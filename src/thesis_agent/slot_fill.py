from __future__ import annotations

import copy
import json
import re
import zipfile
from dataclasses import asdict, dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET

from .docx_inspect import NS, W_NS, _paragraph_text
from .ooxml import serialize_package_xml, serialize_xml
from .rebuild import (
    CT_NS,
    REL_NS,
    R_NS,
    _compact,
    _empty_content_types_xml,
    _empty_rels_xml,
    _extract_source_content,
    _has_visible_content,
    _import_source_relationships,
    _resolve_template_docx,
    _sanitize_imported_body_element,
    _toc_titles_from_body,
    _w,
)


@dataclass(frozen=True)
class SlotFillReport:
    template: Path
    source: Path
    output: Path
    title: str | None = None
    student_id: str | None = None
    student_name: str | None = None
    cover_fields_filled: int = 0
    abstract_paragraphs: int = 0
    english_abstract_paragraphs: int = 0
    toc_entries: int = 0
    body_elements: int = 0
    reference_items: int = 0
    acknowledgement_paragraphs: int = 0
    imported_relationships: int = 0
    imported_parts: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    def to_json(self) -> str:
        return json.dumps(asdict(self), ensure_ascii=False, indent=2, default=str)


@dataclass(frozen=True)
class TemplateSlots:
    front: list[ET.Element]
    zh_abstract_heading: ET.Element
    zh_abstract_body: ET.Element
    zh_keywords: ET.Element
    en_abstract_heading: ET.Element
    en_abstract_body: ET.Element
    en_keywords: ET.Element
    toc_heading: ET.Element
    toc_entry_level1: ET.Element
    toc_entry_level2: ET.Element
    toc_section_break: ET.Element | None
    first_main_heading: ET.Element
    later_main_heading: ET.Element
    second_heading: ET.Element
    third_heading: ET.Element
    body_paragraph: ET.Element
    figure_caption: ET.Element
    table_caption: ET.Element
    reference_heading: ET.Element
    reference_item: ET.Element
    acknowledgement_heading: ET.Element
    acknowledgement_body: ET.Element
    final_section: ET.Element | None


@dataclass(frozen=True)
class AbstractParts:
    zh_paragraphs: list[str]
    zh_keywords: str | None
    en_paragraphs: list[str]
    en_keywords: str | None


def fill_standard_template_docx(template_path: Path, source_docx: Path, output_path: Path) -> SlotFillReport:
    """Fill a formal standard thesis template with extracted student content.

    The template owns all page geometry, front-matter declarations, section
    breaks, paragraph properties and run styles. The source document contributes
    text, tables and images only.
    """
    if source_docx.suffix.lower() != ".docx":
        raise ValueError("slot filling currently supports .docx sources only")
    template_docx = _resolve_template_docx(template_path)
    if template_docx is None:
        raise ValueError(f"No usable .docx template found for {template_path}")
    if not zipfile.is_zipfile(source_docx):
        raise ValueError(f"Not a valid docx file: {source_docx}")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    warnings: list[str] = []
    with zipfile.ZipFile(template_docx) as template_zip, zipfile.ZipFile(source_docx) as source_zip:
        template_root = ET.fromstring(template_zip.read("word/document.xml"))
        template_body = template_root.find("w:body", NS)
        if template_body is None:
            raise ValueError(f"Template has no w:body: {template_docx}")
        source_root = ET.fromstring(source_zip.read("word/document.xml"))
        source_body = source_root.find("w:body", NS)
        if source_body is None:
            raise ValueError(f"Source has no w:body: {source_docx}")

        slots = _extract_template_slots(template_body)
        extracted = _extract_source_content(source_body)
        warnings.extend(extracted.warnings)
        metadata = _extract_metadata(extracted.cover, list(source_body), source_docx)
        if extracted.title:
            metadata.setdefault("title", extracted.title)
        abstract = _parse_abstract_parts(extracted.abstract)
        if not abstract.zh_paragraphs:
            warnings.append("未识别到中文摘要正文，保留模板摘要样式并留空。")
        if not abstract.en_paragraphs:
            warnings.append("未识别到英文摘要正文，保留模板英文摘要样式并留空。")

        target_rels_xml = (
            template_zip.read("word/_rels/document.xml.rels")
            if "word/_rels/document.xml.rels" in template_zip.namelist()
            else _empty_rels_xml()
        )
        target_rels = ET.fromstring(target_rels_xml)
        content_types_xml = (
            template_zip.read("[Content_Types].xml")
            if "[Content_Types].xml" in template_zip.namelist()
            else _empty_content_types_xml()
        )
        content_types = ET.fromstring(content_types_xml)
        imported = _import_source_relationships(
            elements=[*extracted.body, *extracted.references, *extracted.acknowledgements],
            source_zip=source_zip,
            target_rels=target_rels,
            content_types=content_types,
            existing_names=set(template_zip.namelist()),
        )

        toc_titles = _toc_titles_from_body(extracted.body, extracted.references, extracted.acknowledgements)
        new_elements: list[ET.Element] = []
        front, filled_fields = _fill_cover_fields(slots.front, metadata)
        new_elements.extend(front)
        new_elements.extend(_filled_abstract_elements(slots, abstract))
        new_elements.extend(_filled_toc_elements(slots, toc_titles))
        body_elements = _filled_body_elements(slots, extracted.body)
        new_elements.extend(body_elements)
        reference_elements = _filled_reference_elements(slots, extracted.references)
        new_elements.extend(reference_elements)
        acknowledgement_elements = _filled_acknowledgement_elements(slots, extracted.acknowledgements)
        new_elements.extend(acknowledgement_elements)
        if slots.final_section is not None:
            new_elements.append(copy.deepcopy(slots.final_section))

        for child in list(template_body):
            template_body.remove(child)
        for element in new_elements:
            template_body.append(element)

        document_xml = serialize_xml(template_root)
        rels_xml = serialize_package_xml(target_rels, REL_NS)
        content_types_data = serialize_package_xml(content_types, CT_NS)
        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as out_zip:
            written: set[str] = set()
            for item in template_zip.infolist():
                if item.filename == "word/document.xml":
                    data = document_xml
                elif item.filename == "word/_rels/document.xml.rels":
                    data = rels_xml
                elif item.filename == "[Content_Types].xml":
                    data = content_types_data
                else:
                    data = template_zip.read(item.filename)
                out_zip.writestr(item, data)
                written.add(item.filename)
            if "word/_rels/document.xml.rels" not in written:
                out_zip.writestr("word/_rels/document.xml.rels", rels_xml)
                written.add("word/_rels/document.xml.rels")
            if "[Content_Types].xml" not in written:
                out_zip.writestr("[Content_Types].xml", content_types_data)
                written.add("[Content_Types].xml")
            for name, data in imported.parts.items():
                if name not in written:
                    out_zip.writestr(name, data)
                    written.add(name)

    return SlotFillReport(
        template=template_docx,
        source=source_docx,
        output=output_path,
        title=metadata.get("title"),
        student_id=metadata.get("student_id"),
        student_name=metadata.get("student_name"),
        cover_fields_filled=filled_fields,
        abstract_paragraphs=len(abstract.zh_paragraphs),
        english_abstract_paragraphs=len(abstract.en_paragraphs),
        toc_entries=len(toc_titles),
        body_elements=len(body_elements),
        reference_items=max(0, len(reference_elements) - 1),
        acknowledgement_paragraphs=max(0, len(acknowledgement_elements) - 1),
        imported_relationships=imported.relationships,
        imported_parts=sorted(imported.parts),
        warnings=warnings,
    )


def final_output_filename(report: SlotFillReport, suffix: str = ".docx") -> str:
    student_id = _filename_part(report.student_id)
    student_name = _filename_part(report.student_name)
    title = _filename_part(report.title)
    return f"{student_id}-{student_name}-{title}{suffix}"


def _filename_part(value: str | None) -> str:
    if not value:
        return "XX"
    cleaned = re.sub(r"\s+", "", value)
    cleaned = re.sub(r'[\\/:*?"<>|]+', "", cleaned)
    cleaned = cleaned.strip(". ")
    return cleaned or "XX"


def _extract_template_slots(template_body: ET.Element) -> TemplateSlots:
    children = list(template_body)
    paragraphs = [(idx, _paragraph_text(el).strip()) for idx, el in enumerate(children) if el.tag == _w("p")]
    abstract_idx = _required_index(paragraphs, lambda text: _compact(text) == "摘要", "template 摘要 heading")
    zh_keywords_idx = _required_index(paragraphs, lambda text: text.startswith("关键词"), "template 中文关键词")
    en_abstract_idx = _required_index(paragraphs, lambda text: _compact(text).upper() == "ABSTRACT", "template ABSTRACT heading")
    en_keywords_idx = _required_index(paragraphs, lambda text: re.match(r"^Key\s*words?[:：]", text, re.I) is not None, "template English keywords")
    toc_idx = _required_index(paragraphs, lambda text: _compact(text) in {"目录", "目錄"}, "template TOC heading")
    first_body_idx = _required_index(paragraphs, lambda text: re.fullmatch(r"1\s*绪论", text) is not None, "template first body heading")
    reference_idx = _required_index(paragraphs, lambda text: _compact(text) == "参考文献" and idx_gt(paragraphs, text, first_body_idx), "template references")
    acknowledgement_idx = _required_index(paragraphs, lambda text: _compact(text) in {"致谢", "致謝"} and idx_gt(paragraphs, text, reference_idx), "template acknowledgement")

    toc_entry_level1 = _require_paragraph(_first_paragraph_between(children, toc_idx + 1, first_body_idx, lambda text: re.match(r"^[1-9]\s+", text) is not None), "template TOC level 1")
    toc_entry_level2 = _coalesce_element(
        _first_paragraph_between(children, toc_idx + 1, first_body_idx, lambda text: re.match(r"^[1-9]\.\d+", text) is not None),
        toc_entry_level1,
    )
    second_heading = _require_paragraph(_first_paragraph_between(children, first_body_idx + 1, reference_idx, lambda text: re.match(r"^[1-9]\.\d+\s*", text) is not None), "template second heading")
    third_heading = _coalesce_element(
        _first_paragraph_between(children, first_body_idx + 1, reference_idx, lambda text: re.match(r"^[1-9]\.\d+\.\d+\s*", text) is not None),
        second_heading,
    )
    later_main = _coalesce_element(
        _first_paragraph_between(children, first_body_idx + 1, reference_idx, lambda text: re.match(r"^[2-9]\s+", text) is not None),
        children[first_body_idx],
    )
    body_para = _first_paragraph_between(
        children,
        first_body_idx + 1,
        reference_idx,
        lambda text: bool(text) and _body_text_candidate(text),
    )
    body_para = _require_paragraph(body_para, "template body paragraph")
    figure_caption = _coalesce_element(_first_paragraph_between(children, first_body_idx + 1, reference_idx, lambda text: re.match(r"^图\s*\d", text) is not None), body_para)
    table_caption = _coalesce_element(_first_paragraph_between(children, first_body_idx + 1, reference_idx, lambda text: re.match(r"^表\s*\d", text) is not None), figure_caption)
    reference_item = _require_paragraph(_first_paragraph_between(children, reference_idx + 1, acknowledgement_idx, lambda text: bool(text and not _compact(text) == "致谢")), "template reference item")
    acknowledgement_body = _coalesce_element(_first_paragraph_between(children, acknowledgement_idx + 1, len(children), lambda text: bool(text)), body_para)
    toc_section_break = _last_section_break_between(children, toc_idx + 1, first_body_idx)
    final_section = children[-1] if children and children[-1].tag == _w("sectPr") else None

    return TemplateSlots(
        front=[copy.deepcopy(el) for el in children[:abstract_idx]],
        zh_abstract_heading=children[abstract_idx],
        zh_abstract_body=_first_visible_paragraph(children, abstract_idx + 1, zh_keywords_idx),
        zh_keywords=children[zh_keywords_idx],
        en_abstract_heading=children[en_abstract_idx],
        en_abstract_body=_first_visible_paragraph(children, en_abstract_idx + 1, en_keywords_idx),
        en_keywords=children[en_keywords_idx],
        toc_heading=children[toc_idx],
        toc_entry_level1=toc_entry_level1,
        toc_entry_level2=toc_entry_level2,
        toc_section_break=toc_section_break,
        first_main_heading=children[first_body_idx],
        later_main_heading=later_main,
        second_heading=second_heading,
        third_heading=third_heading,
        body_paragraph=body_para,
        figure_caption=figure_caption,
        table_caption=table_caption,
        reference_heading=children[reference_idx],
        reference_item=reference_item,
        acknowledgement_heading=children[acknowledgement_idx],
        acknowledgement_body=acknowledgement_body,
        final_section=final_section,
    )


def idx_gt(paragraphs: list[tuple[int, str]], text: str, index: int) -> bool:
    return any(idx > index and value == text for idx, value in paragraphs)


def _required_index(paragraphs: list[tuple[int, str]], predicate, label: str) -> int:
    for idx, text in paragraphs:
        if predicate(text):
            return idx
    raise ValueError(f"Could not find {label} in standard template")


def _require_paragraph(element: ET.Element | None, label: str) -> ET.Element:
    if element is None:
        raise ValueError(f"Could not find {label} in standard template")
    return element


def _coalesce_element(element: ET.Element | None, fallback: ET.Element) -> ET.Element:
    return element if element is not None else fallback


def _first_visible_paragraph(children: list[ET.Element], start: int, end: int) -> ET.Element:
    paragraph = _first_paragraph_between(children, start, end, lambda text: bool(text))
    if paragraph is None:
        raise ValueError("Template slot has no visible paragraph exemplar")
    return paragraph


def _first_paragraph_between(children: list[ET.Element], start: int, end: int, predicate) -> ET.Element | None:
    for element in children[start:end]:
        if element.tag != _w("p"):
            continue
        text = _paragraph_text(element).strip()
        if predicate(text):
            return element
    return None


def _last_section_break_between(children: list[ET.Element], start: int, end: int) -> ET.Element | None:
    for element in reversed(children[start:end]):
        if element.find(".//w:sectPr", NS) is not None:
            return copy.deepcopy(element)
    return None


def _body_text_candidate(text: str) -> bool:
    compact = _compact(text)
    if not compact:
        return False
    if _is_heading_text(text) or _is_caption_text(text):
        return False
    if compact in {"正文内容"}:
        return True
    return len(compact) > 30


def _extract_metadata(cover: list[ET.Element], all_children: list[ET.Element], source_docx: Path) -> dict[str, str]:
    text = "\n".join(_paragraph_text(el) for el in (cover or all_children[:50]) if el.tag == _w("p"))
    full_text = "\n".join(_paragraph_text(el) for el in all_children if el.tag == _w("p"))
    metadata: dict[str, str] = {}
    labels = {
        "学院": "college",
        "学生姓名": "student_name",
        "姓名": "student_name",
        "学生学号": "student_id",
        "学号": "student_id",
        "专业": "major",
        "指导教师": "advisor",
    }
    for line in [part.strip() for part in text.splitlines() if part.strip()]:
        cleaned = re.sub(r"[_＿]{2,}", "", line).strip()
        for label, key in labels.items():
            match = re.search(rf"{label}\s*[:：]\s*(.+)$", cleaned)
            if match and match.group(1).strip():
                metadata[key] = match.group(1).strip()
    title = _infer_title_from_lines(text.splitlines()) if cover else None
    if title:
        metadata["title"] = title
    else:
        fallback_title = _fallback_title_from_document(full_text, source_docx)
        if fallback_title:
            metadata["title"] = fallback_title
    filename_parts = re.split(r"[-_]", source_docx.stem)
    for part in filename_parts:
        if re.fullmatch(r"[\u4e00-\u9fff]{2,4}", part) and "论文" not in part:
            metadata.setdefault("student_name", part)
    if "物联网" in source_docx.stem:
        metadata.setdefault("major", "物联网工程")
    return metadata


def _infer_title_from_lines(lines: list[str]) -> str | None:
    candidates = [line.strip() for line in lines if line.strip()]
    for idx, line in enumerate(candidates):
        if _compact(line) in {"学士学位论文", "毕业设计（论文）", "本科毕业设计（论文）"}:
            for candidate in candidates[idx + 1: idx + 5]:
                if _looks_like_title(candidate):
                    return candidate
    for candidate in candidates[:20]:
        if _looks_like_title(candidate):
            return candidate
    return None


def _looks_like_title(text: str) -> bool:
    compact = _compact(text)
    if len(compact) < 8 or len(compact) > 45:
        return False
    if re.match(r"^[1-9](?:\.\d+)*", compact):
        return False
    banned = [
        "上海电机学院",
        "学生姓名",
        "学生学号",
        "指导教师",
        "学士学位论文",
        "毕业论文",
        "初稿",
        "摘要",
        "ABSTRACT",
        "关键词",
        "关键字",
        "Keywords",
        "KeyWords",
        "目录",
    ]
    return bool(re.search(r"[\u4e00-\u9fff]", compact)) and not any(item in compact for item in banned)


def _fallback_title_from_document(text: str, source_docx: Path) -> str | None:
    compact = _compact(text)
    candidates = [
        ("危化品" in compact and "监管" in compact, "智能危化品监管系统"),
        ("冷链" in compact and "温控" in compact, "冷链物流温控追踪系统"),
        ("水产" in compact and "监测系统" in compact, "智慧渔业水产养殖监测系统设计"),
        ("停车" in compact and "空气" in compact and ("检测" in compact or "监测" in compact), "地下停车场空气检测系统"),
        ("停车" in compact and "管理系统" in compact, "停车场管理系统"),
    ]
    for matched, title in candidates:
        if matched:
            return title
    stem = re.sub(r"(毕业论文|论文|初稿|终稿|修改稿)", "", source_docx.stem)
    parts = [part.strip() for part in re.split(r"[-_]", stem) if part.strip()]
    for part in reversed(parts):
        if _looks_like_title(part):
            return part
    return None


def _fill_cover_fields(front: list[ET.Element], metadata: dict[str, str]) -> tuple[list[ET.Element], int]:
    result = [copy.deepcopy(el) for el in front]
    filled = 0
    for paragraph in result:
        if paragraph.tag != _w("p"):
            continue
        text = _paragraph_text(paragraph).strip()
        compact = _compact(text)
        if compact == "论文题目" and metadata.get("title"):
            _set_paragraph_text(paragraph, metadata["title"])
            filled += 1
            continue
        field_key = _cover_field_key(text)
        if field_key and metadata.get(field_key):
            if _replace_cover_underline_value(paragraph, metadata[field_key]):
                filled += 1
                continue
            prefix = text.split("：", 1)[0] + "：" if "：" in text else text.split(":", 1)[0] + ":"
            _set_paragraph_text(paragraph, f"{prefix}{metadata[field_key]}")
            filled += 1
    return result, filled


def _replace_cover_underline_value(paragraph: ET.Element, value: str) -> bool:
    if not value:
        return False
    for run in paragraph.findall("w:r", NS):
        for text_node in run.findall("w:t", NS):
            text = text_node.text or ""
            match = re.search(r"_{4,}", text)
            if not match:
                continue
            residual = max(4, match.end() - match.start() - _display_width_units(value))
            replacement = value + "_" * residual
            text_node.text = text[: match.start()] + replacement + text[match.end() :]
            if re.search(r"^\s|\s$", text_node.text or "") or re.search(r"\s{2,}", text_node.text or ""):
                text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            _set_run_underline(run)
            return True
    return False


def _display_width_units(text: str) -> int:
    width = 0
    for char in text:
        width += 2 if "\u4e00" <= char <= "\u9fff" else 1
    return width


def _set_run_underline(run: ET.Element) -> None:
    rpr = run.find("w:rPr", NS)
    if rpr is None:
        rpr = ET.Element(_w("rPr"))
        run.insert(0, rpr)
    underline = rpr.find("w:u", NS)
    if underline is None:
        underline = ET.Element(_w("u"))
        rpr.append(underline)
    underline.set(_w("val"), "single")


def _cover_field_key(text: str) -> str | None:
    compact = _compact(text)
    if compact.startswith("学院：") or compact.startswith("学院:"):
        return "college"
    if compact.startswith("学生姓名：") or compact.startswith("学生姓名:"):
        return "student_name"
    if compact.startswith("学生学号：") or compact.startswith("学生学号:"):
        return "student_id"
    if compact.startswith("专业：") or compact.startswith("专业:"):
        return "major"
    if compact.startswith("指导教师：") or compact.startswith("指导教师:"):
        return "advisor"
    return None


def _parse_abstract_parts(elements: list[ET.Element]) -> AbstractParts:
    section: str | None = None
    zh: list[str] = []
    en: list[str] = []
    zh_keywords: str | None = None
    en_keywords: str | None = None
    for element in elements:
        if element.tag != _w("p"):
            continue
        text = _paragraph_text(element).strip()
        compact = _compact(text)
        if not text:
            continue
        if compact == "摘要":
            section = "zh"
            continue
        if compact.upper() == "ABSTRACT":
            section = "en"
            continue
        if text.startswith("关键词"):
            zh_keywords = _normalize_keywords(text, english=False)
            section = None
            continue
        if re.match(r"^Key\s*words?[:：]", text, re.I):
            en_keywords = _normalize_keywords(text, english=True)
            section = None
            continue
        if section == "zh":
            zh.append(text)
        elif section == "en":
            en.append(text)
    return AbstractParts(zh_paragraphs=zh, zh_keywords=zh_keywords, en_paragraphs=en, en_keywords=en_keywords)


def _normalize_keywords(text: str, english: bool) -> str:
    if english:
        match = re.match(r"^Key\s*words?[:：]\s*(.+)$", text, re.I)
        body = match.group(1).strip() if match else text.strip()
        parts = [part.strip(" ，,.;；。") for part in re.split(r"[,，;；]", body) if part.strip(" ，,.;；。")]
        return "Key words: " + ", ".join(parts)
    match = re.match(r"^关键词[:：]\s*(.+)$", text)
    body = match.group(1).strip() if match else text.strip()
    parts = [part.strip(" ，,.;；。") for part in re.split(r"[,，;；]", body) if part.strip(" ，,.;；。")]
    return "关键词：" + "，".join(parts)


def _filled_abstract_elements(slots: TemplateSlots, abstract: AbstractParts) -> list[ET.Element]:
    elements = [_paragraph_from_template(slots.zh_abstract_heading, "摘  要")]
    zh_paragraphs = abstract.zh_paragraphs or [""]
    elements.extend(_paragraph_from_template(slots.zh_abstract_body, text) for text in zh_paragraphs if text or len(zh_paragraphs) == 1)
    elements.append(_paragraph_from_template(slots.zh_keywords, abstract.zh_keywords or "关键词："))
    elements.append(_paragraph_from_template(slots.en_abstract_heading, "ABSTRACT"))
    en_paragraphs = abstract.en_paragraphs or [""]
    elements.extend(_paragraph_from_template(slots.en_abstract_body, text) for text in en_paragraphs if text or len(en_paragraphs) == 1)
    elements.append(_paragraph_from_template(slots.en_keywords, abstract.en_keywords or "Key words: "))
    return elements


def _filled_toc_elements(slots: TemplateSlots, titles: list[str]) -> list[ET.Element]:
    elements = [_paragraph_from_template(slots.toc_heading, "目  录")]
    for title in titles:
        title = _normalize_heading_spacing(title)
        template = slots.toc_entry_level2 if _toc_level(title) == 2 else slots.toc_entry_level1
        elements.append(_toc_paragraph_from_template(template, title, "1"))
    if slots.toc_section_break is not None:
        elements.append(copy.deepcopy(slots.toc_section_break))
    return elements


def _filled_body_elements(slots: TemplateSlots, body: list[ET.Element]) -> list[ET.Element]:
    elements: list[ET.Element] = []
    seen_first_main = False
    for source in body:
        if source.tag == _w("tbl"):
            elements.append(_sanitize_imported_body_element(source))
            continue
        if source.tag != _w("p"):
            continue
        text = _paragraph_text(source).strip()
        has_embedded_visual = _has_embedded_visual(source)
        if not text and has_embedded_visual:
            elements.append(_visual_only_paragraph_from_source(source))
            continue
        if not text:
            continue
        if _is_main_heading_text(text):
            template = slots.first_main_heading if not seen_first_main else slots.later_main_heading
            elements.append(_paragraph_from_template(template, _normalize_heading_spacing(text)))
            seen_first_main = True
        elif re.match(r"^[1-9]\.\d+\.\d+\s*", text):
            elements.append(_paragraph_from_template(slots.third_heading, _normalize_heading_spacing(text)))
        elif re.match(r"^[1-9]\.\d+\s*", text):
            elements.append(_paragraph_from_template(slots.second_heading, _normalize_heading_spacing(text)))
        elif _is_caption_text(text):
            template = slots.table_caption if text.strip().startswith("表") else slots.figure_caption
            elements.append(_paragraph_from_template(template, _normalize_caption_spacing(text)))
        else:
            elements.append(_paragraph_from_template(slots.body_paragraph, text))
        if has_embedded_visual and not _is_caption_text(text):
            elements.append(_visual_only_paragraph_from_source(source))
    return elements


def _filled_reference_elements(slots: TemplateSlots, references: list[ET.Element]) -> list[ET.Element]:
    items = [
        _paragraph_text(element).strip()
        for element in references
        if element.tag == _w("p") and _paragraph_text(element).strip() and _compact(_paragraph_text(element)) != "参考文献"
    ]
    if not items:
        return []
    return [_paragraph_from_template(slots.reference_heading, "参考文献")] + [
        _paragraph_from_template(slots.reference_item, item) for item in items
    ]


def _filled_acknowledgement_elements(slots: TemplateSlots, acknowledgements: list[ET.Element]) -> list[ET.Element]:
    items = [
        _paragraph_text(element).strip()
        for element in acknowledgements
        if element.tag == _w("p") and _paragraph_text(element).strip() and _compact(_paragraph_text(element)) not in {"致谢", "致謝"}
    ]
    if not items:
        return []
    return [_paragraph_from_template(slots.acknowledgement_heading, "致  谢")] + [
        _paragraph_from_template(slots.acknowledgement_body, item) for item in items
    ]


def _paragraph_from_template(template: ET.Element, text: str) -> ET.Element:
    paragraph = ET.Element(_w("p"))
    for name, value in template.attrib.items():
        if "rsid" not in name and not name.endswith("paraId") and not name.endswith("textId"):
            paragraph.set(name, value)
    ppr = template.find("w:pPr", NS)
    if ppr is not None:
        paragraph.append(copy.deepcopy(ppr))
    if template.find(".//w:br[@w:type='page']", NS) is not None:
        run = ET.SubElement(paragraph, _w("r"))
        br = ET.SubElement(run, _w("br"))
        br.set(_w("type"), "page")
    run = ET.SubElement(paragraph, _w("r"))
    rpr = _first_run_properties(template)
    if rpr is not None:
        run.append(copy.deepcopy(rpr))
    t = ET.SubElement(run, _w("t"))
    t.text = text
    if re.search(r"^\s|\s$", text) or re.search(r"\s{2,}", text):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return paragraph


def _set_paragraph_text(paragraph: ET.Element, text: str) -> None:
    ppr = paragraph.find("w:pPr", NS)
    preserved_ppr = copy.deepcopy(ppr) if ppr is not None else None
    rpr = _first_run_properties(paragraph)
    for child in list(paragraph):
        paragraph.remove(child)
    if preserved_ppr is not None:
        paragraph.append(preserved_ppr)
    run = ET.SubElement(paragraph, _w("r"))
    if rpr is not None:
        run.append(copy.deepcopy(rpr))
    t = ET.SubElement(run, _w("t"))
    t.text = text
    if re.search(r"^\s|\s$", text) or re.search(r"\s{2,}", text):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def _toc_paragraph_from_template(template: ET.Element, title: str, label: str) -> ET.Element:
    paragraph = _paragraph_from_template(template, "")
    _set_toc_entry_layout(paragraph, _toc_level(title))
    for run in list(paragraph.findall("w:r", NS)):
        paragraph.remove(run)
    run = ET.SubElement(paragraph, _w("r"))
    rpr = _first_run_properties(template)
    if rpr is not None:
        run.append(copy.deepcopy(rpr))
    t = ET.SubElement(run, _w("t"))
    t.text = title
    if re.search(r"\s", title):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    ET.SubElement(run, _w("tab"))
    page = ET.SubElement(run, _w("t"))
    page.text = label
    return paragraph


def _visual_only_paragraph_from_source(source: ET.Element) -> ET.Element:
    paragraph = _sanitize_imported_body_element(source)
    for run in list(paragraph.findall("w:r", NS)):
        if run.find(".//w:drawing", NS) is None and run.find(".//w:pict", NS) is None and run.find(".//w:object", NS) is None:
            paragraph.remove(run)
            continue
        for text_node in run.findall(".//w:t", NS):
            text_node.text = ""
    ppr = paragraph.find("w:pPr", NS)
    if ppr is None:
        ppr = ET.Element(_w("pPr"))
        paragraph.insert(0, ppr)
    for name in ("ind", "numPr", "pStyle", "pageBreakBefore"):
        existing = ppr.find(f"w:{name}", NS)
        if existing is not None:
            ppr.remove(existing)
    spacing = ppr.find("w:spacing", NS)
    if spacing is None:
        spacing = ET.Element(_w("spacing"))
        ppr.append(spacing)
    spacing.set(_w("before"), "0")
    spacing.set(_w("after"), "0")
    spacing.set(_w("line"), "240")
    spacing.set(_w("lineRule"), "auto")
    jc = ppr.find("w:jc", NS)
    if jc is None:
        jc = ET.Element(_w("jc"))
        ppr.append(jc)
    jc.set(_w("val"), "center")
    return paragraph


def _has_embedded_visual(element: ET.Element) -> bool:
    return (
        element.find(".//w:drawing", NS) is not None
        or element.find(".//w:pict", NS) is not None
        or element.find(".//w:object", NS) is not None
    )


def _set_toc_entry_layout(paragraph: ET.Element, level: int) -> None:
    ppr = paragraph.find("w:pPr", NS)
    if ppr is None:
        ppr = ET.Element(_w("pPr"))
        paragraph.insert(0, ppr)
    ind = ppr.find("w:ind", NS)
    if ind is None:
        ind = ET.Element(_w("ind"))
        ppr.append(ind)
    ind.set(_w("left"), "420" if level == 2 else "0")
    for name in ("firstLine", "firstLineChars", "hanging", "hangingChars"):
        ind.attrib.pop(_w(name), None)
    tabs = ppr.find("w:tabs", NS)
    if tabs is None:
        tabs = ET.Element(_w("tabs"))
        ppr.append(tabs)
    for tab in list(tabs):
        tabs.remove(tab)
    tab = ET.SubElement(tabs, _w("tab"))
    tab.set(_w("val"), "right")
    tab.set(_w("leader"), "dot")
    tab.set(_w("pos"), "9070")


def _first_run_properties(paragraph: ET.Element) -> ET.Element | None:
    for run in paragraph.findall("w:r", NS):
        text = "".join(t.text or "" for t in run.findall("w:t", NS))
        if text.strip():
            rpr = run.find("w:rPr", NS)
            return copy.deepcopy(rpr) if rpr is not None else None
    rpr = paragraph.find("w:pPr/w:rPr", NS)
    return copy.deepcopy(rpr) if rpr is not None else None


def _is_heading_text(text: str) -> bool:
    return _is_main_heading_text(text) or re.match(r"^[1-9]\.\d+(?:\.\d+)?\s*", text.strip()) is not None


def _is_main_heading_text(text: str) -> bool:
    return re.match(r"^[1-9]\s*[\u4e00-\u9fffA-Za-z].{0,60}$", text.strip()) is not None


def _is_caption_text(text: str) -> bool:
    return re.match(r"^(图|表)\s*\d+\s*[-－]\s*\d+", text.strip()) is not None


def _normalize_heading_spacing(text: str) -> str:
    stripped = text.strip()
    match = re.match(r"^([1-9])\s*(.+)$", stripped)
    if match and not stripped.startswith(match.group(1) + "."):
        return f"{match.group(1)} {match.group(2).strip()}"
    match = re.match(r"^([1-9](?:\.\d+)+)\s*(.+)$", stripped)
    if match:
        return f"{match.group(1)} {match.group(2).strip()}"
    return stripped


def _normalize_caption_spacing(text: str) -> str:
    return re.sub(r"^(图|表)\s*([0-9]+)\s*[-－]\s*([0-9]+)\s*", r"\1 \2-\3 ", text.strip())


def _toc_level(title: str) -> int:
    return 2 if re.match(r"^[1-9]\.\d+", title.strip()) else 1
