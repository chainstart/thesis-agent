from __future__ import annotations

import copy
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET

from .docx_inspect import NS, W_NS, _paragraph_text
from .ooxml import serialize_package_xml, serialize_xml
from .reference_style import normalize_reference_text


R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)
ET.register_namespace("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")
ET.register_namespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
ET.register_namespace("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture")


@dataclass(frozen=True)
class FormatFixReport:
    input: Path
    output: Path
    removed_trailing_empty_paragraphs: int = 0
    removed_trailing_section_paragraphs: int = 0
    removed_empty_paragraphs_in_runs: int = 0
    removed_empty_paragraphs_after_sections: int = 0
    removed_empty_paragraphs_before_page_breaks: int = 0
    removed_orphan_empty_body_paragraphs: int = 0
    removed_layout_paragraphs_before_page_break_headings: int = 0
    merged_empty_section_paragraphs: int = 0
    removed_existing_front_matter_paragraphs: int = 0
    inserted_front_matter_paragraphs: int = 0
    front_matter_lines_normalized: int = 0
    abstract_paragraphs_normalized: int = 0
    toc_headings_fixed: int = 0
    toc_page_breaks_inserted: int = 0
    toc_section_breaks_moved: int = 0
    page_number_restart_applied: bool = False
    front_page_numbers_normalized: int = 0
    front_matter_page_breaks_inserted: int = 0
    headers_normalized: int = 0
    body_paragraphs_normalized: int = 0
    tables_normalized: int = 0
    acknowledgement_paragraphs_normalized: int = 0
    heading_styles_applied: int = 0
    sub_heading_styles_applied: int = 0
    reference_paragraphs_normalized: int = 0
    reference_empty_paragraphs_removed: int = 0
    reference_field_codes_removed: int = 0
    caption_paragraphs_normalized: int = 0
    caption_kinds_fixed: int = 0
    caption_keep_rules_applied: int = 0
    warnings: list[str] = field(default_factory=list)


@dataclass(frozen=True)
class TableTemplateFormat:
    tbl_pr: ET.Element | None
    body_font_size: str
    grid_width: int | None


def fix_docx_format(
    input_path: Path,
    output_path: Path,
    max_empty_run: int = 1,
    template_path: Path | None = None,
    preserve_template_front_matter: bool = False,
) -> FormatFixReport:
    if input_path.suffix.lower() != ".docx":
        raise ValueError("fix-format currently supports .docx targets only")
    if not zipfile.is_zipfile(input_path):
        raise ValueError(f"Not a valid docx file: {input_path}")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    warnings: list[str] = []
    with zipfile.ZipFile(input_path) as zin:
        document_xml = zin.read("word/document.xml")
        styles_xml = zin.read("word/styles.xml") if "word/styles.xml" in zin.namelist() else None
        rels_xml = zin.read("word/_rels/document.xml.rels") if "word/_rels/document.xml.rels" in zin.namelist() else _empty_rels_xml()
        content_types_xml = zin.read("[Content_Types].xml") if "[Content_Types].xml" in zin.namelist() else _empty_content_types_xml()
        root = ET.fromstring(document_xml)
        heading1_style = _find_heading_style_id(styles_xml, "heading 1") if styles_xml else None

        body = root.find("w:body", NS)
        if body is None:
            raise ValueError("word/document.xml does not contain w:body")

        removed_trailing_empty, removed_trailing_sections = _remove_trailing_empty_paragraphs(body)
        if preserve_template_front_matter:
            merged_empty_sections = 0
            removed_runs = 0
            removed_after_sections = 0
            removed_before_page_breaks = 0
        else:
            merged_empty_sections = _merge_empty_section_paragraphs_into_previous(body)
            removed_runs = _reduce_empty_paragraph_runs(body, max_empty_run=max_empty_run)
            removed_after_sections = _remove_empty_paragraphs_after_section_breaks(body)
            removed_before_page_breaks = _remove_empty_paragraphs_before_page_breaks(body)
        removed_orphan_empty_body = _remove_orphan_empty_body_paragraphs(body)
        inserted_front_matter = 0
        removed_existing_front_matter = 0
        extra_parts: dict[str, bytes] = {}
        resolved_template = _resolve_template_docx(template_path)
        table_template = _table_template_format(resolved_template) if resolved_template else None
        if resolved_template:
            front_result = _front_matter_from_template(
                resolved_template,
                rels_xml=rels_xml,
                content_types_xml=content_types_xml,
                existing_names=set(zin.namelist()),
            )
            if _missing_required_front_matter(body):
                removed_existing_front_matter = _remove_existing_front_matter_before_abstract(body)
                inserted_front_matter = _insert_front_matter(body, front_result.elements)
                rels_xml = front_result.rels_xml
                content_types_xml = front_result.content_types_xml
                extra_parts = front_result.extra_parts
        elif template_path:
            warnings.append(f"Template front matter was not inserted because no .docx template was found for {template_path}.")

        if preserve_template_front_matter:
            front_breaks = 0
            front_lines = _normalize_authorization_title_offset(body)
        else:
            front_breaks = _ensure_front_matter_page_boundaries(body)
            front_lines = _normalize_front_matter_lines(body)
        abstract_paragraphs = _normalize_abstract_and_keywords(body)
        toc_page_breaks_inserted = _normalize_toc_content_controls(body)
        toc_headings_fixed = _fix_toc_headings(body)
        toc_section_breaks_moved = _move_toc_heading_section_breaks_to_previous(body)
        toc_page_breaks_inserted += _ensure_page_break_before_toc_blocks(body)
        restart_applied, front_numbers = _restart_page_number_at_first_body_heading(body)
        body_paragraphs = _normalize_body_paragraph_format(body)
        tables_normalized = _normalize_table_format(body, table_template)
        headings_applied = _apply_main_heading_style(body, heading1_style)
        removed_before_page_break_headings = _remove_layout_paragraphs_before_page_break_headings(body)
        sub_headings_applied = _apply_sub_heading_style(body)
        acknowledgement_paragraphs = _normalize_acknowledgement_body(body)
        reference_paragraphs, reference_empty_removed, reference_fields_removed = _normalize_reference_paragraphs(body)
        caption_kinds = _fix_caption_kinds(body)
        caption_paragraphs = _normalize_caption_format(body)
        captions_applied = _apply_caption_keep_rules(body)

        fixed_document_xml = _serialize_xml(root)
        header_title = _document_title_for_header(body) or _fallback_header_title_from_body(body) or _existing_header_title(zin)
        header_updates: dict[str, bytes] = {}
        if header_title:
            for name in zin.namelist():
                if name.startswith("word/header") and name.endswith(".xml"):
                    header_updates[name] = _normalized_header_xml(header_title)
        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    data = fixed_document_xml
                elif item.filename in header_updates:
                    data = header_updates[item.filename]
                elif item.filename == "word/_rels/document.xml.rels":
                    data = rels_xml
                elif item.filename == "[Content_Types].xml":
                    data = content_types_xml
                else:
                    data = zin.read(item.filename)
                zout.writestr(item, data)
            if "word/_rels/document.xml.rels" not in zin.namelist():
                zout.writestr("word/_rels/document.xml.rels", rels_xml)
            if "[Content_Types].xml" not in zin.namelist():
                zout.writestr("[Content_Types].xml", content_types_xml)
            for name, data in extra_parts.items():
                if name not in zin.namelist():
                    zout.writestr(name, data)

    if not heading1_style:
        warnings.append("Heading 1 style was not found; main heading style normalization was skipped.")
    if not restart_applied:
        warnings.append("Could not locate the section before the first main body heading; page number restart was not applied.")
    return FormatFixReport(
        input=input_path,
        output=output_path,
        removed_trailing_empty_paragraphs=removed_trailing_empty,
        removed_trailing_section_paragraphs=removed_trailing_sections,
        removed_empty_paragraphs_in_runs=removed_runs,
        removed_empty_paragraphs_after_sections=removed_after_sections,
        removed_empty_paragraphs_before_page_breaks=removed_before_page_breaks,
        removed_orphan_empty_body_paragraphs=removed_orphan_empty_body,
        removed_layout_paragraphs_before_page_break_headings=removed_before_page_break_headings,
        merged_empty_section_paragraphs=merged_empty_sections,
        removed_existing_front_matter_paragraphs=removed_existing_front_matter,
        inserted_front_matter_paragraphs=inserted_front_matter,
        front_matter_lines_normalized=front_lines,
        abstract_paragraphs_normalized=abstract_paragraphs,
        toc_headings_fixed=toc_headings_fixed,
        toc_page_breaks_inserted=toc_page_breaks_inserted,
        toc_section_breaks_moved=toc_section_breaks_moved,
        page_number_restart_applied=restart_applied,
        front_page_numbers_normalized=front_numbers,
        front_matter_page_breaks_inserted=front_breaks,
        headers_normalized=len(header_updates),
        body_paragraphs_normalized=body_paragraphs,
        tables_normalized=tables_normalized,
        acknowledgement_paragraphs_normalized=acknowledgement_paragraphs,
        heading_styles_applied=headings_applied,
        sub_heading_styles_applied=sub_headings_applied,
        reference_paragraphs_normalized=reference_paragraphs,
        reference_empty_paragraphs_removed=reference_empty_removed,
        reference_field_codes_removed=reference_fields_removed,
        caption_paragraphs_normalized=caption_paragraphs,
        caption_kinds_fixed=caption_kinds,
        caption_keep_rules_applied=captions_applied,
        warnings=warnings,
    )


@dataclass(frozen=True)
class FrontMatterImport:
    elements: list[ET.Element]
    rels_xml: bytes
    content_types_xml: bytes
    extra_parts: dict[str, bytes]


def _resolve_template_docx(template_path: Path | None) -> Path | None:
    if template_path is None:
        return None
    if template_path.name.startswith("~$"):
        return None
    if template_path.suffix.lower() == ".docx" and template_path.exists():
        return template_path
    candidate = template_path.with_suffix(".docx")
    if not candidate.name.startswith("~$") and candidate.exists():
        return candidate
    return None


def _table_template_format(template_path: Path) -> TableTemplateFormat | None:
    try:
        with zipfile.ZipFile(template_path) as template_zip:
            root = ET.fromstring(template_zip.read("word/document.xml"))
    except (KeyError, zipfile.BadZipFile, ET.ParseError):
        return None
    table = root.find(".//w:tbl", NS)
    if table is None:
        return None
    tbl_pr = table.find("w:tblPr", NS)
    return TableTemplateFormat(
        tbl_pr=copy.deepcopy(tbl_pr) if tbl_pr is not None else None,
        body_font_size=_first_visible_run_size(table) or "24",
        grid_width=_table_grid_width(table),
    )


def _first_visible_run_size(element: ET.Element) -> str | None:
    for run in element.iter(_w("r")):
        text = "".join(node.text or "" for node in run.findall("w:t", NS)).strip()
        if not text:
            continue
        rpr = run.find("w:rPr", NS)
        size = rpr.find("w:sz", NS) if rpr is not None else None
        if size is not None and size.get(_w("val")):
            return size.get(_w("val"))
        size = rpr.find("w:szCs", NS) if rpr is not None else None
        if size is not None and size.get(_w("val")):
            return size.get(_w("val"))
        break
    return None


def _front_matter_from_template(
    template_path: Path,
    rels_xml: bytes,
    content_types_xml: bytes,
    existing_names: set[str],
) -> FrontMatterImport:
    with zipfile.ZipFile(template_path) as template_zip:
        template_root = ET.fromstring(template_zip.read("word/document.xml"))
        template_body = template_root.find("w:body", NS)
        if template_body is None:
            raise ValueError(f"Template has no w:body: {template_path}")
        elements = [copy.deepcopy(el) for el in _extract_front_matter_elements(template_body)]
        _normalize_front_matter_section_references(elements, rels_xml)
        rels_xml, content_types_xml, extra_parts = _import_element_relationships(
            elements=elements,
            template_zip=template_zip,
            target_rels_xml=rels_xml,
            target_content_types_xml=content_types_xml,
            existing_names=existing_names,
        )
    return FrontMatterImport(
        elements=elements,
        rels_xml=rels_xml,
        content_types_xml=content_types_xml,
        extra_parts=extra_parts,
    )


def _extract_front_matter_elements(template_body: ET.Element) -> list[ET.Element]:
    elements: list[ET.Element] = []
    for child in list(template_body):
        text = _paragraph_text(child) if child.tag == _w("p") else ""
        if _is_template_abstract_marker(text):
            break
        elements.append(child)
    elements.append(_page_break_paragraph())
    return elements


def _is_template_abstract_marker(text: str) -> bool:
    compact = re.sub(r"\s+", "", text)
    return "摘要正文" in compact or "小二号黑体摘要" in compact or "摘 要" in text


def _missing_required_front_matter(body: ET.Element) -> bool:
    text = "\n".join(_paragraph_text(p) for p in body.iter(_w("p")))
    compact = re.sub(r"\s+", "", text)
    return not all(marker in compact for marker in ["学术诚信声明", "AI使用情况声明", "版权使用授权书"])


def _insert_front_matter(body: ET.Element, elements: list[ET.Element]) -> int:
    for offset, element in enumerate(elements):
        body.insert(offset, copy.deepcopy(element))
    return sum(1 for element in elements if element.tag == _w("p"))


def _remove_existing_front_matter_before_abstract(body: ET.Element) -> int:
    children = list(body)
    abstract_idx = None
    for idx, child in enumerate(children):
        text = _paragraph_text(child) if child.tag == _w("p") else ""
        if _is_abstract_heading(text):
            abstract_idx = idx
            break
    if abstract_idx is None or abstract_idx == 0:
        return 0
    front_text = "".join(_paragraph_text(child) for child in children[:abstract_idx] if child.tag == _w("p"))
    compact = re.sub(r"\s+", "", front_text)
    markers = ["学术诚信声明", "AI使用情况声明", "版权使用授权书"]
    if not any(marker in compact for marker in markers):
        return 0
    removed = 0
    for child in children[:abstract_idx]:
        if child.tag == _w("p"):
            removed += 1
        body.remove(child)
    return removed


def _is_abstract_heading(text: str) -> bool:
    compact = re.sub(r"\s+", "", text)
    return compact in {"摘要", "Abstract", "ABSTRACT"} or compact == "摘要正文"


def _page_break_paragraph() -> ET.Element:
    p = ET.Element(_w("p"))
    r = ET.SubElement(p, _w("r"))
    br = ET.SubElement(r, _w("br"))
    br.set(_w("type"), "page")
    return p


def _ensure_front_matter_page_boundaries(body: ET.Element) -> int:
    anchors = _front_matter_page_anchor_indices(list(body))
    inserted = 0
    for original_idx in anchors:
        idx = original_idx + inserted
        children = list(body)
        if idx <= 0 or idx >= len(children):
            continue
        anchor = children[idx]
        if _contains_explicit_page_break(anchor) or _contains_section_properties(anchor):
            continue
        if _has_recent_page_boundary_before(children, idx):
            continue
        body.insert(idx, _page_break_paragraph())
        inserted += 1
    return inserted


def _has_recent_page_boundary_before(children: list[ET.Element], idx: int) -> bool:
    for previous in reversed(children[:idx]):
        if _contains_explicit_page_break(previous) or _contains_section_properties(previous):
            return True
        if _has_visible_content(previous):
            return False
    return True


def _front_matter_page_anchor_indices(children: list[ET.Element]) -> list[int]:
    anchors: list[int] = []
    seen: set[int] = set()
    for idx, child in enumerate(children):
        if child.tag != _w("p"):
            continue
        compact = re.sub(r"\s+", "", _paragraph_text(child))
        is_anchor = (
            compact in {"毕业设计（论文）学术诚信声明", "毕业设计（论文）AI使用情况声明", "毕业设计（论文）版权使用授权书", "摘要", "ABSTRACT", "目录", "目錄"}
            or re.fullmatch(r"1\s*绪论", _paragraph_text(child).strip()) is not None
        )
        if not is_anchor:
            continue
        anchor_idx = idx
        if compact.startswith("毕业设计（论文）") and idx > 0:
            previous_text = re.sub(r"\s+", "", _paragraph_text(children[idx - 1]) if children[idx - 1].tag == _w("p") else "")
            if _looks_like_school_heading(previous_text):
                anchor_idx = idx - 1
        if anchor_idx > 0 and anchor_idx not in seen:
            anchors.append(anchor_idx)
            seen.add(anchor_idx)
    return anchors


def _looks_like_school_heading(text: str) -> bool:
    if not text:
        return False
    if len(text) > 20:
        return False
    return bool(re.search(r"[\u4e00-\u9fff](大学|学院)$", text))


def _normalize_front_matter_section_references(elements: list[ET.Element], target_rels_xml: bytes) -> None:
    target_header_rid = _target_header_relationship_id(target_rels_xml)
    for element in elements:
        for sect in element.findall(".//w:sectPr", NS):
            for ref_name in ("headerReference", "footerReference"):
                for ref in list(sect.findall(f"w:{ref_name}", NS)):
                    sect.remove(ref)
            if target_header_rid:
                header = ET.Element(_w("headerReference"))
                header.set(_w("type"), "default")
                header.set(f"{{{R_NS}}}id", target_header_rid)
                sect.insert(0, header)


def _target_header_relationship_id(rels_xml: bytes) -> str | None:
    root = ET.fromstring(rels_xml)
    for rel in root.findall(f"{{{REL_NS}}}Relationship"):
        if rel.attrib.get("Type", "").endswith("/header"):
            return rel.attrib.get("Id")
    return None


def _document_title_for_header(body: ET.Element) -> str | None:
    paragraphs = [p for p in body.findall("w:p", NS)]
    texts = [_paragraph_text(p).strip() for p in paragraphs]
    for idx, text in enumerate(texts):
        compact = re.sub(r"\s+", "", text)
        if compact in {"学士学位论文", "毕业设计（论文）", "本科毕业设计（论文）"}:
            for candidate in texts[idx + 1: idx + 5]:
                cleaned = candidate.strip()
                if _looks_like_header_title(cleaned):
                    return cleaned
    for text in texts[:30]:
        if _looks_like_header_title(text):
            return text.strip()
    return None


def _looks_like_header_title(text: str) -> bool:
    compact = re.sub(r"\s+", "", text)
    if not compact:
        return False
    if len(compact) < 8 or len(compact) > 40:
        return False
    if re.match(r"^[0-9ivxlcdmIVXLCDM]+", compact):
        return False
    if re.match(r"^[0-9一二三四五六七八九十]+[．.、)]", compact):
        return False
    banned = [
        "学校名称",
        "学术诚信声明",
        "使用情况声明",
        "版权使用授权书",
        "保密",
        "不保密",
        "AI工具",
        "本人承诺",
        "未使用",
        "生成或篡改",
        "原始数据",
        "学生姓名",
        "学生学号",
        "指导教师",
        "专业",
        "学院",
        "大学",
        "日期",
    ]
    if any(item in compact for item in banned):
        return False
    return bool(re.search(r"[\u4e00-\u9fff]", compact))


def _fallback_header_title_from_body(body: ET.Element) -> str | None:
    text = "\n".join(_paragraph_text(p) for p in body.iter(_w("p")))
    compact = re.sub(r"\s+", "", text)
    for match in re.finditer(r"([\u4e00-\u9fffA-Za-z0-9]{4,32}(?:系统|平台|装置|软件|应用|模型|算法|方案)(?:设计|实现|研究|开发)?)", compact):
        candidate = match.group(1)
        if _looks_like_header_title(candidate):
            return candidate
    return None


def _existing_header_title(archive: zipfile.ZipFile) -> str | None:
    for name in archive.namelist():
        if not name.startswith("word/header") or not name.endswith(".xml"):
            continue
        try:
            root = ET.fromstring(archive.read(name))
        except ET.ParseError:
            continue
        text = "".join(t.text or "" for t in root.iter(_w("t")))
        cleaned = re.sub(r"\s+", "", text)
        cleaned = re.sub(r"([ivxlcdmIVXLCDM]+|\d+)$", "", cleaned).strip()
        cleaned = re.sub(r"^([ivxlcdmIVXLCDM]+|\d+)", "", cleaned).strip()
        if _looks_like_header_title(cleaned):
            return cleaned
    return None


def _normalized_header_xml(title: str) -> bytes:
    header = ET.Element(_w("hdr"))
    paragraph = ET.SubElement(header, _w("p"))
    ppr = ET.SubElement(paragraph, _w("pPr"))
    spacing = ET.SubElement(ppr, _w("spacing"))
    spacing.set(_w("before"), "0")
    spacing.set(_w("after"), "0")
    border = ET.SubElement(ppr, _w("pBdr"))
    bottom = ET.SubElement(border, _w("bottom"))
    bottom.set(_w("val"), "single")
    bottom.set(_w("sz"), "4")
    bottom.set(_w("space"), "1")
    bottom.set(_w("color"), "auto")
    tabs = ET.SubElement(ppr, _w("tabs"))
    tab = ET.SubElement(tabs, _w("tab"))
    tab.set(_w("val"), "right")
    tab.set(_w("pos"), "9070")
    jc = ET.SubElement(ppr, _w("jc"))
    jc.set(_w("val"), "left")

    title_run = ET.SubElement(paragraph, _w("r"))
    _set_run_font(title_run, east_asia="黑体", ascii_font="Times New Roman", size="18", bold=False)
    title_text = ET.SubElement(title_run, _w("t"))
    title_text.text = title

    tab_run = ET.SubElement(paragraph, _w("r"))
    ET.SubElement(tab_run, _w("tab"))

    begin_run = ET.SubElement(paragraph, _w("r"))
    begin = ET.SubElement(begin_run, _w("fldChar"))
    begin.set(_w("fldCharType"), "begin")

    instr_run = ET.SubElement(paragraph, _w("r"))
    instr = ET.SubElement(instr_run, _w("instrText"))
    instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instr.text = " PAGE "

    separate_run = ET.SubElement(paragraph, _w("r"))
    separate = ET.SubElement(separate_run, _w("fldChar"))
    separate.set(_w("fldCharType"), "separate")

    value_run = ET.SubElement(paragraph, _w("r"))
    _set_run_font(value_run, east_asia="黑体", ascii_font="Times New Roman", size="18", bold=False)
    value = ET.SubElement(value_run, _w("t"))
    value.text = "1"

    end_run = ET.SubElement(paragraph, _w("r"))
    end = ET.SubElement(end_run, _w("fldChar"))
    end.set(_w("fldCharType"), "end")
    return _serialize_xml(header)


def _import_element_relationships(
    elements: list[ET.Element],
    template_zip: zipfile.ZipFile,
    target_rels_xml: bytes,
    target_content_types_xml: bytes,
    existing_names: set[str],
) -> tuple[bytes, bytes, dict[str, bytes]]:
    template_rels_xml = template_zip.read("word/_rels/document.xml.rels") if "word/_rels/document.xml.rels" in template_zip.namelist() else _empty_rels_xml()
    template_rels = ET.fromstring(template_rels_xml)
    target_rels = ET.fromstring(target_rels_xml)
    target_content_types = ET.fromstring(target_content_types_xml)
    template_content_types = ET.fromstring(template_zip.read("[Content_Types].xml")) if "[Content_Types].xml" in template_zip.namelist() else ET.fromstring(_empty_content_types_xml())

    target_rids = {rel.attrib.get("Id", "") for rel in target_rels.findall(f"{{{REL_NS}}}Relationship")}
    used_rids = sorted(_relationship_ids_in_elements(elements) - target_rids)
    if not used_rids:
        return target_rels_xml, target_content_types_xml, {}
    rid_map: dict[str, str] = {}
    extra_parts: dict[str, bytes] = {}

    for old_rid in used_rids:
        template_rel = _find_relationship(template_rels, old_rid)
        if template_rel is None:
            continue
        new_rid = _next_rid(target_rids)
        target_rids.add(new_rid)
        rid_map[old_rid] = new_rid

        new_attrib = dict(template_rel.attrib)
        new_attrib["Id"] = new_rid
        target = new_attrib.get("Target", "")
        mode = new_attrib.get("TargetMode")
        if mode != "External" and target:
            source_part = _word_part_name(target)
            if source_part in template_zip.namelist():
                new_part = _unique_import_part_name(source_part, existing_names | set(extra_parts))
                extra_parts[new_part] = template_zip.read(source_part)
                new_attrib["Target"] = new_part.removeprefix("word/")
                _ensure_content_type_override(
                    target_content_types,
                    part_name="/" + new_part,
                    content_type=_content_type_for_part(template_content_types, source_part, new_attrib.get("Type", "")),
                )

        ET.SubElement(target_rels, f"{{{REL_NS}}}Relationship", new_attrib)

    _remap_relationship_ids(elements, rid_map)
    return _serialize_package_xml(target_rels, REL_NS), _serialize_package_xml(target_content_types, CT_NS), extra_parts


def _relationship_ids_in_elements(elements: list[ET.Element]) -> set[str]:
    ids: set[str] = set()
    for element in elements:
        for node in element.iter():
            for key, value in node.attrib.items():
                if key == f"{{{R_NS}}}id":
                    ids.add(value)
    return ids


def _remap_relationship_ids(elements: list[ET.Element], rid_map: dict[str, str]) -> None:
    if not rid_map:
        return
    for element in elements:
        for node in element.iter():
            for key, value in list(node.attrib.items()):
                if key == f"{{{R_NS}}}id" and value in rid_map:
                    node.set(key, rid_map[value])


def _find_relationship(root: ET.Element, rid: str) -> ET.Element | None:
    for rel in root.findall(f"{{{REL_NS}}}Relationship"):
        if rel.attrib.get("Id") == rid:
            return rel
    return None


def _next_rid(existing: set[str]) -> str:
    max_id = 0
    for rid in existing:
        match = re.match(r"rId(\d+)$", rid)
        if match:
            max_id = max(max_id, int(match.group(1)))
    return f"rId{max_id + 1}"


def _word_part_name(target: str) -> str:
    normalized = target.lstrip("/")
    if normalized.startswith("word/"):
        return normalized
    return f"word/{normalized}"


def _unique_import_part_name(source_part: str, used_names: set[str]) -> str:
    if source_part not in used_names:
        return source_part
    path = Path(source_part)
    stem = path.stem
    suffix = path.suffix
    parent = path.parent.as_posix()
    index = 1
    while True:
        candidate = f"{parent}/{stem}_imported{index}{suffix}"
        if candidate not in used_names:
            return candidate
        index += 1


def _ensure_content_type_override(root: ET.Element, part_name: str, content_type: str) -> None:
    for override in root.findall(f"{{{CT_NS}}}Override"):
        if override.attrib.get("PartName") == part_name:
            override.set("ContentType", content_type)
            return
    ET.SubElement(root, f"{{{CT_NS}}}Override", {"PartName": part_name, "ContentType": content_type})


def _content_type_for_part(root: ET.Element, source_part: str, rel_type: str) -> str:
    part_name = "/" + source_part
    for override in root.findall(f"{{{CT_NS}}}Override"):
        if override.attrib.get("PartName") == part_name:
            return override.attrib.get("ContentType", "application/xml")
    if rel_type.endswith("/header"):
        return "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
    if rel_type.endswith("/footer"):
        return "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
    return "application/xml"


def _empty_rels_xml() -> bytes:
    return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'


def _empty_content_types_xml() -> bytes:
    return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>'


def _remove_trailing_empty_paragraphs(body: ET.Element) -> tuple[int, int]:
    children = list(body)
    last_content_idx = -1
    for idx, child in enumerate(children):
        if _has_visible_content(child):
            last_content_idx = idx

    removed_empty = 0
    removed_sections = 0
    for child in children[last_content_idx + 1:]:
        if child.tag != _w("p"):
            continue
        if child.find(".//w:sectPr", NS) is not None:
            removed_sections += 1
        else:
            removed_empty += 1
        body.remove(child)
    return removed_empty, removed_sections


def _reduce_empty_paragraph_runs(body: ET.Element, max_empty_run: int) -> int:
    removed = 0
    empty_run = 0
    for child in list(body):
        if _is_removable_empty_paragraph(child):
            empty_run += 1
            if empty_run > max_empty_run:
                body.remove(child)
                removed += 1
        elif child.tag == _w("p") and not _has_visible_content(child):
            empty_run = 0
        else:
            empty_run = 0
    return removed


def _merge_empty_section_paragraphs_into_previous(body: ET.Element) -> int:
    merged = 0
    children = list(body)
    idx = 0
    while idx < len(children):
        child = children[idx]
        if (
            child.tag != _w("p")
            or _has_visible_content(child)
            or _contains_explicit_page_break(child)
            or child.find(".//w:sectPr", NS) is None
        ):
            idx += 1
            continue
        previous = _previous_paragraph_with_content(children, idx)
        if previous is None:
            idx += 1
            continue
        source_ppr = child.find("w:pPr", NS)
        sect = source_ppr.find("w:sectPr", NS) if source_ppr is not None else None
        if sect is None:
            idx += 1
            continue
        target_ppr = _ensure_ppr(previous)
        existing = target_ppr.find("w:sectPr", NS)
        if existing is not None:
            idx += 1
            continue
        source_ppr.remove(sect)
        target_ppr.append(sect)
        body.remove(child)
        children.pop(idx)
        merged += 1
    return merged


def _remove_empty_paragraphs_after_section_breaks(body: ET.Element) -> int:
    removed = 0
    children = list(body)
    idx = 0
    while idx + 1 < len(children):
        child = children[idx]
        if child.tag != _w("p") or child.find(".//w:sectPr", NS) is None:
            idx += 1
            continue
        next_child = children[idx + 1]
        while _is_removable_empty_paragraph(next_child):
            body.remove(next_child)
            children.pop(idx + 1)
            removed += 1
            if idx + 1 >= len(children):
                break
            next_child = children[idx + 1]
        idx += 1
    return removed


def _remove_empty_paragraphs_before_page_breaks(body: ET.Element) -> int:
    removed = 0
    children = list(body)
    idx = 0
    while idx + 1 < len(children):
        child = children[idx]
        next_child = children[idx + 1]
        if _is_removable_empty_paragraph(child) and _contains_explicit_page_break(next_child):
            body.remove(child)
            children.pop(idx)
            removed += 1
            continue
        idx += 1
    return removed


def _remove_orphan_empty_body_paragraphs(body: ET.Element) -> int:
    removed = 0
    in_body = False
    for child in list(body):
        if child.tag != _w("p"):
            continue
        text = _paragraph_text(child).strip()
        compact = re.sub(r"\s+", "", text)
        if _is_first_main_heading(text):
            in_body = True
        if in_body and compact in {"参考文献", "致谢", "附录"}:
            break
        if not in_body or text:
            continue
        if _empty_paragraph_is_layout_content(child):
            continue
        body.remove(child)
        removed += 1
    return removed


def _empty_paragraph_is_layout_content(paragraph: ET.Element) -> bool:
    return (
        paragraph.find(".//w:drawing", NS) is not None
        or paragraph.find(".//w:pict", NS) is not None
        or paragraph.find(".//w:object", NS) is not None
        or paragraph.find(".//w:br[@w:type='page']", NS) is not None
        or paragraph.find(".//w:sectPr", NS) is not None
    )


def _remove_layout_paragraphs_before_page_break_headings(body: ET.Element) -> int:
    removed = 0
    children = list(body)
    idx = 0
    while idx + 1 < len(children):
        child = children[idx]
        next_child = children[idx + 1]
        if not _has_page_break_before(next_child):
            idx += 1
            continue
        if child.tag == _w("p") and child.find(".//w:sectPr", NS) is None and (_is_removable_empty_paragraph(child) or _contains_explicit_page_break(child)):
            body.remove(child)
            children.pop(idx)
            removed += 1
            continue
        idx += 1
    return removed


def _has_page_break_before(block: ET.Element) -> bool:
    if block.tag != _w("p"):
        return False
    node = block.find("w:pPr/w:pageBreakBefore", NS)
    if node is None:
        return False
    return node.attrib.get(_w("val")) not in {"0", "false", "False"}


def _restart_page_number_at_first_body_heading(body: ET.Element) -> tuple[bool, int]:
    children = list(body)
    first_heading_idx = None
    for idx, child in enumerate(children):
        if child.tag == _w("p") and _is_first_main_heading(_paragraph_text(child)):
            first_heading_idx = idx
            break
    if first_heading_idx is None:
        return False, 0

    sections = _direct_section_properties(body)
    previous_sections = [(idx, sect) for idx, sect in sections if idx < first_heading_idx]
    if previous_sections:
        body_break_idx, body_break_section = _ensure_body_start_section(body, children, first_heading_idx, previous_sections)
        _set_section_type_next_page(body_break_section)
        sections = _direct_section_properties(body)
        body_sections = [(idx, sect) for idx, sect in sections if idx > first_heading_idx]
        if not body_sections:
            final_sect = body.find("w:sectPr", NS)
            if final_sect is not None:
                body_sections = [(len(children), final_sect)]
        first_body_section_idx = body_sections[0][0] if body_sections else None
        front_count = 0
        first_front = True
        for idx, sect in sections:
            if idx <= body_break_idx:
                if _is_cover_section(sect):
                    _remove_page_number_type(sect)
                    continue
                _set_page_number_type(sect, "upperRoman", "1" if first_front else None)
                first_front = False
                front_count += 1
            elif first_body_section_idx is not None and idx == first_body_section_idx:
                _set_page_number_type(sect, "decimal", "1")
            else:
                _set_page_number_type(sect, "decimal", None)
        return True, front_count

    for idx, sect in sections:
        if idx >= first_heading_idx:
            _set_page_number_type(sect, "decimal", "1")
            return True, 0
    final_sect = body.find("w:sectPr", NS)
    if final_sect is not None:
        _set_page_number_type(final_sect, "decimal", "1")
        return True, 0
    return False, 0


def _ensure_body_start_section(
    body: ET.Element,
    children: list[ET.Element],
    first_heading_idx: int,
    previous_sections: list[tuple[int, ET.Element]],
) -> tuple[int, ET.Element]:
    if first_heading_idx > 0 and children[first_heading_idx - 1].tag != _w("p"):
        p = ET.Element(_w("p"))
        ppr = ET.SubElement(p, _w("pPr"))
        source = previous_sections[-1][1] if previous_sections else None
        sect = copy.deepcopy(source) if source is not None else ET.Element(_w("sectPr"))
        ppr.append(sect)
        body.insert(first_heading_idx, p)
        return first_heading_idx, sect
    for idx in range(first_heading_idx - 1, -1, -1):
        child = children[idx]
        if child.tag != _w("p"):
            continue
        ppr = _ensure_ppr(child)
        existing = ppr.find("w:sectPr", NS)
        if existing is not None:
            return idx, existing
        source = previous_sections[-1][1] if previous_sections else None
        sect = copy.deepcopy(source) if source is not None else ET.Element(_w("sectPr"))
        ppr.append(sect)
        return idx, sect
    return previous_sections[-1]


def _direct_section_properties(body: ET.Element) -> list[tuple[int, ET.Element]]:
    sections: list[tuple[int, ET.Element]] = []
    for idx, child in enumerate(list(body)):
        if child.tag == _w("p"):
            ppr = child.find("w:pPr", NS)
            sect = ppr.find("w:sectPr", NS) if ppr is not None else None
            if sect is not None:
                sections.append((idx, sect))
        elif child.tag == _w("sectPr"):
            sections.append((idx, child))
    return sections


def _set_page_number_type(sect: ET.Element, fmt: str, start: str | None) -> None:
    pg_num = sect.find("w:pgNumType", NS)
    if pg_num is None:
        pg_num = ET.Element(_w("pgNumType"))
        sect.insert(0, pg_num)
    pg_num.set(_w("fmt"), fmt)
    if start is None:
        pg_num.attrib.pop(_w("start"), None)
    else:
        pg_num.set(_w("start"), start)


def _remove_page_number_type(sect: ET.Element) -> None:
    pg_num = sect.find("w:pgNumType", NS)
    if pg_num is not None:
        sect.remove(pg_num)


def _is_cover_section(sect: ET.Element) -> bool:
    page_margin = sect.find("w:pgMar", NS)
    if page_margin is None:
        return False
    zero_margins = all(page_margin.attrib.get(_w(name)) == "0" for name in ("top", "right", "bottom", "left"))
    if not zero_margins:
        return False
    has_header = sect.find("w:headerReference", NS) is not None
    has_footer = sect.find("w:footerReference", NS) is not None
    return not has_header and not has_footer


def _set_section_type_next_page(sect: ET.Element) -> None:
    section_type = sect.find("w:type", NS)
    if section_type is None:
        section_type = ET.Element(_w("type"))
        sect.insert(0, section_type)
    section_type.set(_w("val"), "nextPage")


def _normalize_front_matter_lines(body: ET.Element) -> int:
    fixed = 0
    for child in list(body):
        if child.tag != _w("p"):
            continue
        text = _paragraph_text(child)
        compact = re.sub(r"\s+", "", text)
        if compact in {"摘要", "Abstract", "ABSTRACT", "目录", "目錄"}:
            break
        if text.count("日期") >= 2 and "年" in text and "月" in text and "日" in text:
            _replace_paragraph_visible_text(child, "日期：    年   月   日        日期：    年   月   日")
            fixed += 1
    return fixed


def _normalize_authorization_title_offset(body: ET.Element) -> int:
    children = list(body)
    fixed = 0
    for idx, child in enumerate(children):
        if child.tag != _w("p"):
            continue
        if "版权使用授权书" not in _paragraph_text(child):
            continue
        school_idx = None
        for prev_idx in range(idx - 1, -1, -1):
            previous = children[prev_idx]
            if previous.tag != _w("p"):
                continue
            text = _paragraph_text(previous).strip()
            if text:
                school_idx = prev_idx if text == "上海电机学院" else None
                break
        if school_idx is None:
            continue
        blank_indices: list[int] = []
        for prev_idx in range(school_idx - 1, -1, -1):
            previous = children[prev_idx]
            if previous.tag != _w("p") or _paragraph_text(previous).strip():
                break
            if (
                previous.find(".//w:br[@w:type='page']", NS) is not None
                or previous.find("w:pPr/w:sectPr", NS) is not None
            ):
                break
            blank_indices.append(prev_idx)
        blank_indices = sorted(blank_indices)
        for blank_idx in blank_indices:
            if _set_blank_front_line_spacing(children[blank_idx], "480"):
                fixed += 1
        target_blank_lines = 2
        while len(blank_indices) > target_blank_lines:
            remove_idx = blank_indices.pop(0)
            body.remove(children[remove_idx])
            children.pop(remove_idx)
            blank_indices = [item - 1 if item > remove_idx else item for item in blank_indices]
            school_idx -= 1
            fixed += 1
        while len(blank_indices) < target_blank_lines:
            blank = _front_blank_paragraph("480")
            body.insert(school_idx, blank)
            children.insert(school_idx, blank)
            blank_indices.append(school_idx)
            school_idx += 1
            fixed += 1
        break
    return fixed


def _front_blank_paragraph(line: str) -> ET.Element:
    paragraph = ET.Element(_w("p"))
    ppr = ET.SubElement(paragraph, _w("pPr"))
    spacing = ET.SubElement(ppr, _w("spacing"))
    spacing.set(_w("line"), line)
    spacing.set(_w("lineRule"), "auto")
    return paragraph


def _set_blank_front_line_spacing(paragraph: ET.Element, line: str) -> bool:
    if _paragraph_text(paragraph).strip():
        return False
    ppr = _ensure_ppr(paragraph)
    spacing = ppr.find("w:spacing", NS)
    if spacing is None:
        spacing = ET.Element(_w("spacing"))
        _insert_ppr_child_ordered(ppr, spacing, "spacing")
    changed = spacing.attrib.get(_w("line")) != line or spacing.attrib.get(_w("lineRule")) != "auto"
    spacing.set(_w("line"), line)
    spacing.set(_w("lineRule"), "auto")
    return changed


def _normalize_body_paragraph_format(body: ET.Element) -> int:
    children = list(body)
    first_heading_idx = None
    for idx, child in enumerate(children):
        if child.tag == _w("p") and _is_first_main_heading(_paragraph_text(child)):
            first_heading_idx = idx
            break
    if first_heading_idx is None:
        return 0

    count = 0
    in_references = False
    for child in children[first_heading_idx + 1:]:
        if child.tag != _w("p"):
            continue
        text = _paragraph_text(child).strip()
        compact = re.sub(r"\s+", "", text)
        if compact in {"参考文献", "致谢", "附录"}:
            in_references = compact == "参考文献" or compact in {"致谢", "附录"}
        if in_references:
            continue
        if not _is_normal_body_paragraph(child, text):
            continue
        _set_body_paragraph_format(child)
        count += 1
    return count


def _normalize_table_format(body: ET.Element, template: TableTemplateFormat | None = None) -> int:
    count = 0
    for table in body.iter(_w("tbl")):
        cell_font_size = _table_cell_font_size(table, template.body_font_size if template else "24")
        tbl_pr = _replace_table_pr_from_template(table, template.tbl_pr if template else None)
        _set_table_centered(tbl_pr)
        _ensure_three_line_table_borders(tbl_pr)
        _remove_table_source_layout(tbl_pr)
        _normalize_table_cells(table)
        _normalize_table_grid(table, template.grid_width if template else None)
        _normalize_table_header_border(table)
        for paragraph in table.iter(_w("p")):
            ppr = _ensure_ppr(paragraph)
            spacing = ppr.find("w:spacing", NS)
            if spacing is None:
                spacing = ET.Element(_w("spacing"))
                _insert_ppr_child_ordered(ppr, spacing, "spacing")
            spacing.set(_w("before"), "0")
            spacing.set(_w("after"), "0")
            spacing.set(_w("line"), "240")
            spacing.set(_w("lineRule"), "auto")
            jc_p = ppr.find("w:jc", NS)
            if jc_p is None:
                jc_p = ET.Element(_w("jc"))
                _insert_ppr_child_ordered(ppr, jc_p, "jc")
            jc_p.set(_w("val"), "center")
            for run in paragraph.findall(".//w:r", NS):
                _set_run_font(run, east_asia="宋体", ascii_font="Times New Roman", size=cell_font_size, bold=False)
        count += 1
    return count


def _table_cell_font_size(table: ET.Element, template_size: str) -> str:
    if _table_max_columns(table) >= 6:
        return str(min(int(template_size), 21)) if template_size.isdigit() else "21"
    return template_size


def _replace_table_pr_from_template(table: ET.Element, template_tbl_pr: ET.Element | None) -> ET.Element:
    existing = table.find("w:tblPr", NS)
    if template_tbl_pr is not None:
        tbl_pr = copy.deepcopy(template_tbl_pr)
        if existing is not None:
            table.remove(existing)
        table.insert(0, tbl_pr)
        return tbl_pr
    if existing is None:
        existing = ET.Element(_w("tblPr"))
        table.insert(0, existing)
    return existing


def _set_table_centered(tbl_pr: ET.Element) -> None:
    jc = tbl_pr.find("w:jc", NS)
    if jc is None:
        jc = ET.Element(_w("jc"))
        tbl_pr.append(jc)
    jc.set(_w("val"), "center")


def _remove_table_source_layout(tbl_pr: ET.Element) -> None:
    for name in ("tblStyle", "tblLayout", "shd"):
        existing = tbl_pr.find(f"w:{name}", NS)
        if existing is not None:
            tbl_pr.remove(existing)


def _ensure_three_line_table_borders(tbl_pr: ET.Element) -> None:
    borders = tbl_pr.find("w:tblBorders", NS)
    if borders is None:
        borders = ET.Element(_w("tblBorders"))
        tbl_pr.append(borders)
    for name in ("top", "bottom"):
        border = borders.find(f"w:{name}", NS)
        if border is None:
            border = ET.Element(_w(name))
            borders.append(border)
        border.set(_w("val"), "single")
        border.set(_w("sz"), "4")
        border.set(_w("space"), "0")
        border.set(_w("color"), "auto")


def _normalize_table_cells(table: ET.Element) -> None:
    for tc_pr in table.findall(".//w:tcPr", NS):
        for name in ("shd", "tcBorders", "tcMar", "vAlign", "noWrap", "tcFitText"):
            existing = tc_pr.find(f"w:{name}", NS)
            if existing is not None:
                tc_pr.remove(existing)


def _normalize_table_grid(table: ET.Element, total_width: int | None) -> None:
    columns = _table_max_columns(table)
    if total_width is None or columns <= 0:
        return
    widths = _distributed_widths(total_width, columns)
    grid = table.find("w:tblGrid", NS)
    if grid is None:
        grid = ET.Element(_w("tblGrid"))
        insert_at = 1 if table.find("w:tblPr", NS) is not None else 0
        table.insert(insert_at, grid)
    for child in list(grid):
        grid.remove(child)
    for width in widths:
        col = ET.SubElement(grid, _w("gridCol"))
        col.set(_w("w"), str(width))
    for row in table.findall("w:tr", NS):
        col_idx = 0
        for cell in row.findall("w:tc", NS):
            span = max(1, _cell_grid_span(cell))
            width = sum(widths[col_idx : min(columns, col_idx + span)]) or widths[-1]
            tc_pr = cell.find("w:tcPr", NS)
            if tc_pr is None:
                tc_pr = ET.Element(_w("tcPr"))
                cell.insert(0, tc_pr)
            tc_w = tc_pr.find("w:tcW", NS)
            if tc_w is None:
                tc_w = ET.Element(_w("tcW"))
                tc_pr.insert(0, tc_w)
            tc_w.set(_w("w"), str(width))
            tc_w.set(_w("type"), "dxa")
            col_idx += span


def _normalize_table_header_border(table: ET.Element) -> None:
    first_row = table.find("w:tr", NS)
    if first_row is None:
        return
    for cell in first_row.findall("w:tc", NS):
        tc_pr = cell.find("w:tcPr", NS)
        if tc_pr is None:
            tc_pr = ET.Element(_w("tcPr"))
            cell.insert(0, tc_pr)
        borders = tc_pr.find("w:tcBorders", NS)
        if borders is None:
            borders = ET.Element(_w("tcBorders"))
            tc_pr.append(borders)
        bottom = borders.find("w:bottom", NS)
        if bottom is None:
            bottom = ET.Element(_w("bottom"))
            borders.append(bottom)
        bottom.set(_w("val"), "single")
        bottom.set(_w("sz"), "4")
        bottom.set(_w("space"), "0")
        bottom.set(_w("color"), "auto")


def _table_grid_width(table: ET.Element) -> int | None:
    grid = table.find("w:tblGrid", NS)
    if grid is None:
        return None
    total = 0
    for col in grid.findall("w:gridCol", NS):
        try:
            total += int(col.get(_w("w")) or 0)
        except ValueError:
            return None
    return total or None


def _table_max_columns(table: ET.Element) -> int:
    max_cols = 0
    for row in table.findall("w:tr", NS):
        cols = sum(max(1, _cell_grid_span(cell)) for cell in row.findall("w:tc", NS))
        max_cols = max(max_cols, cols)
    return max_cols


def _cell_grid_span(cell: ET.Element) -> int:
    span = cell.find("w:tcPr/w:gridSpan", NS)
    if span is None:
        return 1
    try:
        return int(span.get(_w("val")) or 1)
    except ValueError:
        return 1


def _distributed_widths(total: int, columns: int) -> list[int]:
    base, remainder = divmod(total, columns)
    return [base + (1 if idx < remainder else 0) for idx in range(columns)]


def _normalize_abstract_and_keywords(body: ET.Element) -> int:
    count = 0
    section: str | None = None
    for child in list(body):
        if child.tag != _w("p"):
            continue
        text = _paragraph_text(child).strip()
        compact = re.sub(r"\s+", "", text)
        if compact in {"目录", "目錄"} or _is_first_main_heading(text):
            break
        if compact == "摘要":
            _format_front_heading(
                child,
                east_asia="黑体",
                ascii_font="黑体",
                size="36",
                bold=True,
                before="0",
                template_spacing="zh_abstract",
            )
            section = "zh_abstract"
            count += 1
            continue
        if compact in {"ABSTRACT"}:
            _format_front_heading(
                child,
                east_asia="Times New Roman",
                ascii_font="Times New Roman",
                size="36",
                bold=True,
                before="624",
                template_spacing="en_abstract",
            )
            section = "en_abstract"
            count += 1
            continue
        if re.match(r"^关键词[:：]", text):
            _normalize_keyword_text(child, english=False)
            _format_keyword_paragraph(child, english=False)
            section = None
            count += 1
            continue
        if re.match(r"^Keywords?[:：]", text, re.IGNORECASE):
            _normalize_keyword_text(child, english=True)
            _format_keyword_paragraph(child, english=True)
            section = None
            count += 1
            continue
        if not text:
            continue
        if section == "zh_abstract":
            _set_abstract_paragraph_format(child, east_asia="宋体", ascii_font="Times New Roman")
            count += 1
        elif section == "en_abstract":
            _set_abstract_paragraph_format(child, east_asia="Times New Roman", ascii_font="Times New Roman")
            count += 1
    return count


def _format_front_heading(
    p: ET.Element,
    east_asia: str,
    ascii_font: str,
    size: str,
    bold: bool,
    before: str = "0",
    template_spacing: str | None = None,
) -> None:
    ppr = _ensure_ppr(p)
    _set_front_heading_spacing(p, ppr, before, template_spacing)
    jc = ppr.find("w:jc", NS)
    if jc is None:
        jc = ET.Element(_w("jc"))
        _insert_ppr_child_ordered(ppr, jc, "jc")
    jc.set(_w("val"), "center")
    ind = ppr.find("w:ind", NS)
    if ind is not None:
        for name in ("firstLine", "firstLineChars", "hanging", "hangingChars"):
            ind.attrib.pop(_w(name), None)
    preserve_template_runs = template_spacing is not None and p.find(".//w:br[@w:type='page']", NS) is not None
    for run in p.findall("w:r", NS):
        if preserve_template_runs and not _run_has_visible_text(run):
            continue
        _set_run_font(run, east_asia=east_asia, ascii_font=ascii_font, size=size, bold=bold)


def _run_has_visible_text(run: ET.Element) -> bool:
    return any((node.text or "").strip() for node in run.findall("w:t", NS))


def _set_front_heading_spacing(p: ET.Element, ppr: ET.Element, before: str, template_spacing: str | None) -> None:
    spacing = ppr.find("w:spacing", NS)
    has_page_break = p.find(".//w:br[@w:type='page']", NS) is not None
    if has_page_break and template_spacing == "zh_abstract":
        if spacing is not None:
            ppr.remove(spacing)
        return
    if has_page_break and template_spacing == "en_abstract":
        if spacing is not None:
            ppr.remove(spacing)
        return
    if spacing is None:
        spacing = ET.Element(_w("spacing"))
        _insert_ppr_child_ordered(ppr, spacing, "spacing")
    spacing.set(_w("before"), before)
    spacing.set(_w("after"), "240")
    spacing.set(_w("line"), "300")
    spacing.set(_w("lineRule"), "auto")


def _set_abstract_paragraph_format(p: ET.Element, east_asia: str, ascii_font: str) -> None:
    ppr = _ensure_ppr(p)
    for name in ("numPr", "keepNext", "keepLines", "pageBreakBefore"):
        existing = ppr.find(f"w:{name}", NS)
        if existing is not None:
            ppr.remove(existing)
    spacing = ppr.find("w:spacing", NS)
    if spacing is None:
        spacing = ET.Element(_w("spacing"))
        _insert_ppr_child_ordered(ppr, spacing, "spacing")
    spacing.set(_w("before"), "240")
    spacing.set(_w("after"), "240")
    spacing.set(_w("line"), "300")
    spacing.set(_w("lineRule"), "auto")
    ind = ppr.find("w:ind", NS)
    if ind is None:
        ind = ET.Element(_w("ind"))
        _insert_ppr_child_ordered(ppr, ind, "ind")
    ind.set(_w("firstLine"), "480")
    ind.set(_w("firstLineChars"), "200")
    jc = ppr.find("w:jc", NS)
    if jc is None:
        jc = ET.Element(_w("jc"))
        _insert_ppr_child_ordered(ppr, jc, "jc")
    jc.set(_w("val"), "both")
    for run in p.findall("w:r", NS):
        _set_run_font(run, east_asia=east_asia, ascii_font=ascii_font, size="24", bold=False)


def _normalize_keyword_text(p: ET.Element, english: bool) -> None:
    text = _paragraph_text(p).strip()
    if english:
        match = re.match(r"^(Keywords?)[:：]\s*(.*)$", text, re.IGNORECASE)
        if not match:
            return
        words = [part.strip(" .;；,，") for part in re.split(r"[,，;；]", match.group(2)) if part.strip(" .;；,，")]
        normalized = "Keywords: " + ", ".join(words)
    else:
        match = re.match(r"^关键词[:：]\s*(.*)$", text)
        if not match:
            return
        words = [part.strip(" 。.;；,，") for part in re.split(r"[,，;；、]", match.group(1)) if part.strip(" 。.;；,，")]
        normalized = "关键词：" + "，".join(words)
    _replace_paragraph_visible_text(p, normalized)


def _format_keyword_paragraph(p: ET.Element, english: bool) -> None:
    ppr = _ensure_ppr(p)
    for name in ("numPr", "keepNext", "keepLines", "pageBreakBefore"):
        existing = ppr.find(f"w:{name}", NS)
        if existing is not None:
            ppr.remove(existing)
    spacing = ppr.find("w:spacing", NS)
    if english:
        if spacing is None:
            spacing = ET.Element(_w("spacing"))
            _insert_ppr_child_ordered(ppr, spacing, "spacing")
        spacing.attrib.clear()
        spacing.set(_w("line"), "360")
        spacing.set(_w("lineRule"), "auto")
    elif spacing is not None:
        ppr.remove(spacing)
    ind = ppr.find("w:ind", NS)
    if ind is not None:
        for name in ("firstLine", "firstLineChars", "hanging", "hangingChars"):
            ind.attrib.pop(_w(name), None)
    east_asia = "Times New Roman" if english else "宋体"
    ascii_font = "Times New Roman" if english else "Times New Roman"
    for run in p.findall("w:r", NS):
        _set_run_font(run, east_asia=east_asia, ascii_font=ascii_font, size="24", bold=False)


def _is_normal_body_paragraph(p: ET.Element, text: str) -> bool:
    if not text or len(text) < 6:
        return False
    if _is_main_heading(text) or _is_sub_heading(text):
        return False
    if "\t" in text or "..." in text or "…" in text:
        return False
    if re.match(r"^(图|表)\s*\d+\s*[-－]\s*\d+", text):
        return False
    if re.match(r"^(\[\d+\]|\d+[\.\u3001、])", text):
        return False
    return bool(re.search(r"[\u4e00-\u9fff]", text))


def _is_sub_heading(text: str) -> bool:
    stripped = text.strip()
    if "\t" in stripped or "..." in stripped or "…" in stripped:
        return False
    return bool(re.match(r"^[1-9]\.\d+(?:\.\d+)?\s+[\u4e00-\u9fffA-Za-z].{0,60}$", stripped))


def _set_body_paragraph_format(p: ET.Element) -> None:
    ppr = _ensure_ppr(p)
    for name in ("numPr", "keepNext", "keepLines", "pageBreakBefore"):
        existing = ppr.find(f"w:{name}", NS)
        if existing is not None:
            ppr.remove(existing)
    spacing = ppr.find("w:spacing", NS)
    if spacing is None:
        spacing = ET.Element(_w("spacing"))
        _insert_ppr_child_ordered(ppr, spacing, "spacing")
    spacing.set(_w("before"), "0")
    spacing.set(_w("after"), "0")
    spacing.set(_w("line"), "300")
    spacing.set(_w("lineRule"), "auto")

    ind = ppr.find("w:ind", NS)
    if ind is None:
        ind = ET.Element(_w("ind"))
        _insert_ppr_child_ordered(ppr, ind, "ind")
    ind.set(_w("firstLine"), "480")
    ind.set(_w("firstLineChars"), "200")

    jc = ppr.find("w:jc", NS)
    if jc is None:
        jc = ET.Element(_w("jc"))
        _insert_ppr_child_ordered(ppr, jc, "jc")
    jc.set(_w("val"), "both")

    for run in p.findall(".//w:r", NS):
        _set_run_font(run, east_asia="宋体", ascii_font="Times New Roman", size="24", bold=False)


def _normalize_acknowledgement_body(body: ET.Element) -> int:
    children = list(body)
    in_ack = False
    count = 0
    for child in children:
        if child.tag != _w("p"):
            continue
        text = _paragraph_text(child).strip()
        compact = re.sub(r"\s+", "", text)
        if compact == "致谢":
            in_ack = True
            continue
        if not in_ack:
            continue
        if not text:
            continue
        if _is_main_heading(text) or compact in {"参考文献", "附录"}:
            break
        _set_acknowledgement_paragraph_format(child)
        count += 1
    return count


def _normalize_reference_paragraphs(body: ET.Element) -> tuple[int, int, int]:
    in_references = False
    count = 0
    removed_empty = 0
    removed_fields = 0
    expected_number = 1
    for child in list(body):
        if child.tag != _w("p"):
            continue
        text = _paragraph_text(child).strip()
        compact = re.sub(r"\s+", "", text)
        if compact == "参考文献":
            in_references = True
            continue
        if in_references and compact in {"致谢", "附录"}:
            break
        if not in_references:
            continue
        if not text:
            if _is_removable_empty_paragraph(child):
                body.remove(child)
                removed_empty += 1
            continue
        if _set_reference_number(child, expected_number):
            removed_fields += 1
        _set_reference_paragraph_format(child)
        count += 1
        expected_number += 1
    return count, removed_empty, removed_fields


def _set_reference_number(p: ET.Element, number: int) -> bool:
    text = _paragraph_text(p).strip()
    body = re.sub(r"^\[\d+\]\s*", "", text, count=1).strip()
    body = _normalize_reference_text(body)
    normalized = f"[{number}] {body}"
    had_field_codes = _has_reference_field_or_inline_style(p)
    if normalized != text or had_field_codes:
        _replace_paragraph_plain_text(p, normalized)
    return had_field_codes


def _normalize_reference_text(text: str) -> str:
    return normalize_reference_text(text)


def _has_reference_field_or_inline_style(p: ET.Element) -> bool:
    return (
        p.find(".//w:fldChar", NS) is not None
        or p.find(".//w:instrText", NS) is not None
        or p.find(".//w:hyperlink", NS) is not None
        or p.find(".//w:rStyle", NS) is not None
    )


def _set_reference_paragraph_format(p: ET.Element) -> None:
    ppr = _ensure_ppr(p)
    for name in ("pStyle", "numPr", "keepNext", "pageBreakBefore", "autoSpaceDE", "autoSpaceDN", "adjustRightInd"):
        existing = ppr.find(f"w:{name}", NS)
        if existing is not None:
            ppr.remove(existing)
    spacing = ppr.find("w:spacing", NS)
    if spacing is None:
        spacing = ET.Element(_w("spacing"))
        _insert_ppr_child_ordered(ppr, spacing, "spacing")
    spacing.set(_w("before"), "0")
    spacing.set(_w("after"), "0")
    spacing.set(_w("line"), "300")
    spacing.set(_w("lineRule"), "auto")
    ind = ppr.find("w:ind", NS)
    if ind is None:
        ind = ET.Element(_w("ind"))
        _insert_ppr_child_ordered(ppr, ind, "ind")
    ind.set(_w("left"), "480")
    ind.set(_w("hanging"), "480")
    ind.set(_w("hangingChars"), "200")
    for name in ("leftChars", "firstLine", "firstLineChars"):
        ind.attrib.pop(_w(name), None)
    jc = ppr.find("w:jc", NS)
    if jc is None:
        jc = ET.Element(_w("jc"))
        _insert_ppr_child_ordered(ppr, jc, "jc")
    jc.set(_w("val"), "left")
    _set_on_off_ppr(p, "keepLines", True)
    for run in p.findall(".//w:r", NS):
        _set_run_font(run, east_asia="宋体", ascii_font="Times New Roman", size="24", bold=False)


def _set_acknowledgement_paragraph_format(p: ET.Element) -> None:
    ppr = _ensure_ppr(p)
    for name in ("numPr", "keepNext", "keepLines", "pageBreakBefore"):
        existing = ppr.find(f"w:{name}", NS)
        if existing is not None:
            ppr.remove(existing)
    spacing = ppr.find("w:spacing", NS)
    if spacing is None:
        spacing = ET.Element(_w("spacing"))
        _insert_ppr_child_ordered(ppr, spacing, "spacing")
    spacing.set(_w("before"), "0")
    spacing.set(_w("after"), "0")
    spacing.set(_w("line"), "240")
    spacing.set(_w("lineRule"), "auto")
    ind = ppr.find("w:ind", NS)
    if ind is None:
        ind = ET.Element(_w("ind"))
        _insert_ppr_child_ordered(ppr, ind, "ind")
    ind.set(_w("firstLine"), "480")
    ind.set(_w("firstLineChars"), "200")
    jc = ppr.find("w:jc", NS)
    if jc is None:
        jc = ET.Element(_w("jc"))
        _insert_ppr_child_ordered(ppr, jc, "jc")
    jc.set(_w("val"), "both")
    for run in p.findall("w:r", NS):
        _set_run_font(run, east_asia="宋体", ascii_font="Times New Roman", size="21", bold=False)


def _apply_main_heading_style(body: ET.Element, style_id: str | None) -> int:
    count = 0
    for child in list(body):
        if child.tag != _w("p"):
            continue
        text = _paragraph_text(child)
        if not _is_main_heading(text):
            continue
        normalized = _normalize_main_heading_text(text)
        if normalized != text.strip():
            _replace_paragraph_visible_text(child, normalized)
            text = normalized
        if style_id:
            _set_paragraph_style(child, style_id)
        _format_main_heading(child, page_break_before=not _is_first_main_heading(text))
        count += 1
    for child in list(body):
        if child.tag != _w("p"):
            continue
        text = re.sub(r"\s+", "", _paragraph_text(child))
        if text not in {"参考文献", "致谢"}:
            continue
        _format_main_heading(child, page_break_before=True)
        count += 1
    return count


def _apply_sub_heading_style(body: ET.Element) -> int:
    count = 0
    first_body_seen = False
    for child in list(body):
        if child.tag != _w("p"):
            continue
        text = _paragraph_text(child).strip()
        compact = re.sub(r"\s+", "", text)
        if _is_first_main_heading(text):
            first_body_seen = True
        if not first_body_seen:
            continue
        if compact in {"参考文献", "致谢", "附录"}:
            break
        if "\t" in text or "..." in text or "…" in text:
            continue
        if re.match(r"^[1-9]\.\d+\.\d+\s*[\u4e00-\u9fffA-Za-z]", text):
            _format_sub_heading(child, east_asia="黑体", ascii_font="黑体", size="24")
            count += 1
        elif re.match(r"^[1-9]\.\d+\s*[\u4e00-\u9fffA-Za-z]", text):
            _format_sub_heading(child, east_asia="宋体", ascii_font="Times New Roman", size="28")
            count += 1
    return count


def _format_sub_heading(p: ET.Element, east_asia: str, ascii_font: str, size: str) -> None:
    ppr = _ensure_ppr(p)
    for name in ("numPr", "pageBreakBefore", "autoSpaceDE", "autoSpaceDN"):
        existing = ppr.find(f"w:{name}", NS)
        if existing is not None:
            ppr.remove(existing)
    spacing = ppr.find("w:spacing", NS)
    if spacing is None:
        spacing = ET.Element(_w("spacing"))
        _insert_ppr_child_ordered(ppr, spacing, "spacing")
    spacing.set(_w("before"), "240")
    spacing.set(_w("after"), "0")
    spacing.set(_w("line"), "300")
    spacing.set(_w("lineRule"), "auto")
    ind = ppr.find("w:ind", NS)
    if ind is not None:
        for name in ("firstLine", "firstLineChars", "hanging", "hangingChars", "left", "leftChars"):
            ind.attrib.pop(_w(name), None)
    jc = ppr.find("w:jc", NS)
    if jc is None:
        jc = ET.Element(_w("jc"))
        _insert_ppr_child_ordered(ppr, jc, "jc")
    jc.set(_w("val"), "left")
    _set_on_off_ppr(p, "keepNext", True)
    _set_on_off_ppr(p, "keepLines", True)
    for run in p.findall("w:r", NS):
        _set_run_font(run, east_asia=east_asia, ascii_font=ascii_font, size=size, bold=False)


def _format_main_heading(p: ET.Element, page_break_before: bool) -> None:
    ppr = _ensure_ppr(p)
    _remove_explicit_page_break_runs(p)
    spacing = ppr.find("w:spacing", NS)
    if spacing is None:
        spacing = ET.Element(_w("spacing"))
        _insert_ppr_child_ordered(ppr, spacing, "spacing")
    spacing.set(_w("before"), "0")
    spacing.set(_w("after"), "240")
    spacing.set(_w("line"), "300")
    spacing.set(_w("lineRule"), "auto")
    jc = ppr.find("w:jc", NS)
    if jc is None:
        jc = ET.Element(_w("jc"))
        _insert_ppr_child_ordered(ppr, jc, "jc")
    jc.set(_w("val"), "center")
    if page_break_before:
        _set_on_off_ppr(p, "pageBreakBefore", True)
    else:
        _remove_ppr_child(p, "pageBreakBefore")
    for name in ("numPr",):
        existing = ppr.find(f"w:{name}", NS)
        if existing is not None:
            ppr.remove(existing)
    _set_on_off_ppr(p, "keepNext", True)
    _set_on_off_ppr(p, "keepLines", True)
    for run in p.findall("w:r", NS):
        _set_run_font(run, east_asia="黑体", ascii_font="黑体", size="36", bold=True)


def _remove_explicit_page_break_runs(paragraph: ET.Element) -> None:
    for run in list(paragraph.findall("w:r", NS)):
        page_breaks = run.findall("w:br[@w:type='page']", NS)
        if not page_breaks:
            continue
        for br in page_breaks:
            run.remove(br)
        if (
            not "".join(node.text or "" for node in run.findall("w:t", NS)).strip()
            and run.find(".//w:drawing", NS) is None
            and run.find(".//w:pict", NS) is None
            and run.find(".//w:object", NS) is None
            and not run.findall("w:br", NS)
        ):
            paragraph.remove(run)


def _normalize_main_heading_text(text: str) -> str:
    stripped = text.strip()
    match = re.match(r"^([1-9])\s*(.+)$", stripped)
    if not match:
        return stripped
    return f"{match.group(1)} {match.group(2).strip()}"


def _apply_caption_keep_rules(body: ET.Element) -> int:
    count = 0
    children = list(body)
    for idx, child in enumerate(children):
        if child.tag != _w("p"):
            continue
        text = _paragraph_text(child).strip()
        if not re.match(r"^(图|表)\s*\d+\s*[-－]\s*\d+", text):
            continue
        _set_on_off_ppr(child, "keepLines", True)
        if text.startswith("表"):
            _set_on_off_ppr(child, "keepNext", True)
        elif text.startswith("图"):
            _remove_ppr_child(child, "keepNext")
            previous = _previous_paragraph_with_content(children, idx)
            previous_text = _paragraph_text(previous).strip() if previous is not None else ""
            if previous is not None and not _is_normal_body_paragraph(previous, previous_text):
                _set_on_off_ppr(previous, "keepNext", True)
                _set_on_off_ppr(previous, "keepLines", True)
                if _paragraph_has_visual(previous):
                    _center_visual_paragraph(previous)
        count += 1
    return count


def _fix_caption_kinds(body: ET.Element) -> int:
    fixed = 0
    children = list(body)
    for idx, child in enumerate(children):
        if child.tag != _w("p"):
            continue
        text = _paragraph_text(child).strip()
        if not re.match(r"^表\s*\d+\s*[-－]\s*\d+", text):
            continue
        previous = _previous_paragraph_with_content(children, idx)
        if previous is None or not _paragraph_has_visual(previous):
            continue
        _replace_paragraph_visible_text(child, re.sub(r"^表", "图", text, count=1))
        fixed += 1
    return fixed


def _normalize_caption_format(body: ET.Element) -> int:
    count = 0
    for child in list(body):
        if child.tag != _w("p"):
            continue
        text = _paragraph_text(child).strip()
        if not re.match(r"^(图|表)\s*\d+\s*[-－]\s*\d+", text):
            continue
        ppr = _ensure_ppr(child)
        for name in (
            "pStyle",
            "numPr",
            "ind",
            "tabs",
            "keepNext",
            "pageBreakBefore",
            "autoSpaceDE",
            "autoSpaceDN",
            "adjustRightInd",
        ):
            existing = ppr.find(f"w:{name}", NS)
            if existing is not None:
                ppr.remove(existing)
        spacing = ppr.find("w:spacing", NS)
        if spacing is None:
            spacing = ET.Element(_w("spacing"))
            _insert_ppr_child_ordered(ppr, spacing, "spacing")
        spacing.set(_w("before"), "0")
        spacing.set(_w("after"), "0")
        spacing.set(_w("line"), "240")
        spacing.set(_w("lineRule"), "auto")
        jc = ppr.find("w:jc", NS)
        if jc is None:
            jc = ET.Element(_w("jc"))
            _insert_ppr_child_ordered(ppr, jc, "jc")
        jc.set(_w("val"), "center")
        for run in child.findall("w:r", NS):
            _set_run_font(run, east_asia="宋体", ascii_font="Times New Roman", size="18", bold=True)
        count += 1
    return count


def _paragraph_has_visual(paragraph: ET.Element) -> bool:
    return (
        paragraph.find(".//w:drawing", NS) is not None
        or paragraph.find(".//w:pict", NS) is not None
        or paragraph.find(".//w:object", NS) is not None
    )


def _center_visual_paragraph(paragraph: ET.Element) -> None:
    ppr = _ensure_ppr(paragraph)
    for name in ("ind", "tabs"):
        existing = ppr.find(f"w:{name}", NS)
        if existing is not None:
            ppr.remove(existing)
    jc = ppr.find("w:jc", NS)
    if jc is None:
        jc = ET.Element(_w("jc"))
        _insert_ppr_child_ordered(ppr, jc, "jc")
    jc.set(_w("val"), "center")


def _fix_toc_headings(body: ET.Element) -> int:
    fixed = 0
    for container in _paragraph_sequence_containers(body):
        fixed += _fix_toc_headings_in_container(container)
    return fixed


def _normalize_toc_content_controls(body: ET.Element) -> int:
    inserted_breaks = 0
    children = list(body)
    idx = 0
    while idx < len(children):
        child = children[idx]
        if child.tag != _w("sdt"):
            idx += 1
            continue
        first = _first_nonempty_paragraph(child)
        if first is None or _toc_heading_key(first) not in {"目录", "目錄"}:
            idx += 1
            continue

        _clear_visible_text(first)
        _remove_ppr_child(first, "pageBreakBefore")
        _remove_ppr_child(first, "keepNext")
        _remove_ppr_child(first, "keepLines")

        heading = _clean_toc_heading_paragraph(first)
        body.insert(idx, heading)
        children.insert(idx, heading)
        if idx > 0 and not _contains_explicit_page_break(children[idx - 1]) and not _contains_section_properties(children[idx - 1]):
            page_break = _page_break_paragraph()
            body.insert(idx, page_break)
            children.insert(idx, page_break)
            inserted_breaks += 1
            idx += 1
        idx += 2
    return inserted_breaks


def _first_nonempty_paragraph(block: ET.Element) -> ET.Element | None:
    for paragraph in block.iter(_w("p")):
        if _paragraph_text(paragraph).strip():
            return paragraph
    return None


def _toc_heading_key(paragraph: ET.Element) -> str:
    return re.sub(r"\s+", "", _paragraph_text(paragraph))


def _clear_visible_text(paragraph: ET.Element) -> None:
    for node in paragraph.findall(".//w:t", NS):
        node.text = ""


def _clean_toc_heading_paragraph(source: ET.Element) -> ET.Element:
    paragraph = ET.Element(_w("p"))
    source_ppr = source.find("w:pPr", NS)
    if source_ppr is not None:
        ppr = copy.deepcopy(source_ppr)
        for name in ("sectPr", "pageBreakBefore"):
            existing = ppr.find(f"w:{name}", NS)
            if existing is not None:
                ppr.remove(existing)
        paragraph.append(ppr)
    run = ET.SubElement(paragraph, _w("r"))
    source_rpr = source.find(".//w:rPr", NS)
    if source_rpr is not None:
        run.append(copy.deepcopy(source_rpr))
    text = ET.SubElement(run, _w("t"))
    text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text.text = "目  录"
    return paragraph


def _ensure_page_break_before_toc_blocks(body: ET.Element) -> int:
    inserted = 0
    children = list(body)
    idx = 0
    while idx < len(children):
        child = children[idx]
        if not _block_starts_with_toc_heading(child):
            idx += 1
            continue
        if idx == 0 or _contains_explicit_page_break(children[idx - 1]) or _contains_section_properties(children[idx - 1]):
            idx += 1
            continue
        page_break = _page_break_paragraph()
        body.insert(idx, page_break)
        children.insert(idx, page_break)
        inserted += 1
        idx += 2
    return inserted


def _move_toc_heading_section_breaks_to_previous(body: ET.Element) -> int:
    moved = 0
    children = list(body)
    for idx, child in enumerate(children):
        if child.tag != _w("p") or _toc_heading_key(child) not in {"目录", "目錄"}:
            continue
        source_ppr = child.find("w:pPr", NS)
        sect = source_ppr.find("w:sectPr", NS) if source_ppr is not None else None
        if sect is None:
            continue
        target = _previous_paragraph_child(children, idx)
        if target is None:
            continue
        source_ppr.remove(sect)
        target_ppr = _ensure_ppr(target)
        existing = target_ppr.find("w:sectPr", NS)
        if existing is not None:
            target_ppr.remove(existing)
        target_ppr.append(sect)
        moved += 1
    return moved


def _previous_paragraph_child(children: list[ET.Element], start_idx: int) -> ET.Element | None:
    for previous in reversed(children[:start_idx]):
        if previous.tag == _w("p"):
            return previous
    return None


def _block_starts_with_toc_heading(block: ET.Element) -> bool:
    for paragraph in block.iter(_w("p")):
        text = re.sub(r"\s+", "", _paragraph_text(paragraph))
        if not text:
            continue
        return text in {"目录", "目錄"}
    return False


def _contains_explicit_page_break(block: ET.Element) -> bool:
    return block.find(".//w:br[@w:type='page']", NS) is not None


def _contains_section_properties(block: ET.Element) -> bool:
    return block.find(".//w:sectPr", NS) is not None


def _paragraph_sequence_containers(body: ET.Element) -> list[ET.Element]:
    containers = [body]
    containers.extend(body.findall(".//w:sdtContent", NS))
    containers.extend(body.findall(".//w:tc", NS))
    return containers


def _fix_toc_headings_in_container(container: ET.Element) -> int:
    fixed = 0
    children = list(container)
    idx = 0
    while idx < len(children):
        child = children[idx]
        if child.tag != _w("p"):
            idx += 1
            continue
        text = re.sub(r"\s+", "", _paragraph_text(child))
        if text == "目" and idx + 1 < len(children) and children[idx + 1].tag == _w("p"):
            next_text = re.sub(r"\s+", "", _paragraph_text(children[idx + 1]))
            if next_text == "录":
                _replace_paragraph_plain_text(child, "目  录")
                _move_section_properties(children[idx + 1], child)
                container.remove(children[idx + 1])
                children.pop(idx + 1)
                _format_toc_heading(child)
                fixed += 1
        elif text in {"目录", "目錄"}:
            _replace_paragraph_plain_text(child, "目  录")
            _format_toc_heading(child)
            fixed += 1
        idx += 1
    return fixed


def _format_toc_heading(p: ET.Element) -> None:
    ppr = _ensure_ppr(p)
    spacing = ppr.find("w:spacing", NS)
    if spacing is None:
        spacing = ET.Element(_w("spacing"))
        _insert_ppr_child_ordered(ppr, spacing, "spacing")
    spacing.set(_w("before"), "0")
    spacing.set(_w("after"), "240")
    spacing.set(_w("line"), "300")
    spacing.set(_w("lineRule"), "auto")
    jc = ppr.find("w:jc", NS)
    if jc is None:
        jc = ET.Element(_w("jc"))
        _insert_ppr_child_ordered(ppr, jc, "jc")
    jc.set(_w("val"), "center")
    _set_on_off_ppr(p, "keepNext", True)
    _set_on_off_ppr(p, "keepLines", True)
    for run in p.findall(".//w:r", NS):
        _set_run_font(run, east_asia="黑体", ascii_font="Times New Roman", size="36", bold=True)


def _set_run_font(run: ET.Element, east_asia: str, ascii_font: str, size: str, bold: bool) -> None:
    rpr = run.find("w:rPr", NS)
    if rpr is None:
        rpr = ET.Element(_w("rPr"))
        run.insert(0, rpr)
    fonts = rpr.find("w:rFonts", NS)
    if fonts is None:
        fonts = ET.Element(_w("rFonts"))
        rpr.insert(0, fonts)
    fonts.set(_w("eastAsia"), east_asia)
    fonts.set(_w("ascii"), ascii_font)
    fonts.set(_w("hAnsi"), ascii_font)
    for name in ("eastAsiaTheme", "asciiTheme", "hAnsiTheme", "cstheme"):
        fonts.attrib.pop(_w(name), None)

    for name in ("sz", "szCs"):
        node = rpr.find(f"w:{name}", NS)
        if node is None:
            node = ET.Element(_w(name))
            rpr.append(node)
        node.set(_w("val"), size)

    b = rpr.find("w:b", NS)
    if bold:
        if b is None:
            rpr.append(ET.Element(_w("b")))
    elif b is not None:
        rpr.remove(b)


def _remove_ppr_child(paragraph: ET.Element, name: str) -> None:
    ppr = paragraph.find("w:pPr", NS)
    if ppr is None:
        return
    existing = ppr.find(f"w:{name}", NS)
    if existing is not None:
        ppr.remove(existing)


def _replace_paragraph_visible_text(p: ET.Element, text: str) -> None:
    text_nodes = list(p.findall(".//w:t", NS))
    if not text_nodes:
        run = ET.SubElement(p, _w("r"))
        node = ET.SubElement(run, _w("t"))
        node.text = text
        return
    text_nodes[0].text = text
    if re.search(r"\s", text):
        text_nodes[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    for node in text_nodes[1:]:
        node.text = ""


def _replace_paragraph_plain_text(p: ET.Element, text: str) -> None:
    ppr = p.find("w:pPr", NS)
    bookmark_starts = [copy.deepcopy(node) for node in p.findall(".//w:bookmarkStart", NS)]
    bookmark_ends = [copy.deepcopy(node) for node in p.findall(".//w:bookmarkEnd", NS)]
    for child in list(p):
        if ppr is not None and child is ppr:
            continue
        p.remove(child)
    for bookmark in bookmark_starts:
        p.append(bookmark)
    run = ET.SubElement(p, _w("r"))
    node = ET.SubElement(run, _w("t"))
    if re.search(r"\s", text):
        node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    node.text = text
    for bookmark in bookmark_ends:
        p.append(bookmark)


def _move_section_properties(source: ET.Element, target: ET.Element) -> None:
    source_ppr = source.find("w:pPr", NS)
    if source_ppr is None:
        return
    sect = source_ppr.find("w:sectPr", NS)
    if sect is None:
        return
    source_ppr.remove(sect)
    target_ppr = _ensure_ppr(target)
    old = target_ppr.find("w:sectPr", NS)
    if old is not None:
        target_ppr.remove(old)
    target_ppr.append(sect)


def _previous_paragraph_with_content(children: list[ET.Element], start_idx: int) -> ET.Element | None:
    for previous in reversed(children[:start_idx]):
        if previous.tag == _w("p") and _has_visible_content(previous):
            return previous
        if previous.tag == _w("tbl"):
            return None
    return None


def _find_heading_style_id(styles_xml: bytes | None, style_name: str) -> str | None:
    if not styles_xml:
        return None
    root = ET.fromstring(styles_xml)
    wanted = style_name.strip().lower()
    for style in root.findall("w:style", NS):
        if style.attrib.get(_w("type")) != "paragraph":
            continue
        name = style.find("w:name", NS)
        value = name.attrib.get(_w("val"), "") if name is not None else ""
        if value.strip().lower() == wanted:
            return style.attrib.get(_w("styleId"))
    return None


def _set_paragraph_style(p: ET.Element, style_id: str) -> None:
    ppr = _ensure_ppr(p)
    style = ppr.find("w:pStyle", NS)
    if style is None:
        style = ET.Element(_w("pStyle"))
        ppr.insert(0, style)
    style.set(_w("val"), style_id)


def _set_on_off_ppr(p: ET.Element, name: str, enabled: bool) -> None:
    ppr = _ensure_ppr(p)
    node = ppr.find(f"w:{name}", NS)
    if node is not None:
        ppr.remove(node)
    else:
        node = ET.Element(_w(name))
    _insert_ppr_child_ordered(ppr, node, name)
    if not enabled:
        node.set(_w("val"), "0")
    elif _w("val") in node.attrib:
        node.attrib.pop(_w("val"))


def _insert_ppr_child_ordered(ppr: ET.Element, node: ET.Element, name: str) -> None:
    order = {
        "pStyle": 10,
        "keepNext": 20,
        "keepLines": 30,
        "pageBreakBefore": 40,
        "widowControl": 50,
        "numPr": 60,
        "spacing": 120,
        "ind": 130,
        "jc": 160,
        "sectPr": 1000,
    }
    target_order = order.get(name, 500)
    insert_at = 0
    for idx, child in enumerate(list(ppr)):
        child_name = child.tag.split("}", 1)[-1]
        if order.get(child_name, 500) <= target_order:
            insert_at = idx + 1
    ppr.insert(insert_at, node)


def _ensure_ppr(p: ET.Element) -> ET.Element:
    ppr = p.find("w:pPr", NS)
    if ppr is None:
        ppr = ET.Element(_w("pPr"))
        p.insert(0, ppr)
    return ppr


def _is_first_main_heading(text: str) -> bool:
    return bool(re.match(r"^1\s+绪论\s*$", text.strip()))


def _is_main_heading(text: str) -> bool:
    stripped = text.strip()
    if "\t" in stripped or "..." in stripped or "…" in stripped:
        return False
    return bool(re.match(r"^[1-9]\s*[\u4e00-\u9fffA-Za-z].{0,40}$", stripped))


def _is_removable_empty_paragraph(el: ET.Element) -> bool:
    return (
        el.tag == _w("p")
        and not _has_visible_content(el)
        and el.find(".//w:sectPr", NS) is None
        and el.find(".//w:br[@w:type='page']", NS) is None
    )


def _has_visible_content(el: ET.Element) -> bool:
    if el.tag == _w("tbl"):
        return True
    if el.tag != _w("p"):
        return any(_has_visible_content(child) for child in list(el))
    if _paragraph_text(el).strip():
        return True
    for tag in ("drawing", "pict", "object", "fldSimple"):
        if el.find(f".//w:{tag}", NS) is not None:
            return True
    return False


def _serialize_xml(root: ET.Element) -> bytes:
    return serialize_xml(root)


def _serialize_package_xml(root: ET.Element, default_namespace: str) -> bytes:
    return serialize_package_xml(root, default_namespace)


def _w(local: str) -> str:
    return f"{{{W_NS}}}{local}"
