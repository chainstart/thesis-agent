from __future__ import annotations

import copy
import json
import re
import zipfile
from dataclasses import asdict, dataclass, field
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET

from .docx_inspect import NS, W_NS, _paragraph_text
from .ooxml import serialize_package_xml, serialize_xml


R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)


@dataclass(frozen=True)
class RebuildReport:
    template: Path
    source: Path
    output: Path
    title: str | None = None
    cover_elements: int = 0
    template_front_matter_elements: int = 0
    abstract_elements: int = 0
    toc_entries: int = 0
    body_elements: int = 0
    reference_elements: int = 0
    acknowledgement_elements: int = 0
    appendix_elements: int = 0
    imported_relationships: int = 0
    imported_parts: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    def to_json(self) -> str:
        return json.dumps(asdict(self), ensure_ascii=False, indent=2, default=str)


@dataclass(frozen=True)
class ExtractedContent:
    title: str | None
    cover: list[ET.Element]
    abstract: list[ET.Element]
    body: list[ET.Element]
    references: list[ET.Element]
    acknowledgements: list[ET.Element]
    appendices: list[ET.Element]
    toc_titles: list[str]
    warnings: list[str] = field(default_factory=list)


def rebuild_thesis_docx(template_path: Path, source_docx: Path, output_path: Path) -> RebuildReport:
    if source_docx.suffix.lower() != ".docx":
        raise ValueError("rebuild currently supports .docx sources only")
    if not zipfile.is_zipfile(source_docx):
        raise ValueError(f"Not a valid docx file: {source_docx}")
    template_docx = _resolve_template_docx(template_path)
    if template_docx is None:
        raise ValueError(f"No usable .docx template found for {template_path}")

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

        extracted = _extract_source_content(source_body)
        warnings.extend(extracted.warnings)
        template_front = _extract_template_front_matter(template_body)
        template_final_section = _template_final_section(template_body)

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

        source_owned_elements: list[ET.Element] = []
        source_owned_elements.extend(extracted.abstract)
        source_owned_elements.extend(extracted.body)
        source_owned_elements.extend(extracted.references)
        source_owned_elements.extend(extracted.acknowledgements)
        source_owned_elements.extend(extracted.appendices)

        import_result = _import_source_relationships(
            elements=source_owned_elements,
            source_zip=source_zip,
            target_rels=target_rels,
            content_types=content_types,
            existing_names=set(template_zip.namelist()),
        )
        source_owned_ids = {id(el) for el in source_owned_elements}

        new_elements: list[ET.Element] = []
        # The target document is rebuilt from the standard template package.
        # Student cover pages are frequently non-standard and can collide with
        # the template declarations, so the cover is used only for metadata
        # inference and is not inserted into the rebuilt body.
        new_elements.extend(copy.deepcopy(el) for el in template_front)
        new_elements.extend(extracted.abstract)
        _append_page_break_if_needed(new_elements)
        new_elements.append(_toc_heading_paragraph())
        new_elements.extend(_toc_entry_paragraph(title) for title in extracted.toc_titles)
        _append_page_break_if_needed(new_elements)
        new_elements.extend(extracted.body)
        if extracted.references:
            if not _starts_with_heading(extracted.references, "参考文献"):
                new_elements.append(_main_heading_paragraph("参考文献"))
            new_elements.extend(extracted.references)
        if extracted.acknowledgements:
            if not _starts_with_heading(extracted.acknowledgements, "致谢"):
                new_elements.append(_page_break_paragraph())
                new_elements.append(_main_heading_paragraph("致  谢"))
            new_elements.extend(extracted.acknowledgements)
        if extracted.appendices:
            new_elements.extend(extracted.appendices)
        if template_final_section is not None:
            new_elements.append(copy.deepcopy(template_final_section))

        for child in list(template_body):
            template_body.remove(child)
        for element in new_elements:
            if id(element) in source_owned_ids:
                template_body.append(_sanitize_imported_body_element(element))
            else:
                template_body.append(copy.deepcopy(element))

        document_xml = _serialize_xml(template_root)
        rels_xml = _serialize_package_xml(target_rels, REL_NS)
        content_types_data = _serialize_package_xml(content_types, CT_NS)
        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as out_zip:
            written = set()
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
            for name, data in import_result.parts.items():
                if name not in written:
                    out_zip.writestr(name, data)
                    written.add(name)

    return RebuildReport(
        template=template_docx,
        source=source_docx,
        output=output_path,
        title=extracted.title,
        cover_elements=len(extracted.cover),
        template_front_matter_elements=len(template_front),
        abstract_elements=len(extracted.abstract),
        toc_entries=len(extracted.toc_titles),
        body_elements=len(extracted.body),
        reference_elements=len(extracted.references),
        acknowledgement_elements=len(extracted.acknowledgements),
        appendix_elements=len(extracted.appendices),
        imported_relationships=import_result.relationships,
        imported_parts=sorted(import_result.parts),
        warnings=warnings,
    )


def _resolve_template_docx(template_path: Path) -> Path | None:
    if template_path.name.startswith("~$"):
        return None
    if template_path.suffix.lower() == ".docx" and template_path.exists():
        return template_path
    candidate = template_path.with_suffix(".docx")
    if not candidate.name.startswith("~$") and candidate.exists():
        return candidate
    return None


def _extract_source_content(source_body: ET.Element) -> ExtractedContent:
    children = [copy.deepcopy(el) for el in list(source_body) if el.tag != _w("sectPr")]
    paragraphs = [(idx, _paragraph_text(el).strip()) for idx, el in enumerate(children) if el.tag == _w("p")]
    block_texts = [(idx, _element_text_with_fields(el).strip()) for idx, el in enumerate(children)]
    first_decl = _first_index(paragraphs, lambda text: _compact(text) in {"毕业设计（论文）学术诚信声明", "毕业设计（论文）AI使用情况声明", "毕业设计（论文）版权使用授权书"})
    abstract_idx = _first_index(paragraphs, _is_abstract_heading)
    toc_idx = _first_index(block_texts, _is_toc_heading)
    first_body_idx = _first_index(paragraphs, _is_first_body_heading)
    reference_idx = _first_index(paragraphs, lambda text: _compact(text) == "参考文献")
    ack_idx = _first_index(paragraphs, lambda text: _compact(text) in {"致谢", "致謝"})
    appendix_idx = _first_index(paragraphs, lambda text: _compact(text).startswith("附录") or _compact(text).startswith("附件"))

    content_start_candidates = [idx for idx in [first_decl, abstract_idx, toc_idx, first_body_idx] if idx is not None]
    content_start = min(content_start_candidates) if content_start_candidates else 0
    cover = children[:content_start] if _looks_like_cover(children[:content_start]) else []

    abstract_start = abstract_idx
    abstract_end_candidates = [idx for idx in [toc_idx, first_body_idx] if idx is not None and abstract_start is not None and idx > abstract_start]
    abstract_end = min(abstract_end_candidates) if abstract_end_candidates else (first_body_idx if first_body_idx is not None else len(children))
    abstract = children[abstract_start:abstract_end] if abstract_start is not None else []

    body_start = first_body_idx
    body_end_candidates = [idx for idx in [reference_idx, ack_idx, appendix_idx] if idx is not None and body_start is not None and idx > body_start]
    body_end = min(body_end_candidates) if body_end_candidates else len(children)
    body = children[body_start:body_end] if body_start is not None else []

    references: list[ET.Element] = []
    if reference_idx is not None:
        ref_end_candidates = [idx for idx in [ack_idx, appendix_idx] if idx is not None and idx > reference_idx]
        ref_end = min(ref_end_candidates) if ref_end_candidates else len(children)
        references = children[reference_idx:ref_end]

    acknowledgements: list[ET.Element] = []
    if ack_idx is not None:
        ack_end = appendix_idx if appendix_idx is not None and appendix_idx > ack_idx else len(children)
        acknowledgements = children[ack_idx:ack_end]

    appendices: list[ET.Element] = []
    if appendix_idx is not None:
        appendices = children[appendix_idx:len(children)]

    title = _infer_title(cover, children)
    toc_titles = _toc_titles_from_body(body, references, acknowledgements, appendices)
    warnings: list[str] = []
    if not abstract:
        warnings.append("未从源文档识别到摘要区域，重建稿将缺少源摘要。")
    if not body:
        warnings.append("未从源文档识别到 `1 绪论` 起始正文，重建稿正文可能不完整。")
    if not references:
        warnings.append("未从源文档识别到参考文献区域。")
    return ExtractedContent(
        title=title,
        cover=cover,
        abstract=abstract,
        body=body,
        references=references,
        acknowledgements=acknowledgements,
        appendices=appendices,
        toc_titles=toc_titles,
        warnings=warnings,
    )


def _extract_template_front_matter(template_body: ET.Element) -> list[ET.Element]:
    elements: list[ET.Element] = []
    for child in list(template_body):
        text = _paragraph_text(child) if child.tag == _w("p") else ""
        if _is_template_abstract_marker(text):
            break
        if child.tag == _w("sectPr"):
            continue
        elements.append(copy.deepcopy(child))
    return elements


def _is_template_abstract_marker(text: str) -> bool:
    compact = _compact(text)
    return "摘要正文" in compact or "小二号黑体摘要" in compact or compact in {"摘要", "摘 要"}


def _template_final_section(template_body: ET.Element) -> ET.Element | None:
    direct = template_body.find("w:sectPr", NS)
    if direct is not None:
        return direct
    for child in reversed(list(template_body)):
        if child.tag == _w("p"):
            sect = child.find("w:pPr/w:sectPr", NS)
            if sect is not None:
                return sect
    return None


def _looks_like_cover(elements: list[ET.Element]) -> bool:
    text = "\n".join(_paragraph_text(el) for el in elements if el.tag == _w("p"))
    compact = _compact(text)
    return any(marker in compact for marker in ["学士学位论文", "学生姓名", "学生学号", "指导教师", "学院"])


def _infer_title(cover: list[ET.Element], all_children: list[ET.Element]) -> str | None:
    candidates = [_paragraph_text(el).strip() for el in (cover or all_children[:30]) if el.tag == _w("p")]
    for idx, text in enumerate(candidates):
        if _compact(text) in {"学士学位论文", "本科毕业设计（论文）", "毕业设计（论文）"}:
            for candidate in candidates[idx + 1: idx + 5]:
                if _looks_like_title(candidate):
                    return candidate.strip()
    for candidate in candidates:
        if _looks_like_title(candidate):
            return candidate.strip()
    return None


def _looks_like_title(text: str) -> bool:
    compact = _compact(text)
    if len(compact) < 8 or len(compact) > 40:
        return False
    if "：" in compact or ":" in compact:
        return False
    banned = ["学校名称", "学生姓名", "学生学号", "指导教师", "专业", "学院", "大学", "摘要", "目录", "声明", "授权书", "关键词", "关键字", "KeyWords", "Keywords"]
    return bool(re.search(r"[\u4e00-\u9fff]", compact)) and not any(item in compact for item in banned)


def _toc_titles_from_body(
    body: list[ET.Element],
    references: list[ET.Element],
    acknowledgements: list[ET.Element],
    appendices: list[ET.Element] | None = None,
) -> list[str]:
    titles: list[str] = []
    seen: set[str] = set()
    for element in body:
        if element.tag != _w("p"):
            continue
        text = _paragraph_text(element).strip()
        if _is_toc_heading_candidate(text):
            key = _compact(text)
            if key not in seen:
                titles.append(_normalize_heading_text(text))
                seen.add(key)
    if references:
        titles.append("参考文献")
    if acknowledgements:
        titles.append("致谢")
    if appendices:
        titles.append("附录")
    return titles


def _is_toc_heading_candidate(text: str) -> bool:
    stripped = text.strip()
    if re.match(r"^[1-9]\s*[\u4e00-\u9fffA-Za-z].{0,60}$", stripped):
        return True
    if re.match(r"^[1-9]\.\d+\s*[\u4e00-\u9fffA-Za-z].{0,60}$", stripped):
        return True
    return False


def _normalize_heading_text(text: str) -> str:
    return re.sub(r"\s+", " ", text.strip())


def _toc_heading_paragraph() -> ET.Element:
    return _simple_paragraph("目  录", kind="toc_heading")


def _toc_entry_paragraph(title: str) -> ET.Element:
    p = ET.Element(_w("p"))
    ppr = ET.SubElement(p, _w("pPr"))
    tabs = ET.SubElement(ppr, _w("tabs"))
    tab = ET.SubElement(tabs, _w("tab"))
    tab.set(_w("val"), "right")
    tab.set(_w("leader"), "dot")
    tab.set(_w("pos"), "8300")
    r = ET.SubElement(p, _w("r"))
    _set_run_font(r, east_asia="黑体", ascii_font="Times New Roman", size="28" if _toc_level(title) == 1 else "24")
    t = ET.SubElement(r, _w("t"))
    t.text = title
    if re.search(r"\s", title):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    ET.SubElement(r, _w("tab"))
    page = ET.SubElement(r, _w("t"))
    page.text = "1"
    return p


def _toc_level(title: str) -> int:
    return 2 if re.match(r"^[1-9]\.\d+", title.strip()) else 1


def _main_heading_paragraph(text: str) -> ET.Element:
    return _simple_paragraph(text, kind="main")


def _simple_paragraph(text: str, kind: str = "body") -> ET.Element:
    p = ET.Element(_w("p"))
    ppr = ET.SubElement(p, _w("pPr"))
    spacing = ET.SubElement(ppr, _w("spacing"))
    spacing.set(_w("before"), "0")
    spacing.set(_w("after"), "240" if kind in {"main", "toc_heading"} else "0")
    spacing.set(_w("line"), "300")
    spacing.set(_w("lineRule"), "auto")
    jc = ET.SubElement(ppr, _w("jc"))
    jc.set(_w("val"), "center" if kind in {"main", "toc_heading"} else "both")
    r = ET.SubElement(p, _w("r"))
    _set_run_font(r, east_asia="黑体" if kind in {"main", "toc_heading"} else "宋体", ascii_font="Times New Roman", size="36" if kind in {"main", "toc_heading"} else "24")
    t = ET.SubElement(r, _w("t"))
    t.text = text
    if re.search(r"\s", text):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return p


def _page_break_paragraph() -> ET.Element:
    p = ET.Element(_w("p"))
    r = ET.SubElement(p, _w("r"))
    br = ET.SubElement(r, _w("br"))
    br.set(_w("type"), "page")
    return p


def _append_page_break_if_needed(elements: list[ET.Element]) -> None:
    for element in reversed(elements):
        if _has_visible_content(element):
            if not _contains_page_boundary(element):
                elements.append(_page_break_paragraph())
            return
        if _contains_page_boundary(element):
            return
    elements.append(_page_break_paragraph())


def _contains_page_boundary(element: ET.Element) -> bool:
    return element.find(".//w:br[@w:type='page']", NS) is not None or element.find(".//w:sectPr", NS) is not None


def _has_visible_content(element: ET.Element) -> bool:
    if element.tag == _w("tbl"):
        return True
    if element.find(".//w:drawing", NS) is not None or element.find(".//w:pict", NS) is not None:
        return True
    return bool(_paragraph_text(element).strip()) if element.tag == _w("p") else False


def _starts_with_heading(elements: list[ET.Element], heading: str) -> bool:
    for element in elements:
        if element.tag != _w("p"):
            continue
        text = _compact(_paragraph_text(element))
        if not text:
            continue
        return text == _compact(heading)
    return False


def _sanitize_imported_body_element(element: ET.Element) -> ET.Element:
    cleaned = copy.deepcopy(element)
    for table in cleaned.iter(_w("tbl")):
        _normalize_imported_table(table)
    for paragraph in cleaned.iter(_w("p")):
        for attr in list(paragraph.attrib):
            if attr.endswith("paraId") or attr.endswith("textId") or "rsid" in attr:
                paragraph.attrib.pop(attr, None)
        ppr = paragraph.find("w:pPr", NS)
        if ppr is not None:
            for name in ("sectPr", "pStyle", "numPr", "pageBreakBefore"):
                for node in list(ppr.findall(f"w:{name}", NS)):
                    ppr.remove(node)
        for br in list(paragraph.findall(".//w:br[@w:type='page']", NS)):
            parent = _parent_map(paragraph).get(br)
            if parent is not None:
                parent.remove(br)
    for sect in list(cleaned.findall(".//w:sectPr", NS)):
        parent = _parent_map(cleaned).get(sect)
        if parent is not None:
            parent.remove(sect)
    return cleaned


def _normalize_imported_table(table: ET.Element) -> None:
    tbl_pr = table.find("w:tblPr", NS)
    if tbl_pr is None:
        tbl_pr = ET.Element(_w("tblPr"))
        table.insert(0, tbl_pr)
    for name in ("tblStyle", "tblLayout", "tblLook"):
        for node in list(tbl_pr.findall(f"w:{name}", NS)):
            tbl_pr.remove(node)
    jc = tbl_pr.find("w:jc", NS)
    if jc is None:
        jc = ET.SubElement(tbl_pr, _w("jc"))
    jc.set(_w("val"), "center")
    borders = tbl_pr.find("w:tblBorders", NS)
    if borders is None:
        borders = ET.SubElement(tbl_pr, _w("tblBorders"))
    for name, value in (("top", "single"), ("bottom", "single"), ("left", "nil"), ("right", "nil"), ("insideH", "nil"), ("insideV", "nil")):
        border = borders.find(f"w:{name}", NS)
        if border is None:
            border = ET.SubElement(borders, _w(name))
        border.set(_w("val"), value)
        if value == "single":
            border.set(_w("sz"), "4")
            border.set(_w("space"), "0")
            border.set(_w("color"), "000000")
    for paragraph in table.iter(_w("p")):
        ppr = paragraph.find("w:pPr", NS)
        if ppr is None:
            ppr = ET.Element(_w("pPr"))
            paragraph.insert(0, ppr)
        for name in ("pStyle", "numPr", "ind", "tabs", "pageBreakBefore"):
            for node in list(ppr.findall(f"w:{name}", NS)):
                ppr.remove(node)
        spacing = ppr.find("w:spacing", NS)
        if spacing is None:
            spacing = ET.SubElement(ppr, _w("spacing"))
        spacing.set(_w("before"), "0")
        spacing.set(_w("after"), "0")
        spacing.set(_w("line"), "240")
        spacing.set(_w("lineRule"), "auto")
        jc = ppr.find("w:jc", NS)
        if jc is None:
            jc = ET.SubElement(ppr, _w("jc"))
        jc.set(_w("val"), "center")
        for run in paragraph.findall("w:r", NS):
            _normalize_imported_table_run(run)


def _normalize_imported_table_run(run: ET.Element) -> None:
    rpr = run.find("w:rPr", NS)
    if rpr is None:
        rpr = ET.Element(_w("rPr"))
        run.insert(0, rpr)
    for name in ("rStyle", "color", "u", "highlight"):
        for node in list(rpr.findall(f"w:{name}", NS)):
            rpr.remove(node)
    rfonts = rpr.find("w:rFonts", NS)
    if rfonts is None:
        rfonts = ET.Element(_w("rFonts"))
        rpr.insert(0, rfonts)
    rfonts.set(_w("eastAsia"), "宋体")
    rfonts.set(_w("ascii"), "Times New Roman")
    rfonts.set(_w("hAnsi"), "Times New Roman")
    for name in ("sz", "szCs"):
        node = rpr.find(f"w:{name}", NS)
        if node is None:
            node = ET.SubElement(rpr, _w(name))
        node.set(_w("val"), "24")


def _parent_map(root: ET.Element) -> dict[ET.Element, ET.Element]:
    return {child: parent for parent in root.iter() for child in list(parent)}


@dataclass(frozen=True)
class RelationshipImport:
    relationships: int
    parts: dict[str, bytes]


def _import_source_relationships(
    *,
    elements: list[ET.Element],
    source_zip: zipfile.ZipFile,
    target_rels: ET.Element,
    content_types: ET.Element,
    existing_names: set[str],
) -> RelationshipImport:
    if "word/_rels/document.xml.rels" not in source_zip.namelist():
        return RelationshipImport(relationships=0, parts={})
    source_rels = ET.fromstring(source_zip.read("word/_rels/document.xml.rels"))
    source_rel_map = {rel.attrib.get("Id", ""): rel for rel in source_rels.findall(f"{{{REL_NS}}}Relationship")}
    next_rid = _next_relationship_id(target_rels)
    imported_parts: dict[str, bytes] = {}
    imported_relationships = 0
    rel_cache: dict[str, str] = {}
    used_names = set(existing_names)

    for element in elements:
        for node in element.iter():
            for attr_name, attr_value in list(node.attrib.items()):
                if not _is_relationship_attribute(attr_name) or attr_value not in source_rel_map:
                    continue
                if attr_value in rel_cache:
                    node.set(attr_name, rel_cache[attr_value])
                    continue
                source_rel = source_rel_map[attr_value]
                target = source_rel.attrib.get("Target", "")
                rel_type = source_rel.attrib.get("Type", "")
                mode = source_rel.attrib.get("TargetMode")
                new_rid = f"rId{next_rid}"
                next_rid += 1
                new_rel = ET.SubElement(target_rels, f"{{{REL_NS}}}Relationship")
                new_rel.set("Id", new_rid)
                new_rel.set("Type", rel_type)
                if mode == "External":
                    new_rel.set("Target", target)
                    new_rel.set("TargetMode", "External")
                else:
                    source_part = _resolve_word_part(target)
                    new_part = _unique_part_name(source_part, used_names)
                    used_names.add(new_part)
                    imported_parts[new_part] = source_zip.read(source_part)
                    new_rel.set("Target", _relative_word_target(new_part))
                    _ensure_content_type(content_types, new_part)
                rel_cache[attr_value] = new_rid
                node.set(attr_name, new_rid)
                imported_relationships += 1
    return RelationshipImport(relationships=imported_relationships, parts=imported_parts)


def _is_relationship_attribute(name: str) -> bool:
    return name in {f"{{{R_NS}}}id", f"{{{R_NS}}}embed", f"{{{R_NS}}}link"}


def _next_relationship_id(rels: ET.Element) -> int:
    max_seen = 0
    for rel in rels.findall(f"{{{REL_NS}}}Relationship"):
        match = re.match(r"rId(\d+)$", rel.attrib.get("Id", ""))
        if match:
            max_seen = max(max_seen, int(match.group(1)))
    return max_seen + 1


def _resolve_word_part(target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    base = PurePosixPath("word")
    path = base / target
    parts: list[str] = []
    for part in path.parts:
        if part in {"", "."}:
            continue
        if part == "..":
            if parts:
                parts.pop()
        else:
            parts.append(part)
    return "/".join(parts)


def _unique_part_name(source_part: str, used_names: set[str]) -> str:
    source_path = PurePosixPath(source_part)
    suffix = source_path.suffix
    stem = source_path.stem
    parent = source_path.parent
    candidate = str(parent / f"rebuild_{stem}{suffix}")
    idx = 1
    while candidate in used_names:
        candidate = str(parent / f"rebuild_{stem}_{idx}{suffix}")
        idx += 1
    return candidate


def _relative_word_target(part_name: str) -> str:
    if part_name.startswith("word/"):
        return part_name[len("word/"):]
    return part_name


def _ensure_content_type(content_types: ET.Element, part_name: str) -> None:
    ext = PurePosixPath(part_name).suffix.lower().lstrip(".")
    if not ext:
        return
    if any(node.attrib.get("Extension", "").lower() == ext for node in content_types.findall(f"{{{CT_NS}}}Default")):
        return
    content = {
        "png": "image/png",
        "jpg": "image/jpeg",
        "jpeg": "image/jpeg",
        "gif": "image/gif",
        "bmp": "image/bmp",
        "emf": "image/x-emf",
        "wmf": "image/x-wmf",
        "svg": "image/svg+xml",
    }.get(ext, "application/octet-stream")
    node = ET.Element(f"{{{CT_NS}}}Default")
    node.set("Extension", ext)
    node.set("ContentType", content)
    content_types.insert(0, node)


def _first_index(paragraphs: list[tuple[int, str]], predicate) -> int | None:
    for idx, text in paragraphs:
        if predicate(text):
            return idx
    return None


def _compact(text: str) -> str:
    return re.sub(r"\s+", "", text)


def _element_text_with_fields(element: ET.Element) -> str:
    parts = []
    for node in element.iter():
        if node.tag in {_w("t"), _w("instrText")} and node.text:
            parts.append(node.text)
    return "".join(parts)


def _is_abstract_heading(text: str) -> bool:
    compact = _compact(text)
    return compact in {"摘要", "摘 要", "ABSTRACT"}


def _is_toc_heading(text: str) -> bool:
    compact = _compact(text)
    return compact in {"目录", "目錄"} or compact.startswith("目录TOC") or compact.startswith("目錄TOC")


def _is_first_body_heading(text: str) -> bool:
    return re.fullmatch(r"1\s*绪论", text.strip()) is not None


def _set_run_font(run: ET.Element, *, east_asia: str, ascii_font: str, size: str) -> None:
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
    for name in ("sz", "szCs"):
        node = rpr.find(f"w:{name}", NS)
        if node is None:
            node = ET.SubElement(rpr, _w(name))
        node.set(_w("val"), size)


def _empty_rels_xml() -> bytes:
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{REL_NS}"/>'.encode()


def _empty_content_types_xml() -> bytes:
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="{CT_NS}"/>'.encode()


def _serialize_xml(root: ET.Element) -> bytes:
    return serialize_xml(root)


def _serialize_package_xml(root: ET.Element, namespace: str) -> bytes:
    return serialize_package_xml(root, namespace)


def _w(local: str) -> str:
    return f"{{{W_NS}}}{local}"
