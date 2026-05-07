from __future__ import annotations

import copy
import json
import zipfile
from dataclasses import asdict, dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET

from .docx_inspect import NS, W_NS, _paragraph_text
from .rebuild import (
    CT_NS,
    REL_NS,
    _empty_content_types_xml,
    _empty_rels_xml,
    _import_source_relationships,
    _resolve_template_docx,
    _serialize_package_xml,
    _serialize_xml,
    _w,
)


@dataclass(frozen=True)
class StandardTemplateReport:
    cover: Path
    body: Path
    output: Path
    cover_elements: int = 0
    body_elements: int = 0
    imported_relationships: int = 0
    imported_parts: list[str] = field(default_factory=list)
    merged_styles: int = 0
    warnings: list[str] = field(default_factory=list)

    def to_json(self) -> str:
        return json.dumps(asdict(self), ensure_ascii=False, indent=2, default=str)


def rebuild_standard_template(cover_template: Path, body_template: Path, output_path: Path) -> StandardTemplateReport:
    """Build the canonical thesis template from the official cover and body templates.

    This is intentionally a package-level composition step: the cover page and
    body-format pages keep their original WordprocessingML formatting. It is
    used as the lossless baseline before any student content is inserted.
    """
    cover_docx = _resolve_template_docx(cover_template)
    body_docx = _resolve_template_docx(body_template)
    if cover_docx is None:
        raise ValueError(f"No usable .docx cover template found for {cover_template}")
    if body_docx is None:
        raise ValueError(f"No usable .docx body template found for {body_template}")
    if not zipfile.is_zipfile(cover_docx):
        raise ValueError(f"Not a valid docx file: {cover_docx}")
    if not zipfile.is_zipfile(body_docx):
        raise ValueError(f"Not a valid docx file: {body_docx}")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    warnings: list[str] = []
    with zipfile.ZipFile(cover_docx) as cover_zip, zipfile.ZipFile(body_docx) as body_zip:
        cover_root = ET.fromstring(cover_zip.read("word/document.xml"))
        body_root = ET.fromstring(body_zip.read("word/document.xml"))
        cover_body = cover_root.find("w:body", NS)
        body_body = body_root.find("w:body", NS)
        if cover_body is None:
            raise ValueError(f"Cover template has no w:body: {cover_docx}")
        if body_body is None:
            raise ValueError(f"Body template has no w:body: {body_docx}")

        cover_elements = [copy.deepcopy(el) for el in list(cover_body) if el.tag != _w("sectPr")]
        cover_section = _extract_final_section(cover_body)
        body_elements = [copy.deepcopy(el) for el in list(body_body)]
        if cover_section is None:
            warnings.append("封面模板缺少最终节属性，已使用正文模板前的默认分页衔接。")
            cover_elements.append(_page_break_paragraph())
        else:
            _attach_section_to_last_paragraph(cover_elements, cover_section)
        _compensate_first_body_page_after_cover(body_elements)
        _normalize_front_matter_roman_numbering(body_elements)

        target_rels_xml = (
            body_zip.read("word/_rels/document.xml.rels")
            if "word/_rels/document.xml.rels" in body_zip.namelist()
            else _empty_rels_xml()
        )
        target_rels = ET.fromstring(target_rels_xml)
        content_types_xml = (
            body_zip.read("[Content_Types].xml")
            if "[Content_Types].xml" in body_zip.namelist()
            else _empty_content_types_xml()
        )
        content_types = ET.fromstring(content_types_xml)
        import_result = _import_source_relationships(
            elements=cover_elements,
            source_zip=cover_zip,
            target_rels=target_rels,
            content_types=content_types,
            existing_names=set(body_zip.namelist()),
        )

        styles_data = body_zip.read("word/styles.xml") if "word/styles.xml" in body_zip.namelist() else None
        merged_styles = 0
        if styles_data and "word/styles.xml" in cover_zip.namelist():
            styles_root = ET.fromstring(styles_data)
            cover_styles_root = ET.fromstring(cover_zip.read("word/styles.xml"))
            merged_styles = _merge_missing_styles(styles_root, cover_styles_root)
            styles_data = _serialize_xml(styles_root)

        for child in list(body_body):
            body_body.remove(child)
        for element in cover_elements:
            body_body.append(element)
        for element in body_elements:
            body_body.append(element)

        document_xml = _serialize_xml(body_root)
        rels_xml = _serialize_package_xml(target_rels, REL_NS)
        content_types_data = _serialize_package_xml(content_types, CT_NS)

        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as out_zip:
            written: set[str] = set()
            for item in body_zip.infolist():
                if item.filename == "word/document.xml":
                    data = document_xml
                elif item.filename == "word/_rels/document.xml.rels":
                    data = rels_xml
                elif item.filename == "[Content_Types].xml":
                    data = content_types_data
                elif item.filename == "word/styles.xml" and styles_data is not None:
                    data = styles_data
                else:
                    data = body_zip.read(item.filename)
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

    return StandardTemplateReport(
        cover=cover_docx,
        body=body_docx,
        output=output_path,
        cover_elements=len(cover_elements),
        body_elements=len(body_elements),
        imported_relationships=import_result.relationships,
        imported_parts=sorted(import_result.parts),
        merged_styles=merged_styles,
        warnings=warnings,
    )


def _extract_final_section(body: ET.Element) -> ET.Element | None:
    direct = body.find("w:sectPr", NS)
    if direct is not None:
        return copy.deepcopy(direct)
    for child in reversed(list(body)):
        if child.tag != _w("p"):
            continue
        sect = child.find("w:pPr/w:sectPr", NS)
        if sect is not None:
            return copy.deepcopy(sect)
    return None


def _attach_section_to_last_paragraph(elements: list[ET.Element], section: ET.Element) -> None:
    for element in reversed(elements):
        if element.tag != _w("p"):
            continue
        ppr = element.find("w:pPr", NS)
        if ppr is None:
            ppr = ET.Element(_w("pPr"))
            element.insert(0, ppr)
        for existing in list(ppr.findall("w:sectPr", NS)):
            ppr.remove(existing)
        ppr.append(copy.deepcopy(section))
        return
    elements.append(_page_break_paragraph())


def _compensate_first_body_page_after_cover(elements: list[ET.Element]) -> None:
    """Keep the body template's first page visually stable after a cover section.

    LibreOffice lays out the first page after an inserted section boundary about
    4pt lower than the same body template rendered as a standalone document.
    The official body template starts with empty spacer paragraphs, so the
    compensation is confined to the first spacer and does not touch visible
    declaration text or later thesis styles.
    """
    for element in elements[:3]:
        if element.tag != _w("p") or _paragraph_text(element).strip():
            continue
        ppr = element.find("w:pPr", NS)
        if ppr is None:
            ppr = ET.Element(_w("pPr"))
            element.insert(0, ppr)
        spacing = ppr.find("w:spacing", NS)
        if spacing is None:
            spacing = ET.Element(_w("spacing"))
            ppr.insert(0, spacing)
        spacing.set(_w("line"), "420")
        spacing.set(_w("lineRule"), "auto")
        spacing.set(_w("before"), "1")
        return


def _normalize_front_matter_roman_numbering(elements: list[ET.Element]) -> None:
    first_roman_section = True
    for element in elements:
        if element.tag == _w("p") and _paragraph_text(element).strip().replace(" ", "") == "1绪论":
            break
        for section in element.findall(".//w:sectPr", NS):
            page_number = section.find("w:pgNumType", NS)
            if page_number is None or page_number.attrib.get(_w("fmt")) != "upperRoman":
                continue
            if first_roman_section:
                page_number.set(_w("start"), "1")
                first_roman_section = False
            else:
                page_number.attrib.pop(_w("start"), None)


def _page_break_paragraph() -> ET.Element:
    paragraph = ET.Element(_w("p"))
    run = ET.SubElement(paragraph, _w("r"))
    br = ET.SubElement(run, _w("br"))
    br.set(_w("type"), "page")
    return paragraph


def _merge_missing_styles(target_styles: ET.Element, source_styles: ET.Element) -> int:
    existing = {
        style.attrib.get(_w("styleId"))
        for style in target_styles.findall("w:style", NS)
        if style.attrib.get(_w("styleId"))
    }
    merged = 0
    for style in source_styles.findall("w:style", NS):
        style_id = style.attrib.get(_w("styleId"))
        if not style_id or style_id in existing:
            continue
        target_styles.append(copy.deepcopy(style))
        existing.add(style_id)
        merged += 1
    return merged


def template_text_digest(docx_path: Path) -> str:
    with zipfile.ZipFile(docx_path) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    return "\n".join(
        text
        for paragraph in root.iter(_w("p"))
        if (text := _paragraph_text(paragraph).strip())
    )
