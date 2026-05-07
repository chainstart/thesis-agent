from __future__ import annotations

import copy
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET

from .docx_inspect import NS, W_NS, _paragraph_text
from .ooxml import serialize_xml
from .pdf_inspect import extract_page_texts, pdf_info
from .tools import Toolchain
from .visual_check import _infer_page_number_label, _parse_toc_entry_line, _rendered_heading_label_map


ET.register_namespace("w", W_NS)


@dataclass(frozen=True)
class TocSyncReport:
    input: Path
    output: Path
    updated_entries: int = 0
    inserted_entries: int = 0
    removed_entries: int = 0
    formatted_entries: int = 0
    missing_entries: list[str] = field(default_factory=list)


def sync_static_toc_from_pdf(input_path: Path, pdf_path: Path, output_path: Path, toolchain: Toolchain) -> TocSyncReport:
    if input_path.suffix.lower() != ".docx":
        raise ValueError("TOC sync currently supports .docx targets only")
    if not zipfile.is_zipfile(input_path):
        raise ValueError(f"Not a valid docx file: {input_path}")

    info = pdf_info(pdf_path, toolchain)
    page_texts = extract_page_texts(pdf_path, toolchain, info.pages)
    page_labels = {
        page: label
        for page, text in page_texts.items()
        if (label := _infer_page_number_label(text)) is not None
    }
    heading_labels = _rendered_heading_label_map(page_texts, page_labels)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    updated = 0
    inserted = 0
    removed = 0
    formatted = 0
    seen: set[str] = set()
    missing: list[str] = []
    with zipfile.ZipFile(input_path) as zin:
        root = ET.fromstring(zin.read("word/document.xml"))
        paragraphs = list(root.iter(_w("p")))
        first_body_pos = _first_body_paragraph_pos(paragraphs)
        parent_map = {child: parent for parent in root.iter() for child in list(parent)}
        toc_paragraphs: list[tuple[ET.Element, str, str]] = []

        for pos, paragraph in enumerate(paragraphs):
            if first_body_pos is not None and pos >= first_body_pos:
                break
            parsed = _parse_docx_toc_entry(_paragraph_text(paragraph))
            if parsed is None:
                continue
            title, old_label = parsed
            if _toc_entry_level(title) > 2:
                parent = parent_map.get(paragraph)
                if parent is not None:
                    parent.remove(paragraph)
                    removed += 1
                continue
            key = _heading_key(title)
            if key not in heading_labels:
                continue
            seen.add(key)
            toc_paragraphs.append((paragraph, key, title))
            new_label = heading_labels[key]
            _set_toc_entry_text(paragraph, title, new_label)
            formatted += 1
            if old_label != new_label:
                updated += 1

        if "致谢" in heading_labels and "致谢" not in seen:
            anchor = _toc_anchor_for_ack(toc_paragraphs)
            if anchor is not None:
                anchor_p, _, _ = anchor
                parent = parent_map.get(anchor_p)
                if parent is not None:
                    new_p = copy.deepcopy(anchor_p)
                    _set_toc_entry_text(new_p, "致谢", heading_labels["致谢"])
                    siblings = list(parent)
                    parent.insert(siblings.index(anchor_p) + 1, new_p)
                    inserted += 1
                    formatted += 1
            else:
                missing.append("致谢")

        document_xml = _serialize_xml(root)
        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = document_xml if item.filename == "word/document.xml" else zin.read(item.filename)
                zout.writestr(item, data)

    return TocSyncReport(
        input=input_path,
        output=output_path,
        updated_entries=updated,
        inserted_entries=inserted,
        removed_entries=removed,
        formatted_entries=formatted,
        missing_entries=missing,
    )


def _first_body_paragraph_pos(paragraphs: list[ET.Element]) -> int | None:
    for pos, paragraph in enumerate(paragraphs):
        if re.fullmatch(r"1\s+绪论", _paragraph_text(paragraph).strip()):
            return pos
    return None


def _parse_docx_toc_entry(text: str) -> tuple[str, str] | None:
    parsed = _parse_toc_entry_line(text)
    if parsed is not None:
        return parsed
    stripped = text.strip()
    if not re.match(r"^([1-9](?:\.\d+)*\s*|参考文献|致\s*谢|附录)", stripped):
        return None
    if "\t" not in stripped and "..." not in stripped and "…" not in stripped:
        return None
    match = re.match(r"^(?P<title>.+?)(?P<label>\d+)\s*$", stripped)
    if not match:
        return None
    return match.group("title").strip(" .\t…"), match.group("label")


def _toc_anchor_for_ack(entries: list[tuple[ET.Element, str, str]]) -> tuple[ET.Element, str, str] | None:
    for entry in reversed(entries):
        if entry[1] == "参考文献":
            return entry
    return entries[-1] if entries else None


def _set_toc_entry_text(paragraph: ET.Element, title: str, label: str) -> None:
    title = _normalize_toc_title_spacing(title)
    ppr = paragraph.find("w:pPr", NS)
    preserved = copy.deepcopy(ppr) if ppr is not None else None
    for child in list(paragraph):
        paragraph.remove(child)
    if preserved is not None:
        paragraph.append(preserved)
    level = _toc_entry_level(title)
    _format_toc_entry_paragraph(paragraph, level)
    run = ET.SubElement(paragraph, _w("r"))
    _set_run_font(run, east_asia="黑体", ascii_font="Times New Roman", size="28" if level == 1 else "24", bold=False)
    t = ET.SubElement(run, _w("t"))
    t.text = title
    if re.search(r"\s", title):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    ET.SubElement(run, _w("tab"))
    page = ET.SubElement(run, _w("t"))
    page.text = label


def _toc_entry_level(title: str) -> int:
    stripped = title.strip()
    if re.match(r"^[1-9]\.\d+\.\d+", stripped):
        return 3
    if re.match(r"^[1-9]\.\d+", stripped):
        return 2
    return 1


def _format_toc_entry_paragraph(paragraph: ET.Element, level: int) -> None:
    ppr = paragraph.find("w:pPr", NS)
    if ppr is None:
        ppr = ET.Element(_w("pPr"))
        paragraph.insert(0, ppr)
    spacing = ppr.find("w:spacing", NS)
    if spacing is None:
        spacing = ET.Element(_w("spacing"))
        ppr.append(spacing)
    spacing.set(_w("before"), "0")
    spacing.set(_w("after"), "0")
    spacing.set(_w("line"), "300")
    spacing.set(_w("lineRule"), "auto")

    ind = ppr.find("w:ind", NS)
    if ind is None:
        ind = ET.Element(_w("ind"))
        ppr.append(ind)
    if level == 2:
        ind.set(_w("left"), "420")
    else:
        ind.set(_w("left"), "0")
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


def _heading_key(title: str) -> str:
    title = re.sub(r"(?:\.{2,}|…{2,}|\t+|\s{2,})\d+\s*$", "", title).strip()
    return re.sub(r"\s+", "", title)


def _normalize_toc_title_spacing(title: str) -> str:
    stripped = title.strip()
    match = re.match(r"^([1-9])\s*(.+)$", stripped)
    if match and not stripped.startswith(match.group(1) + "."):
        return f"{match.group(1)} {match.group(2).strip()}"
    match = re.match(r"^([1-9](?:\.\d+)+)\s*(.+)$", stripped)
    if match:
        return f"{match.group(1)} {match.group(2).strip()}"
    if re.sub(r"\s+", "", stripped) == "致谢":
        return "致谢"
    return stripped


def _serialize_xml(root: ET.Element) -> bytes:
    return serialize_xml(root)


def _w(local: str) -> str:
    return f"{{{W_NS}}}{local}"
