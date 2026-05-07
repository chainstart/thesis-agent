from __future__ import annotations

import json
import re
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from xml.etree import ElementTree as ET

from .docx_inspect import NS, W_NS
from .format_fix import _resolve_template_docx


@dataclass(frozen=True)
class TemplateChecklistItem:
    id: str
    paragraph_index: int
    text: str


def extract_red_text_checklist(template_path: Path) -> list[TemplateChecklistItem]:
    docx = _resolve_template_docx(template_path)
    if docx is None:
        return []
    items: list[TemplateChecklistItem] = []
    with zipfile.ZipFile(docx) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    for idx, paragraph in enumerate(root.iter(f"{{{W_NS}}}p")):
        red_text = _red_text_in_paragraph(paragraph)
        if not red_text:
            continue
        text = re.sub(r"\s+", " ", red_text).strip()
        if text:
            items.append(TemplateChecklistItem(id=f"R{len(items) + 1:03d}", paragraph_index=idx, text=text))
    return items


def write_red_text_checklist(template_path: Path, markdown_path: Path, json_path: Path | None = None) -> list[TemplateChecklistItem]:
    items = extract_red_text_checklist(template_path)
    markdown_path.parent.mkdir(parents=True, exist_ok=True)
    lines = [
        "# Template Red-Text Checklist",
        "",
        f"- Source: `{template_path}`",
        f"- Items: {len(items)}",
        "",
    ]
    for item in items:
        lines.append(f"- [ ] `{item.id}` p{item.paragraph_index}: {item.text}")
    markdown_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    if json_path is not None:
        json_path.write_text(json.dumps([asdict(item) for item in items], ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return items


def _red_text_in_paragraph(paragraph: ET.Element) -> str:
    pieces: list[str] = []
    for run in paragraph.findall("w:r", NS):
        color = run.find("w:rPr/w:color", NS)
        value = color.attrib.get(f"{{{W_NS}}}val", "") if color is not None else ""
        if value.upper() not in {"FF0000", "RED", "C00000", "E60000"}:
            continue
        pieces.extend(t.text or "" for t in run.findall("w:t", NS))
        if run.find("w:tab", NS) is not None:
            pieces.append("\t")
    return "".join(pieces)
