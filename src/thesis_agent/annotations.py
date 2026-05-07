from __future__ import annotations

import json
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from xml.etree import ElementTree as ET

from .docx_inspect import NS, W_NS, _paragraph_text
from .ooxml import serialize_xml


V_NS = "urn:schemas-microsoft-com:vml"


@dataclass(frozen=True)
class AnnotationStripReport:
    input: Path
    output: Path
    red_shapes_removed: int = 0
    red_runs_removed: int = 0
    empty_pictures_removed: int = 0
    red_color_properties_removed: int = 0

    def to_json(self) -> str:
        return json.dumps(asdict(self), ensure_ascii=False, indent=2, default=str)


def strip_red_annotations_from_docx(input_path: Path, output_path: Path) -> AnnotationStripReport:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(input_path) as zin:
        root = ET.fromstring(zin.read("word/document.xml"))
        removed_shapes = _remove_red_shapes(root)
        removed_runs = _remove_red_runs(root)
        removed_pictures = _remove_empty_pictures(root)
        removed_colors = _remove_red_color_properties(root)
        document_xml = serialize_xml(root)
        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = document_xml if item.filename == "word/document.xml" else zin.read(item.filename)
                zout.writestr(item, data)
    return AnnotationStripReport(
        input=input_path,
        output=output_path,
        red_shapes_removed=removed_shapes,
        red_runs_removed=removed_runs,
        empty_pictures_removed=removed_pictures,
        red_color_properties_removed=removed_colors,
    )


def _remove_red_shapes(root: ET.Element) -> int:
    parents = _parent_map(root)
    removed = 0
    for shape in list(root.findall(f".//{{{V_NS}}}shape")):
        if not _is_red_shape(shape):
            continue
        parent = parents.get(shape)
        if parent is not None:
            parent.remove(shape)
            removed += 1
    return removed


def _remove_red_runs(root: ET.Element) -> int:
    parents = _parent_map(root)
    removed = 0
    for run in list(root.findall(".//w:r", NS)):
        if run.find(".//w:pict", NS) is not None or run.find(".//w:drawing", NS) is not None:
            continue
        if not _is_red_run(run):
            continue
        parent = parents.get(run)
        if parent is not None:
            parent.remove(run)
            removed += 1
    return removed


def _remove_empty_pictures(root: ET.Element) -> int:
    parents = _parent_map(root)
    removed = 0
    for pict in list(root.findall(".//w:pict", NS)):
        if list(pict):
            continue
        parent = parents.get(pict)
        if parent is not None:
            parent.remove(pict)
            removed += 1
    return removed


def _remove_red_color_properties(root: ET.Element) -> int:
    parents = _parent_map(root)
    removed = 0
    for color in list(root.findall(".//w:color", NS)):
        if not _is_red_value(color.attrib.get(_w("val"), "")):
            continue
        parent = parents.get(color)
        if parent is not None:
            parent.remove(color)
            removed += 1
    for highlight in list(root.findall(".//w:highlight", NS)):
        if not _is_red_value(highlight.attrib.get(_w("val"), "")):
            continue
        parent = parents.get(highlight)
        if parent is not None:
            parent.remove(highlight)
            removed += 1
    return removed


def _is_red_shape(shape: ET.Element) -> bool:
    for attr in ("strokecolor", "fillcolor"):
        if _is_red_value(shape.attrib.get(attr, "")):
            return True
    style = shape.attrib.get("style", "")
    if "red" in style.lower() or "#ff0000" in style.lower():
        return True
    return any(_is_red_run(run) for run in shape.findall(".//w:r", NS))


def _is_red_run(run: ET.Element) -> bool:
    color = run.find("w:rPr/w:color", NS)
    if color is not None and _is_red_value(color.attrib.get(_w("val"), "")):
        return True
    highlight = run.find("w:rPr/w:highlight", NS)
    return highlight is not None and _is_red_value(highlight.attrib.get(_w("val"), ""))


def _is_red_value(value: str) -> bool:
    normalized = value.strip().lower().lstrip("#")
    return normalized in {"red", "ff0000", "f00"}


def _parent_map(root: ET.Element) -> dict[ET.Element, ET.Element]:
    return {child: parent for parent in root.iter() for child in list(parent)}


def _w(local: str) -> str:
    return f"{{{W_NS}}}{local}"
