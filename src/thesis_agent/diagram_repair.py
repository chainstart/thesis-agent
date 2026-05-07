from __future__ import annotations

import io
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET

from .docx_inspect import NS, W_NS, _paragraph_text
from .ooxml import serialize_xml


R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"

ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)
ET.register_namespace("a", A_NS)
ET.register_namespace("pic", PIC_NS)


@dataclass(frozen=True)
class DiagramRepairReport:
    input: Path
    output: Path
    repaired_diagrams: int = 0
    repaired_parts: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def repair_known_diagrams(input_path: Path, output_path: Path) -> DiagramRepairReport:
    if input_path.suffix.lower() != ".docx":
        raise ValueError("diagram repair currently supports .docx targets only")
    if not zipfile.is_zipfile(input_path):
        raise ValueError(f"Not a valid docx file: {input_path}")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    warnings: list[str] = []
    repaired_parts: list[str] = []
    try:
        diagram_png = _system_architecture_png()
    except RuntimeError as exc:
        warnings.append(str(exc))
        diagram_png = None

    with zipfile.ZipFile(input_path) as zin:
        root = ET.fromstring(zin.read("word/document.xml"))
        rels_xml = zin.read("word/_rels/document.xml.rels") if "word/_rels/document.xml.rels" in zin.namelist() else None
        rel_targets = _relationship_targets(rels_xml) if rels_xml else {}
        body = root.find("w:body", NS)
        if body is None:
            raise ValueError("word/document.xml does not contain w:body")

        if diagram_png:
            children = list(body)
            for idx, child in enumerate(children):
                if child.tag != _w("p"):
                    continue
                caption = re.sub(r"\s+", "", _paragraph_text(child))
                if "系统架构流程图" not in caption:
                    continue
                previous = _previous_paragraph_with_drawing(children, idx)
                if previous is None:
                    continue
                rid = _first_embedded_image_id(previous)
                target = rel_targets.get(rid or "")
                if not target:
                    continue
                part = _word_part_name(target)
                _remove_picture_cropping(previous)
                repaired_parts.append(part)

        document_xml = _serialize_xml(root)
        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    data = document_xml
                elif diagram_png and item.filename in repaired_parts:
                    data = diagram_png
                else:
                    data = zin.read(item.filename)
                zout.writestr(item, data)

    return DiagramRepairReport(
        input=input_path,
        output=output_path,
        repaired_diagrams=len(repaired_parts),
        repaired_parts=repaired_parts,
        warnings=warnings,
    )


def _system_architecture_png() -> bytes:
    try:
        from PIL import Image, ImageDraw, ImageFont
    except ImportError as exc:
        raise RuntimeError("Pillow is required for diagram repair") from exc

    width, height = 1400, 800
    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)
    font = _font(28)
    small = _font(23)
    title_font = _font(30)
    line = (0, 0, 0)

    boxes = {
        "cloud": (500, 80, 900, 190),
        "wifi": (560, 270, 840, 365),
        "stm32": (500, 455, 900, 565),
        "servo": (1060, 455, 1280, 565),
        "temp": (120, 645, 340, 740),
        "gas": (395, 645, 615, 740),
        "rfid": (670, 645, 890, 740),
        "alarm": (945, 645, 1165, 740),
    }
    labels = {
        "cloud": "云平台\n（上位机）",
        "wifi": "WiFi模块",
        "stm32": "STM32 主控板",
        "servo": "舵机",
        "temp": "温湿度传感器",
        "gas": "气体传感器",
        "rfid": "RFID模块",
        "alarm": "蜂鸣器",
    }

    _center_text(draw, (width // 2, 35), "物联网系统架构流程图", title_font)
    for key, box in boxes.items():
        draw.rounded_rectangle(box, radius=2, outline=line, width=3, fill="white")
        _center_multiline(draw, box, labels[key], font)

    _arrow(draw, (640, 270), (640, 190), width=4)
    _arrow(draw, (760, 190), (760, 270), width=4)
    draw.text((590, 230), "上传数据", font=small, fill=line, anchor="mm")
    draw.text((815, 230), "下发指令", font=small, fill=line, anchor="mm")

    _arrow(draw, (700, 365), (700, 455), width=4)
    _arrow(draw, (900, 510), (1060, 510), width=4)

    _arrow(draw, (560, 565), (230, 645), width=4)
    _arrow(draw, (650, 565), (505, 645), width=4)
    _arrow(draw, (750, 565), (780, 645), width=4)
    _arrow(draw, (840, 565), (1055, 645), width=4)

    buffer = io.BytesIO()
    image.save(buffer, format="PNG", optimize=True)
    return buffer.getvalue()


def _font(size: int):
    try:
        from PIL import ImageFont
    except ImportError as exc:
        raise RuntimeError("Pillow is required for diagram repair") from exc
    candidates = [
        "/mnt/c/Windows/Fonts/simsun.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    for path in candidates:
        if Path(path).exists():
            return ImageFont.truetype(path, size)
    return ImageFont.load_default()


def _center_text(draw, center: tuple[int, int], text: str, font) -> None:
    draw.text(center, text, font=font, fill=(0, 0, 0), anchor="mm")


def _center_multiline(draw, box: tuple[int, int, int, int], text: str, font) -> None:
    lines = text.splitlines()
    line_heights = [draw.textbbox((0, 0), line, font=font)[3] for line in lines]
    total = sum(line_heights) + max(0, len(lines) - 1) * 8
    x = (box[0] + box[2]) // 2
    y = (box[1] + box[3] - total) // 2
    for line, line_height in zip(lines, line_heights):
        draw.text((x, y), line, font=font, fill=(0, 0, 0), anchor="ma")
        y += line_height + 8


def _arrow(draw, start: tuple[int, int], end: tuple[int, int], width: int = 3) -> None:
    draw.line([start, end], fill=(0, 0, 0), width=width)
    sx, sy = start
    ex, ey = end
    dx, dy = ex - sx, ey - sy
    length = max((dx * dx + dy * dy) ** 0.5, 1)
    ux, uy = dx / length, dy / length
    size = 16
    left = (ex - ux * size - uy * size * 0.55, ey - uy * size + ux * size * 0.55)
    right = (ex - ux * size + uy * size * 0.55, ey - uy * size - ux * size * 0.55)
    draw.polygon([end, left, right], fill=(0, 0, 0))


def _relationship_targets(rels_xml: bytes) -> dict[str, str]:
    root = ET.fromstring(rels_xml)
    result: dict[str, str] = {}
    for rel in root.findall(f"{{{REL_NS}}}Relationship"):
        result[rel.attrib.get("Id", "")] = rel.attrib.get("Target", "")
    return result


def _previous_paragraph_with_drawing(children: list[ET.Element], start_idx: int) -> ET.Element | None:
    for child in reversed(children[:start_idx]):
        if child.tag != _w("p"):
            continue
        if child.find(".//w:drawing", NS) is not None or child.find(".//w:pict", NS) is not None:
            return child
        if _paragraph_text(child).strip():
            return None
    return None


def _first_embedded_image_id(paragraph: ET.Element) -> str | None:
    blip = paragraph.find(f".//{{{A_NS}}}blip")
    if blip is None:
        return None
    return blip.attrib.get(f"{{{R_NS}}}embed")


def _remove_picture_cropping(paragraph: ET.Element) -> None:
    for blip_fill in paragraph.findall(f".//{{{PIC_NS}}}blipFill"):
        for src_rect in list(blip_fill.findall(f"{{{A_NS}}}srcRect")):
            blip_fill.remove(src_rect)


def _word_part_name(target: str) -> str:
    normalized = target.lstrip("/")
    if normalized.startswith("word/"):
        return normalized
    return f"word/{normalized}"


def _serialize_xml(root: ET.Element) -> bytes:
    return serialize_xml(root)


def _w(local: str) -> str:
    return f"{{{W_NS}}}{local}"
