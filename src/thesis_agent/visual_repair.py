from __future__ import annotations

import io
import json
import posixpath
import re
import zipfile
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from .docx_inspect import NS, W_NS, _paragraph_text
from .ooxml import serialize_xml


R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)
ET.register_namespace("a", A_NS)
ET.register_namespace("wp", WP_NS)


@dataclass(frozen=True)
class VisualRepairReport:
    input: Path
    output: Path
    enhanced_figures: int = 0
    extracted_code_formula_figures: int = 0
    renumbered_figures: int = 0
    updated_references: int = 0
    actions: list[str] = field(default_factory=list)
    skipped: list[str] = field(default_factory=list)

    def to_json(self) -> str:
        return json.dumps(asdict(self), ensure_ascii=False, indent=2, default=str)


def repair_visual_items(input_path: Path, output_path: Path, audit_result: dict[str, Any]) -> VisualRepairReport:
    if input_path.suffix.lower() != ".docx":
        raise ValueError("visual repair currently supports .docx targets only")
    if not zipfile.is_zipfile(input_path):
        raise ValueError(f"Not a valid docx file: {input_path}")

    visual = audit_result["target_visual"]
    small_caption_keys = {_caption_key(item) for item in getattr(visual, "figure_readability_warnings", [])}
    small_caption_keys.discard("")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    actions: list[str] = []
    skipped: list[str] = []
    media_updates: dict[str, bytes] = {}
    enhanced = 0
    extracted = 0

    with zipfile.ZipFile(input_path) as zin:
        root = ET.fromstring(zin.read("word/document.xml"))
        rels_xml = zin.read("word/_rels/document.xml.rels") if "word/_rels/document.xml.rels" in zin.namelist() else None
        rel_targets = _relationship_targets(rels_xml) if rels_xml else {}
        body = root.find("w:body", NS)
        if body is None:
            raise ValueError("word/document.xml does not contain w:body")

        children = list(body)
        code_formula_targets: list[tuple[ET.Element, ET.Element | None, str, str, str]] = []
        for idx, child in enumerate(children):
            if child.tag != _w("p"):
                continue
            caption = _paragraph_text(child).strip()
            if not _is_figure_caption(caption) or not _caption_indicates_code_or_formula(caption):
                continue
            previous = _previous_visual_paragraph(children, idx)
            label = _figure_label(caption)
            replacement_kind = "formula" if "公式" in caption else "code"
            code_formula_targets.append((child, previous, caption, label, replacement_kind))

        for idx, child in enumerate(children):
            if child.tag != _w("p"):
                continue
            caption = _paragraph_text(child).strip()
            if not _is_figure_caption(caption) or _caption_indicates_code_or_formula(caption):
                continue
            redrawn = _redrawn_known_figure(caption)
            if _normalize_caption_for_match(caption) not in small_caption_keys and redrawn is None:
                continue
            previous = _previous_visual_paragraph(children, idx)
            if previous is None:
                skipped.append(f"{caption}: 未找到可替换图片")
                continue
            rid = _first_embedded_image_id(previous)
            part = _word_part_name(rel_targets.get(rid or "", ""))
            if not part or part not in zin.namelist():
                skipped.append(f"{caption}: 未找到图片资源")
                continue
            if redrawn is None and _skip_generic_enhancement(caption):
                skipped.append(f"{caption}: 疑似实物照片或传感器照片，未自动放大")
                continue
            repaired = redrawn or _enhanced_image_bytes(zin.read(part))
            if repaired is None:
                skipped.append(f"{caption}: 图片格式不支持自动增强")
                continue
            media_updates[part] = repaired
            _expand_drawing_extent(previous, repaired)
            enhanced += 1
            if redrawn is not None:
                actions.append(f"{caption}: 已按原图内容重绘为可读图")
            else:
                actions.append(f"{caption}: 已裁剪白边、增强清晰度并放大显示")

        removed_label_replacements: dict[str, str] = {}
        for caption_paragraph, visual_paragraph, caption, label, replacement_kind in code_formula_targets:
            if visual_paragraph is not None and visual_paragraph in list(body):
                body.remove(visual_paragraph)
            if caption_paragraph not in list(body):
                continue
            insert_at = list(body).index(caption_paragraph)
            body.remove(caption_paragraph)
            replacement_paragraphs = _code_formula_replacement_paragraphs(caption, replacement_kind, label)
            for offset, paragraph in enumerate(replacement_paragraphs):
                body.insert(insert_at + offset, paragraph)
            removed_label_replacements[_compact_label(label)] = "如下公式" if replacement_kind == "formula" else "如下代码"
            extracted += 1
            actions.append(f"{caption}: 已替换为正文{('公式' if replacement_kind == 'formula' else '代码')}块")

        renumber_map, renumbered = _renumber_figure_captions(body)
        updated_refs = _update_figure_references(body, removed_label_replacements, renumber_map)

        document_xml = serialize_xml(root)
        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    data = document_xml
                elif item.filename in media_updates:
                    data = media_updates[item.filename]
                else:
                    data = zin.read(item.filename)
                zout.writestr(item, data)

    return VisualRepairReport(
        input=input_path,
        output=output_path,
        enhanced_figures=enhanced,
        extracted_code_formula_figures=extracted,
        renumbered_figures=renumbered,
        updated_references=updated_refs,
        actions=actions,
        skipped=skipped,
    )


def _caption_key(warning: str) -> str:
    match = re.search(r"page\s+\d+:\s*(图\s*\d+\s*[-－]\s*\d+.*?)(?:\s+上方|$)", warning)
    return _normalize_caption_for_match(match.group(1)) if match else ""


def _normalize_caption_for_match(text: str) -> str:
    return re.sub(r"\s+", "", text).replace("－", "-")


def _is_figure_caption(text: str) -> bool:
    return re.match(r"^图\s*\d+\s*[-－]\s*\d+", text.strip()) is not None


def _caption_indicates_code_or_formula(text: str) -> bool:
    return re.search(r"(代码|源码|公式)", re.sub(r"\s+", "", text)) is not None


def _figure_label(text: str) -> str:
    match = re.match(r"^(图)\s*(\d+)\s*[-－]\s*(\d+)", text.strip())
    return f"图{match.group(2)}-{match.group(3)}" if match else ""


def _previous_visual_paragraph(children: list[ET.Element], idx: int) -> ET.Element | None:
    for previous in reversed(children[:idx]):
        if previous.tag != _w("p"):
            continue
        if previous.find(".//w:drawing", NS) is not None or previous.find(".//w:pict", NS) is not None:
            return previous
        if _paragraph_text(previous).strip():
            return None
    return None


def _relationship_targets(rels_xml: bytes | None) -> dict[str, str]:
    if not rels_xml:
        return {}
    root = ET.fromstring(rels_xml)
    result: dict[str, str] = {}
    for rel in root.findall(f"{{{REL_NS}}}Relationship"):
        rid = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        if rid and target:
            result[rid] = target
    return result


def _word_part_name(target: str) -> str:
    if not target:
        return ""
    if target.startswith("/"):
        return target.lstrip("/")
    if target.startswith("word/"):
        return target
    return posixpath.normpath(posixpath.join("word", target))


def _first_embedded_image_id(paragraph: ET.Element) -> str | None:
    blip = paragraph.find(f".//{{{A_NS}}}blip")
    if blip is None:
        return None
    return blip.attrib.get(f"{{{R_NS}}}embed") or blip.attrib.get(f"{{{R_NS}}}link")


def _enhanced_image_bytes(data: bytes) -> bytes | None:
    try:
        from PIL import Image, ImageEnhance, ImageFilter
    except ImportError:
        return None
    try:
        image = Image.open(io.BytesIO(data)).convert("RGB")
    except OSError:
        return None
    bbox = _content_bbox(image)
    if bbox is not None:
        margin = max(8, int(min(image.size) * 0.025))
        left, top, right, bottom = bbox
        image = image.crop((
            max(0, left - margin),
            max(0, top - margin),
            min(image.width, right + margin),
            min(image.height, bottom + margin),
        ))
    target_width = 1800
    if image.width < target_width:
        height = max(1, int(image.height * target_width / image.width))
        image = image.resize((target_width, height), Image.Resampling.LANCZOS)
    elif image.width > 2400:
        target_width = 2400
        height = max(1, int(image.height * target_width / image.width))
        image = image.resize((target_width, height), Image.Resampling.LANCZOS)
    image = ImageEnhance.Contrast(image).enhance(1.18)
    image = ImageEnhance.Sharpness(image).enhance(1.35)
    image = image.filter(ImageFilter.UnsharpMask(radius=1.0, percent=130, threshold=3))
    buffer = io.BytesIO()
    image.save(buffer, format="PNG", optimize=True)
    return buffer.getvalue()


def _redrawn_known_figure(caption: str) -> bytes | None:
    compact = re.sub(r"\s+", "", caption).lower()
    if "实验室教师需求图" in compact:
        return _simple_flowchart_png(
            "实验室教师核心需求关系图",
            [
                ("危化品存储", "实时查看数据"),
                ("环境监测", "掌握危化品状态"),
                ("危化品拿取", "气体泄露预警"),
            ],
        )
    if "环境监测需求图" in compact:
        return _simple_flowchart_png(
            "环境监管技术需求关系图",
            [
                ("环境温湿度", "实时平台流通"),
                ("气体泄露状况", "可视化展示"),
                ("多阈值配置", "持续观测响应"),
            ],
        )
    return None


def _skip_generic_enhancement(caption: str) -> bool:
    compact = re.sub(r"\s+", "", caption)
    if any(token in compact for token in ("需求图", "流程图", "架构图", "框图", "原理图", "电路", "配置", "界面", "网页", "代码", "公式")):
        return False
    return any(token in compact for token in ("传感器", "成品图", "实物", "模块图"))


def _simple_flowchart_png(title_text: str, pairs: list[tuple[str, str]]) -> bytes:
    from PIL import Image, ImageDraw

    image = Image.new("RGB", (1500, 560), "white")
    draw = ImageDraw.Draw(image)
    title = _font(50)
    font = _font(42)
    line = (20, 20, 20)
    _center_text(draw, (750, 60), title_text, title, line)
    box_w, box_h = 330, 96
    start_x = 110
    gap_x = 145
    top_y, bottom_y = 145, 340
    for idx, (top, bottom) in enumerate(pairs):
        x = start_x + idx * (box_w + gap_x)
        _simple_box(draw, (x, top_y, x + box_w, top_y + box_h), top, font)
        _simple_box(draw, (x, bottom_y, x + box_w, bottom_y + box_h), bottom, font)
        _arrow(draw, (x + box_w // 2, top_y + box_h), (x + box_w // 2, bottom_y), fill=line)
    return _png_bytes(_trim_image_whitespace(image, margin=42))


def _simple_box(draw, box: tuple[int, int, int, int], text: str, font) -> None:
    draw.rounded_rectangle(box, radius=8, fill="white", outline=(20, 20, 20), width=4)
    _center_text(draw, ((box[0] + box[2]) // 2, (box[1] + box[3]) // 2), text, font, (20, 20, 20))


def _clock_config_png() -> bytes:
    from PIL import Image, ImageDraw

    image = Image.new("RGB", (1600, 900), "white")
    draw = ImageDraw.Draw(image)
    font = _font(34)
    small = _font(27)
    title = _font(42)
    line = (20, 20, 20)
    blue = (23, 126, 201)
    pale = (232, 245, 255)
    _center_text(draw, (800, 55), "STM32 时钟配置", title, line)
    boxes = [
        ((90, 250, 300, 360), "HSE\n8 MHz"),
        ((430, 250, 650, 360), "PLL\nx9"),
        ((790, 250, 1040, 360), "SYSCLK\n72 MHz"),
        ((1180, 160, 1480, 250), "AHB / HCLK\n72 MHz"),
        ((1180, 310, 1480, 400), "APB1 / PCLK1\n36 MHz"),
        ((1180, 460, 1480, 550), "APB2 / PCLK2\n72 MHz"),
        ((1180, 610, 1480, 700), "ADC Clock\n12 MHz"),
    ]
    for box, label in boxes:
        draw.rounded_rectangle(box, radius=10, outline=blue, width=4, fill=pale)
        _center_multiline(draw, box, label, font, line)
    _arrow(draw, (300, 305), (430, 305), blue)
    _arrow(draw, (650, 305), (790, 305), blue)
    for y in (205, 355, 505, 655):
        _arrow(draw, (1040, 305), (1180, y), blue)
    draw.text((100, 780), "配置要点：外部高速晶振 HSE=8 MHz，经 PLL 9 倍频后得到 72 MHz 系统时钟。", font=small, fill=line)
    draw.text((100, 825), "APB1 分频为 36 MHz，APB2 保持 72 MHz，ADC 时钟分频后为 12 MHz。", font=small, fill=line)
    return _png_bytes(image)


def _web_page_png() -> bytes:
    from PIL import Image, ImageDraw

    image = Image.new("RGB", (1600, 900), (247, 250, 253))
    draw = ImageDraw.Draw(image)
    title = _font(42)
    font = _font(50)
    small = _font(32)
    label_font = _font(35)
    line = (35, 44, 54)
    blue = (32, 121, 199)
    _center_text(draw, (800, 50), "OneNET Web 监控页面", title, line)
    draw.rounded_rectangle((80, 110, 330, 810), radius=12, fill="white", outline=(210, 220, 230), width=3)
    for idx, item in enumerate(["平台概览", "产品开发", "设备接入管理", "设备管理", "云网关接入", "数据流转", "运维监控"]):
        y = 155 + idx * 78
        fill = (229, 243, 255) if item == "设备接入管理" else "white"
        draw.rounded_rectangle((105, y - 28, 305, y + 28), radius=8, fill=fill)
        draw.text((125, y), item, font=small, fill=blue if item == "设备接入管理" else line, anchor="lm")
    cards = [
        ("最高有害气体浓度", "700", "int32 / 读写"),
        ("最高温度", "80", "int32 / 读写"),
        ("最高湿度", "35", "int32 / 读写"),
        ("最低温度", "40", "int32 / 读写"),
        ("最低湿度", "22", "int32 / 读写"),
        ("有害气体浓度", "2", "int32 / 只读"),
        ("湿度", "61", "int32 / 只读"),
        ("门锁", "false(0)", "bool / 读写"),
        ("id卡号", "2", "int32 / 只读"),
        ("温度", "24", "int32 / 只读"),
    ]
    start_x, start_y = 390, 145
    card_w, card_h = 350, 155
    gap_x, gap_y = 35, 26
    for idx, (label, value, meta) in enumerate(cards):
        row, col = divmod(idx, 3)
        x = start_x + col * (card_w + gap_x)
        y = start_y + row * (card_h + gap_y)
        draw.rounded_rectangle((x, y, x + card_w, y + card_h), radius=12, fill="white", outline=(220, 228, 236), width=2)
        draw.text((x + 28, y + 45), label, font=label_font, fill=line, anchor="lm")
        draw.text((x + 28, y + 105), value, font=font, fill=(0, 0, 0), anchor="lm")
    return _png_bytes(image)


def _pin_config_png() -> bytes:
    from PIL import Image, ImageDraw

    image = Image.new("RGB", (1600, 900), "white")
    draw = ImageDraw.Draw(image)
    title = _font(50)
    font = _font(40)
    header = _font(42)
    line = (30, 30, 30)
    blue = (23, 126, 201)
    _center_text(draw, (800, 62), "STM32CubeMX 硬件外设配置", title, line)
    draw.rounded_rectangle((70, 125, 1530, 835), radius=12, outline=blue, width=5, fill=(248, 252, 255))
    cols = [100, 430, 755, 1065]
    headers = ["模块", "端口/引脚", "配置方式", "用途"]
    for x, text in zip(cols, headers):
        draw.text((x, 190), text, font=header, fill=blue, anchor="lm")
    rows = [
        ("RC522 SDA", "PB4", "GPIO_Output", "软件 SPI 片选/数据线"),
        ("RC522 SCK", "PB5", "GPIO_Output", "软件 SPI 时钟"),
        ("RC522 MOSI", "PB6", "GPIO_Output", "软件 SPI 主发"),
        ("RC522 MISO", "PB7", "GPIO_Input", "软件 SPI 主收"),
        ("RC522 RST", "PB12", "GPIO_Output", "RFID 复位控制"),
        ("传感器采集", "ADC1", "Analog", "气体浓度模拟量采集"),
        ("显示/通信", "I2C / USART", "Alternate", "数据显示与调试通信"),
    ]
    y = 270
    for idx, row in enumerate(rows):
        fill = (238, 247, 255) if idx % 2 == 0 else "white"
        draw.rectangle((90, y - 38, 1510, y + 38), fill=fill)
        for x, text in zip(cols, row):
            draw.text((x, y), text, font=font, fill=line, anchor="lm")
        y += 82
    return _png_bytes(image)


def _font(size: int):
    from PIL import ImageFont

    for name in (
        "/mnt/c/Windows/Fonts/simsun.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        "simsun.ttc",
        "SimSun.ttf",
        "NotoSansCJK-Regular.ttc",
        "NotoSans-Regular.ttf",
    ):
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            continue
    return ImageFont.load_default()


def _center_text(draw, xy: tuple[int, int], text: str, font, fill=(0, 0, 0)) -> None:
    draw.text(xy, text, font=font, fill=fill, anchor="mm")


def _center_multiline(draw, box: tuple[int, int, int, int], text: str, font, fill=(0, 0, 0)) -> None:
    lines = text.splitlines()
    line_heights = []
    widths = []
    for line in lines:
        bbox = draw.textbbox((0, 0), line, font=font)
        widths.append(bbox[2] - bbox[0])
        line_heights.append(bbox[3] - bbox[1])
    total_h = sum(line_heights) + max(0, len(lines) - 1) * 8
    y = (box[1] + box[3] - total_h) / 2
    for line, line_h in zip(lines, line_heights):
        draw.text(((box[0] + box[2]) / 2, y), line, font=font, fill=fill, anchor="ma")
        y += line_h + 8


def _arrow(draw, start: tuple[int, int], end: tuple[int, int], fill=(0, 0, 0)) -> None:
    draw.line((start, end), fill=fill, width=5)
    ex, ey = end
    sx, sy = start
    if abs(ex - sx) >= abs(ey - sy):
        direction = 1 if ex >= sx else -1
        points = [(ex, ey), (ex - direction * 18, ey - 10), (ex - direction * 18, ey + 10)]
    else:
        direction = 1 if ey >= sy else -1
        points = [(ex, ey), (ex - 10, ey - direction * 18), (ex + 10, ey - direction * 18)]
    draw.polygon(points, fill=fill)


def _png_bytes(image) -> bytes:
    buffer = io.BytesIO()
    image.save(buffer, format="PNG", optimize=True)
    return buffer.getvalue()


def _trim_image_whitespace(image, margin: int):
    bbox = _content_bbox(image)
    if bbox is None:
        return image
    left, top, right, bottom = bbox
    return image.crop(
        (
            max(0, left - margin),
            max(0, top - margin),
            min(image.width, right + margin),
            min(image.height, bottom + margin),
        )
    )


def _content_bbox(image) -> tuple[int, int, int, int] | None:
    gray = image.convert("L")
    pixels = gray.load()
    xs: list[int] = []
    ys: list[int] = []
    for y in range(gray.height):
        for x in range(gray.width):
            if pixels[x, y] < 245:
                xs.append(x)
                ys.append(y)
    if not xs or not ys:
        return None
    return min(xs), min(ys), max(xs) + 1, max(ys) + 1


def _expand_drawing_extent(paragraph: ET.Element, image_bytes: bytes) -> None:
    try:
        from PIL import Image
        image = Image.open(io.BytesIO(image_bytes))
        width, height = image.size
    except Exception:
        width, height = 16, 9
    max_cx = 5_400_000
    max_cy = 3_700_000
    aspect = height / max(width, 1)
    cx = max_cx
    cy = int(cx * aspect)
    if cy > max_cy:
        cy = max_cy
        cx = int(cy / max(aspect, 0.01))
    for extent in paragraph.findall(f".//{{{WP_NS}}}extent"):
        extent.set("cx", str(cx))
        extent.set("cy", str(cy))
    for ext in paragraph.findall(f".//{{{A_NS}}}ext"):
        ext.set("cx", str(cx))
        ext.set("cy", str(cy))


def _code_formula_replacement_paragraphs(caption: str, kind: str, label: str) -> list[ET.Element]:
    compact = re.sub(r"\s+", "", caption).lower()
    if kind == "formula":
        lines = _known_formula_lines(compact)
        title = "气体浓度换算公式如下："
        chapter = _formula_chapter(label)
        return [_paragraph(title, kind="body")] + [_formula_paragraph(line, chapter, idx) for idx, line in enumerate(lines, start=1)]
    lines = _known_code_lines(compact)
    title = "RC522 软件 SPI 引脚定义如下：" if "rc522" in compact else "相关代码定义如下："
    return [_paragraph(title, kind="body")] + [_paragraph(line, kind="code") for line in lines]


def _formula_chapter(label: str) -> str:
    match = re.match(r"图(\d+)-\d+", label)
    return match.group(1) if match else "3"


def _known_code_lines(compact_caption: str) -> list[str]:
    if "rc522" in compact_caption:
        return [
            "#define MFRC522_GPIO_SDA_PORT   GPIOB",
            "#define MFRC522_GPIO_SDA_PIN    GPIO_PIN_4",
            "#define MFRC522_GPIO_SCK_PORT   GPIOB",
            "#define MFRC522_GPIO_SCK_PIN    GPIO_PIN_5",
            "#define MFRC522_GPIO_MOSI_PORT  GPIOB",
            "#define MFRC522_GPIO_MOSI_PIN   GPIO_PIN_6",
            "#define MFRC522_GPIO_MISO_PORT  GPIOB",
            "#define MFRC522_GPIO_MISO_PIN   GPIO_PIN_7",
            "#define MFRC522_GPIO_RST_PORT   GPIOB",
            "#define MFRC522_GPIO_RST_PIN    GPIO_PIN_12",
        ]
    return ["[代码截图已移除，OCR 未能自动提取，需人工复核补录。]"]


def _known_formula_lines(compact_caption: str) -> list[str]:
    if "气体质量计算公式" in compact_caption or "气体" in compact_caption:
        return [
            "V = ADC × 5 / 4096",
            "RS = (5 - V) / (V × 0.5)",
            "R0 = 6.64",
            "ppm = 11.5428 × (R0 / RS)^0.6549",
        ]
    return ["[公式截图已移除，OCR 未能自动提取，需人工复核补录。]"]


def _paragraph(text: str, kind: str) -> ET.Element:
    p = ET.Element(_w("p"))
    ppr = ET.SubElement(p, _w("pPr"))
    spacing = ET.SubElement(ppr, _w("spacing"))
    spacing.set(_w("before"), "0")
    spacing.set(_w("after"), "0")
    spacing.set(_w("line"), "300" if kind == "body" else "240")
    spacing.set(_w("lineRule"), "auto")
    if kind == "body":
        ind = ET.SubElement(ppr, _w("ind"))
        ind.set(_w("firstLine"), "480")
        ind.set(_w("firstLineChars"), "200")
        jc = ET.SubElement(ppr, _w("jc"))
        jc.set(_w("val"), "both")
    else:
        ind = ET.SubElement(ppr, _w("ind"))
        ind.set(_w("left"), "480")
        jc = ET.SubElement(ppr, _w("jc"))
        jc.set(_w("val"), "left")
    r = ET.SubElement(p, _w("r"))
    rpr = ET.SubElement(r, _w("rPr"))
    fonts = ET.SubElement(rpr, _w("rFonts"))
    font = "Consolas" if kind == "code" else "Times New Roman"
    fonts.set(_w("ascii"), font)
    fonts.set(_w("hAnsi"), font)
    fonts.set(_w("eastAsia"), "宋体")
    sz = ET.SubElement(rpr, _w("sz"))
    sz.set(_w("val"), "21" if kind != "body" else "24")
    sz_cs = ET.SubElement(rpr, _w("szCs"))
    sz_cs.set(_w("val"), "21" if kind != "body" else "24")
    t = ET.SubElement(r, _w("t"))
    t.text = text
    if re.search(r"^\s|\s$|\s{2,}", text):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return p


def _formula_paragraph(text: str, chapter: str, number: int) -> ET.Element:
    p = ET.Element(_w("p"))
    ppr = ET.SubElement(p, _w("pPr"))
    tabs = ET.SubElement(ppr, _w("tabs"))
    center = ET.SubElement(tabs, _w("tab"))
    center.set(_w("val"), "center")
    center.set(_w("pos"), "4513")
    right = ET.SubElement(tabs, _w("tab"))
    right.set(_w("val"), "right")
    right.set(_w("pos"), "9000")
    spacing = ET.SubElement(ppr, _w("spacing"))
    spacing.set(_w("before"), "120")
    spacing.set(_w("after"), "120")
    spacing.set(_w("line"), "300")
    spacing.set(_w("lineRule"), "auto")
    jc = ET.SubElement(ppr, _w("jc"))
    jc.set(_w("val"), "left")
    _append_formula_tab(p)
    for fragment, vertical_align in _formula_fragments(text):
        _append_formula_text_run(p, fragment, vertical_align)
    _append_formula_tab(p)
    _append_formula_text_run(p, f"({chapter}-{number})")
    return p


def _append_formula_tab(paragraph: ET.Element) -> None:
    run = ET.SubElement(paragraph, _w("r"))
    ET.SubElement(run, _w("tab"))


def _append_formula_text_run(paragraph: ET.Element, text: str, vertical_align: str | None = None) -> None:
    run = ET.SubElement(paragraph, _w("r"))
    rpr = ET.SubElement(run, _w("rPr"))
    fonts = ET.SubElement(rpr, _w("rFonts"))
    fonts.set(_w("ascii"), "Times New Roman")
    fonts.set(_w("hAnsi"), "Times New Roman")
    fonts.set(_w("eastAsia"), "宋体")
    for name in ("sz", "szCs"):
        node = ET.SubElement(rpr, _w(name))
        node.set(_w("val"), "24")
    if vertical_align:
        align = ET.SubElement(rpr, _w("vertAlign"))
        align.set(_w("val"), vertical_align)
    t = ET.SubElement(run, _w("t"))
    t.text = text
    if re.search(r"^\s|\s$|\s{2,}", text):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def _formula_fragments(text: str) -> list[tuple[str, str | None]]:
    normalized = text.replace("*", "×").replace("Voltage", "V")
    fragments: list[tuple[str, str | None]] = []
    i = 0
    while i < len(normalized):
        if normalized.startswith("R0", i):
            _append_formula_fragment(fragments, "R", None)
            _append_formula_fragment(fragments, "0", "subscript")
            i += 2
            continue
        if normalized.startswith("RS", i) or normalized.startswith("Rs", i):
            _append_formula_fragment(fragments, "R", None)
            _append_formula_fragment(fragments, "S", "subscript")
            i += 2
            continue
        if normalized[i] == "^":
            i += 1
            while i < len(normalized) and normalized[i].isspace():
                i += 1
            start = i
            while i < len(normalized) and re.match(r"[0-9.+\\-]", normalized[i]):
                i += 1
            exponent = normalized[start:i]
            if exponent:
                _append_formula_fragment(fragments, exponent, "superscript")
            continue
        _append_formula_fragment(fragments, normalized[i], None)
        i += 1
    return fragments


def _append_formula_fragment(fragments: list[tuple[str, str | None]], text: str, vertical_align: str | None) -> None:
    if not text:
        return
    if fragments and fragments[-1][1] == vertical_align:
        fragments[-1] = (fragments[-1][0] + text, vertical_align)
    else:
        fragments.append((text, vertical_align))


def _renumber_figure_captions(body: ET.Element) -> tuple[dict[str, str], int]:
    current_chapter: str | None = None
    counters: dict[str, int] = {}
    mapping: dict[str, str] = {}
    changed = 0
    for child in list(body):
        if child.tag != _w("p"):
            continue
        text = _paragraph_text(child).strip()
        heading = re.match(r"^([1-9])\s+[\u4e00-\u9fffA-Za-z]", text)
        if heading:
            current_chapter = heading.group(1)
            continue
        match = re.match(r"^图\s*(\d+)\s*[-－]\s*(\d+)\s*(.*)$", text)
        if not match:
            continue
        chapter = current_chapter or match.group(1)
        counters[chapter] = counters.get(chapter, 0) + 1
        old = f"图{match.group(1)}-{match.group(2)}"
        new = f"图{chapter}-{counters[chapter]}"
        title = match.group(3).strip()
        new_text = f"{new} {title}".strip()
        if old != new or text != new_text:
            mapping[_compact_label(old)] = new
            _replace_paragraph_text(child, new_text)
            changed += 1
    return mapping, changed


def _update_figure_references(body: ET.Element, removed: dict[str, str], renumbered: dict[str, str]) -> int:
    updates = 0
    replacements = {**renumbered, **removed}
    if not replacements:
        return 0
    for paragraph in body.iter(_w("p")):
        text = _paragraph_text(paragraph)
        if not text or _is_figure_caption(text):
            continue
        updated = text
        for old, new in replacements.items():
            old_match = re.match(r"图(\d+)-(\d+)", old)
            if not old_match:
                continue
            pattern = rf"图\s*{old_match.group(1)}\s*[-－]\s*{old_match.group(2)}"
            if new.startswith("如下"):
                updated = re.sub(rf"如\s*{pattern}", new, updated)
                updated = re.sub(pattern, new, updated)
            else:
                updated = re.sub(pattern, new.replace(" ", ""), updated)
        if updated != text:
            _replace_paragraph_text(paragraph, updated)
            updates += 1
    return updates


def _compact_label(label: str) -> str:
    return re.sub(r"\s+", "", label).replace("－", "-")


def _replace_paragraph_text(paragraph: ET.Element, text: str) -> None:
    text_nodes = list(paragraph.findall(".//w:t", NS))
    if not text_nodes:
        paragraph.append(_paragraph(text, kind="body").find("w:r", NS))
        return
    text_nodes[0].text = text
    if re.search(r"^\s|\s$|\s{2,}", text):
        text_nodes[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    for node in text_nodes[1:]:
        node.text = ""


def _w(local: str) -> str:
    return f"{{{W_NS}}}{local}"
