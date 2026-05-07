from __future__ import annotations

import re
import subprocess
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET

from .config import AgentConfig
from .pdf_inspect import extract_page_texts, normalize_text, pdf_info
from .tools import Toolchain, run_cmd


@dataclass(frozen=True)
class PageInk:
    page: int
    ink_ratio: float
    path: Path


@dataclass(frozen=True)
class VisualInspection:
    pages: int
    page_size: str | None
    blank_pages: list[int] = field(default_factory=list)
    near_blank_pages: list[int] = field(default_factory=list)
    page_number_labels: dict[int, str] = field(default_factory=dict)
    heading_pages: dict[str, list[int]] = field(default_factory=dict)
    broken_reference_pages: list[int] = field(default_factory=list)
    toc_title_split_pages: list[int] = field(default_factory=list)
    front_matter_page_number_errors: dict[int, str] = field(default_factory=dict)
    toc_page_number_mismatches: list[str] = field(default_factory=list)
    front_matter_layout_errors: list[str] = field(default_factory=list)
    header_page_number_alignment_errors: list[str] = field(default_factory=list)
    caption_orphan_pages: list[int] = field(default_factory=list)
    figure_readability_warnings: list[str] = field(default_factory=list)
    page_ink: list[PageInk] = field(default_factory=list)


@dataclass(frozen=True)
class CaptionBox:
    page: int
    kind: str
    text: str
    y_min: float
    y_max: float
    page_height: float


def inspect_visual(pdf_path: Path, png_pages: list[Path], toolchain: Toolchain, config: AgentConfig) -> VisualInspection:
    info = pdf_info(pdf_path, toolchain)
    page_texts = extract_page_texts(pdf_path, toolchain, info.pages)
    page_ink = [_ink_ratio(page, path) for page, path in enumerate(png_pages, start=1)]
    page_labels = {
        page: label
        for page, text in page_texts.items()
        if (label := _infer_page_number_label(text)) is not None
    }
    blank_pages = [
        item.page
        for item in page_ink
        if item.ink_ratio <= config.blank_ink_ratio
        and _meaningful_line_count(page_texts.get(item.page, "")) == 0
    ]
    near_blank_pages = [
        item.page
        for item in page_ink
        if item.ink_ratio <= config.near_blank_ink_ratio
        and _meaningful_line_count(page_texts.get(item.page, "")) <= 1
        and not _is_expected_sparse_page(page_texts.get(item.page, ""), page_texts.get(item.page + 1, ""))
    ]
    heading_pages = _find_heading_pages(page_texts, config.main_heading_patterns)
    front_page_errors = _detect_front_matter_page_number_errors(page_texts, page_labels)
    toc_mismatches = _detect_toc_page_number_mismatches(page_texts, page_labels)
    front_layout_errors = _detect_front_matter_layout_errors(page_texts)
    header_alignment_errors = _detect_header_page_number_alignment(pdf_path, toolchain, info.pages)
    broken_pages = [
        page for page, text in page_texts.items()
        if "Error: Reference source not found" in text or "错误!未找到引用源" in text
    ]
    toc_split_pages = _detect_split_toc_title(page_texts)
    caption_boxes = _extract_caption_boxes(pdf_path, toolchain, info.pages)
    caption_orphans = _detect_caption_orphans(page_texts, caption_boxes, png_pages)
    figure_warnings = _detect_figure_readability_warnings(caption_boxes, png_pages)
    return VisualInspection(
        pages=info.pages,
        page_size=info.page_size,
        blank_pages=blank_pages,
        near_blank_pages=near_blank_pages,
        page_number_labels=page_labels,
        heading_pages=heading_pages,
        broken_reference_pages=broken_pages,
        toc_title_split_pages=toc_split_pages,
        front_matter_page_number_errors=front_page_errors,
        toc_page_number_mismatches=toc_mismatches,
        front_matter_layout_errors=front_layout_errors,
        header_page_number_alignment_errors=header_alignment_errors,
        caption_orphan_pages=caption_orphans,
        figure_readability_warnings=figure_warnings,
        page_ink=page_ink,
    )


def _ink_ratio(page: int, image_path: Path) -> PageInk:
    try:
        from PIL import Image
    except ImportError as exc:
        raise RuntimeError("Pillow is required for visual inspection") from exc

    with Image.open(image_path).convert("L") as image:
        pixels = image.getdata()
        total = image.width * image.height
        ink = sum(1 for value in pixels if value < 245)
    return PageInk(page=page, ink_ratio=ink / max(total, 1), path=image_path)


def _find_heading_pages(page_texts: dict[int, str], patterns: list[str]) -> dict[str, list[int]]:
    toc_pages = [page for page, text in page_texts.items() if _is_probable_toc_page(text)]
    after_toc = max(toc_pages) if toc_pages else 0
    result: dict[str, list[int]] = {}
    compiled = [(pattern, re.compile(pattern, re.MULTILINE)) for pattern in patterns]
    for page, text in page_texts.items():
        if page <= after_toc:
            continue
        cleaned_lines = "\n".join(line.strip() for line in text.splitlines() if line.strip())
        for pattern, regex in compiled:
            if regex.search(cleaned_lines):
                result.setdefault(pattern, []).append(page)
    return result


def _is_probable_toc_page(text: str) -> bool:
    normalized = normalize_text(text)
    if "目录" in normalized:
        return True
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    leader_lines = sum(1 for line in lines if re.search(r"\.{4,}\s*\d+\s*$", line))
    numbered_heading_lines = sum(1 for line in lines if re.match(r"^[1-9](\.\d+)?\s+.+\.{4,}\s*\d+\s*$", line))
    back_matter_lines = sum(1 for line in lines if re.match(r"^(参考文献|致谢|附录).+\.{4,}\s*\d+\s*$", line))
    return leader_lines >= 3 or numbered_heading_lines >= 3 or (leader_lines >= 1 and back_matter_lines >= 1)


def _detect_front_matter_page_number_errors(page_texts: dict[int, str], page_labels: dict[int, str]) -> dict[int, str]:
    first_body_page = _find_first_body_page(page_texts)
    if first_body_page is None:
        return {}
    errors: dict[int, str] = {}
    for page, label in sorted(page_labels.items()):
        if page >= first_body_page:
            continue
        if label.isdigit():
            errors[page] = label
    return errors


def _detect_toc_page_number_mismatches(page_texts: dict[int, str], page_labels: dict[int, str]) -> list[str]:
    heading_labels = _rendered_heading_label_map(page_texts, page_labels)
    toc_entries = _toc_entries_from_page_texts(page_texts)
    mismatches: list[str] = []
    for key, (title, toc_label) in toc_entries.items():
        actual = heading_labels.get(key)
        if actual and actual != toc_label:
            mismatches.append(f"{title}: 目录 {toc_label} / 正文 {actual}")
    return mismatches


def _detect_front_matter_layout_errors(page_texts: dict[int, str]) -> list[str]:
    errors: list[str] = []
    incompatible_groups = [
        ("封面", "学术诚信声明", ["学士学位论文", "毕业设计（论文）学士学位论文"], ["学术诚信声明"]),
        ("学术诚信声明", "AI 使用情况声明", ["学术诚信声明"], ["AI 使用情况声明", "AI使用情况声明"]),
        ("AI 使用情况声明", "版权使用授权书", ["AI 使用情况声明", "AI使用情况声明"], ["版权使用授权书"]),
        ("版权使用授权书", "摘要", ["版权使用授权书"], ["摘   要", "摘  要", "摘要"]),
        ("中文摘要", "英文摘要", ["摘   要", "摘  要", "摘要"], ["ABSTRACT"]),
        ("英文摘要", "目录", ["ABSTRACT"], ["目  录", "目录"]),
    ]
    for page, text in page_texts.items():
        compact = re.sub(r"\s+", "", text)
        if "学术诚信声明" in compact and "日期" in compact:
            errors.append(f"page {page}: 学术诚信声明日期页位置异常，疑似模板前置页版式被压缩")
        for left, right, left_markers, right_markers in incompatible_groups:
            if any(re.sub(r"\s+", "", marker) in compact for marker in left_markers) and any(
                re.sub(r"\s+", "", marker) in compact for marker in right_markers
            ):
                errors.append(f"page {page}: {left} 与 {right} 不应混在同一页")
    return errors[:12]


def _rendered_heading_label_map(page_texts: dict[int, str], page_labels: dict[int, str]) -> dict[str, str]:
    toc_pages = [page for page, text in page_texts.items() if _is_probable_toc_page(text)]
    after_toc = max(toc_pages) if toc_pages else 0
    result: dict[str, str] = {}
    for page, text in page_texts.items():
        if page <= after_toc:
            continue
        label = page_labels.get(page, str(page))
        for line in _meaningful_lines(text):
            title = _heading_title_from_rendered_line(line)
            if title:
                result.setdefault(_heading_key(title), label)
    return result


def _toc_entries_from_page_texts(page_texts: dict[int, str]) -> dict[str, tuple[str, str]]:
    entries: dict[str, tuple[str, str]] = {}
    for page, text in page_texts.items():
        if not _is_probable_toc_page(text):
            continue
        for line in _meaningful_lines(text):
            parsed = _parse_toc_entry_line(line)
            if parsed is None:
                continue
            title, label = parsed
            key = _heading_key(title)
            entries.setdefault(key, (title, label))
    return entries


def _find_first_body_page(page_texts: dict[int, str]) -> int | None:
    toc_pages = [page for page, text in page_texts.items() if _is_probable_toc_page(text)]
    after_toc = max(toc_pages) if toc_pages else 0
    for page, text in page_texts.items():
        if page <= after_toc:
            continue
        for line in _meaningful_lines(text):
            if re.fullmatch(r"1\s+绪论", line):
                return page
    return None


def _heading_title_from_rendered_line(line: str) -> str | None:
    stripped = line.strip()
    if _parse_toc_entry_line(stripped) is not None:
        return None
    if re.match(r"^[1-9]\s*[\u4e00-\u9fffA-Za-z].{0,50}$", stripped):
        return stripped
    if re.match(r"^[1-9]\.\d+(?:\.\d+)?\s+[\u4e00-\u9fffA-Za-z].{0,80}$", stripped):
        return stripped
    compact = re.sub(r"\s+", "", stripped)
    if compact in {"参考文献", "致谢", "附录"}:
        return compact
    return None


def _parse_toc_entry_line(line: str) -> tuple[str, str] | None:
    stripped = line.strip()
    if not stripped or "目录" in re.sub(r"\s+", "", stripped):
        return None
    if not re.match(r"^([1-9](?:\.\d+)*\s*|参考文献|致\s*谢|附录)", stripped):
        return None
    match = re.match(r"^(?P<title>.+?)(?:\.{2,}|…{2,}|\t+|\s{2,})(?P<label>\d+)\s*$", stripped)
    if not match:
        match = re.match(r"^(?P<title>(?:[1-9](?:\.\d+)*\s*.+?|参考文献|致\s*谢|附录))(?P<label>\d{1,3})\s*$", stripped)
    if not match:
        return None
    title = match.group("title").strip()
    label = match.group("label")
    return title, label


def _meaningful_lines(text: str) -> list[str]:
    return [line.strip() for line in text.splitlines() if line.strip()]


def _heading_key(title: str) -> str:
    title = re.sub(r"(?:\.{2,}|…{2,}|\t+|\s{2,})\d+\s*$", "", title).strip()
    return re.sub(r"\s+", "", title)


def _extract_caption_boxes(pdf_path: Path, toolchain: Toolchain, pages: int) -> dict[int, list[CaptionBox]]:
    if not toolchain.pdftotext:
        return {}
    try:
        result = run_cmd([toolchain.pdftotext, "-bbox-layout", str(pdf_path), "-"])
    except (subprocess.CalledProcessError, RuntimeError, ET.ParseError):
        return {}
    try:
        root = ET.fromstring(result.stdout)
    except ET.ParseError:
        return {}

    caption_pattern = re.compile(r"^(图|表)\s*\d+\s*[-－]\s*\d+")
    found: dict[int, list[CaptionBox]] = {}
    page_nodes = [node for node in root.iter() if _local_name(node.tag) == "page"]
    for page_index, page_node in enumerate(page_nodes[:pages], start=1):
        page_height = float(page_node.attrib.get("height", "0") or "0")
        if page_height <= 0:
            continue
        for line in page_node.iter():
            if _local_name(line.tag) != "line":
                continue
            words = [word.text or "" for word in line if _local_name(word.tag) == "word"]
            text = " ".join(part for part in words if part).strip()
            match = caption_pattern.match(text)
            if not match:
                continue
            try:
                y_min = float(line.attrib["yMin"])
                y_max = float(line.attrib["yMax"])
            except (KeyError, ValueError):
                continue
            found.setdefault(page_index, []).append(
                CaptionBox(
                    page=page_index,
                    kind=match.group(1),
                    text=text,
                    y_min=y_min,
                    y_max=y_max,
                    page_height=page_height,
                )
            )
    return found


def _detect_header_page_number_alignment(pdf_path: Path, toolchain: Toolchain, pages: int) -> list[str]:
    if not toolchain.pdftotext:
        return []
    try:
        result = run_cmd([toolchain.pdftotext, "-bbox", str(pdf_path), "-"])
        root = ET.fromstring(result.stdout)
    except (subprocess.CalledProcessError, RuntimeError, ET.ParseError):
        return []
    errors: list[str] = []
    page_nodes = [node for node in root.iter() if _local_name(node.tag) == "page"]
    for page_index, page_node in enumerate(page_nodes[:pages], start=1):
        try:
            page_width = float(page_node.attrib.get("width", "0") or "0")
        except ValueError:
            continue
        if page_width <= 0:
            continue
        top_words = []
        for word in page_node.iter():
            if _local_name(word.tag) != "word":
                continue
            try:
                y_min = float(word.attrib["yMin"])
                x_min = float(word.attrib["xMin"])
                x_max = float(word.attrib["xMax"])
            except (KeyError, ValueError):
                continue
            text = (word.text or "").strip()
            if y_min <= 75 and re.fullmatch(r"[ivxlcdmIVXLCDM]+|\d{1,3}", text):
                top_words.append((x_min, x_max, text))
        if not top_words:
            continue
        _x_min, x_max, label = max(top_words, key=lambda item: item[1])
        stale_left = [item for item in top_words if item[1] < page_width * 0.35]
        if stale_left and x_max >= page_width - 90:
            errors.append(f"page {page_index}: 页眉左侧疑似残留旧页码 {stale_left[0][2]}")
            continue
        if x_max < page_width - 90:
            errors.append(f"page {page_index}: 页眉页码 {label} 未靠近右侧页边")
    return errors[:20]


def _detect_figure_readability_warnings(caption_boxes: dict[int, list[CaptionBox]], png_pages: list[Path]) -> list[str]:
    try:
        from PIL import Image
    except ImportError:
        return []
    warnings: list[str] = []
    for page, boxes in sorted(caption_boxes.items()):
        if page < 1 or page > len(png_pages):
            continue
        try:
            image = Image.open(png_pages[page - 1]).convert("L")
        except OSError:
            continue
        scale = image.height / max(boxes[0].page_height, 1)
        for box in boxes:
            caption_top = int(box.y_min * scale)
            crop_top = max(0, caption_top - int(210 * scale))
            crop_bottom = max(crop_top + 1, caption_top - int(8 * scale))
            crop = image.crop((int(image.width * 0.08), crop_top, int(image.width * 0.92), crop_bottom))
            warning = _small_text_warning_for_crop(crop, page, box.text)
            if warning:
                warnings.append(warning)
                break
    return warnings[:12]


def _small_text_warning_for_crop(image, page: int, caption: str) -> str | None:
    width, height = image.size
    if width <= 0 or height <= 0:
        return None
    binary = image.point(lambda value: 0 if value < 180 else 255, mode="1")
    pixels = binary.load()
    visited: set[tuple[int, int]] = set()
    component_heights: list[int] = []
    dark = 0
    for y in range(height):
        for x in range(width):
            if pixels[x, y] != 0:
                continue
            dark += 1
            if (x, y) in visited:
                continue
            stack = [(x, y)]
            visited.add((x, y))
            min_y = max_y = y
            count = 0
            while stack:
                cx, cy = stack.pop()
                count += 1
                min_y = min(min_y, cy)
                max_y = max(max_y, cy)
                for nx in (cx - 1, cx, cx + 1):
                    for ny in (cy - 1, cy, cy + 1):
                        if nx < 0 or ny < 0 or nx >= width or ny >= height or (nx, ny) in visited:
                            continue
                        if pixels[nx, ny] == 0:
                            visited.add((nx, ny))
                            stack.append((nx, ny))
            if 2 <= count <= 160:
                component_heights.append(max_y - min_y + 1)
    ink_ratio = dark / max(width * height, 1)
    if ink_ratio < 0.01 or len(component_heights) < 120:
        return None
    component_heights.sort()
    median_height = component_heights[len(component_heights) // 2]
    if median_height <= 7:
        return f"page {page}: {caption} 上方图形疑似存在过小文字，建议放大或重绘"
    return None


def _detect_caption_orphans(
    page_texts: dict[int, str],
    caption_boxes: dict[int, list[CaptionBox]],
    png_pages: list[Path],
) -> list[int]:
    if not caption_boxes:
        return _detect_caption_orphans_from_text(page_texts)
    png_by_page = {page: path for page, path in enumerate(png_pages, start=1)}
    orphan_pages: list[int] = []
    for page, boxes in caption_boxes.items():
        image_path = png_by_page.get(page)
        for box in boxes:
            has_anchor = False
            if image_path is not None:
                if box.kind == "图":
                    has_anchor = _caption_has_visual_anchor(image_path, box, above=True)
                else:
                    has_anchor = _caption_has_visual_anchor(image_path, box, above=False)
            if not has_anchor and _caption_position_is_risky(box):
                orphan_pages.append(page)
                break
    return orphan_pages


def _detect_caption_orphans_from_text(page_texts: dict[int, str]) -> list[int]:
    orphan_pages: list[int] = []
    caption_pattern = re.compile(r"^(图|表)\s*\d+\s*[-－]\s*\d+", re.MULTILINE)
    for page, text in page_texts.items():
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        if not lines:
            continue
        caption_indices = [idx for idx, line in enumerate(lines) if caption_pattern.match(line)]
        for idx in caption_indices:
            tail_ratio = idx / max(len(lines), 1)
            if tail_ratio >= 0.85:
                orphan_pages.append(page)
                break
    return orphan_pages


def _caption_has_visual_anchor(image_path: Path, box: CaptionBox, above: bool) -> bool:
    try:
        from PIL import Image
    except ImportError:
        return False
    with Image.open(image_path).convert("L") as image:
        scale = image.height / max(box.page_height, 1)
        margin = max(int(8 * scale), 1)
        span = max(int(260 * scale), int(image.height * 0.28))
        if above:
            y_bottom = max(int(box.y_min * scale) - margin, 0)
            y_top = max(y_bottom - span, 0)
        else:
            y_top = min(int(box.y_max * scale) + margin, image.height)
            y_bottom = min(y_top + span, image.height)
        if y_bottom <= y_top:
            return False
        x_margin = int(image.width * 0.08)
        crop = image.crop((x_margin, y_top, image.width - x_margin, y_bottom))
        pixels = list(crop.getdata())
        if not pixels:
            return False
        ink_ratio = sum(1 for value in pixels if value < 245) / len(pixels)
        return ink_ratio >= 0.012


def _caption_position_is_risky(box: CaptionBox) -> bool:
    y_mid = (box.y_min + box.y_max) / 2
    ratio = y_mid / max(box.page_height, 1)
    return ratio <= 0.20 or ratio >= 0.72


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _detect_split_toc_title(page_texts: dict[int, str]) -> list[int]:
    pages: list[int] = []
    for page, text in page_texts.items():
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        next_lines = [line.strip() for line in page_texts.get(page + 1, "").splitlines() if line.strip()]
        if "目" in lines[-3:] and "录" in next_lines[:5]:
            pages.append(page)
    return pages


def _infer_page_number_label(text: str) -> str | None:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    for pos, line in enumerate(lines[:5] + lines[-5:]):
        match = re.search(r"(?:^|\s)([ivxlcdmIVXLCDM]+|\d+)\s*$", line)
        if not match:
            continue
        label = match.group(1)
        if label.isdigit() and len(label) > 4:
            continue
        in_top_header = pos < min(5, len(lines))
        if len(line) > 60 and not (in_top_header and re.search(r"\s{2,}([ivxlcdmIVXLCDM]+|\d+)\s*$", line)):
            continue
        return label
    return None


def _meaningful_line_count(text: str) -> int:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    count = 0
    for line in lines:
        if re.fullmatch(r"[ivxlcdmIVXLCDM]+|\d+", line):
            continue
        if re.search(r"\s([ivxlcdmIVXLCDM]+|\d+)$", line) and len(line) <= 40:
            continue
        count += 1
    return count


def _is_expected_sparse_page(text: str, next_text: str = "") -> bool:
    compact = re.sub(r"\s+", "", text)
    if any(marker in compact for marker in ["作者签名", "指导教师签名", "日期", "保密□", "不保密□"]):
        return True
    if _meaningful_line_count(text) == 1 and re.search(r"[\u4e00-\u9fff]", text):
        next_lines = [line.strip() for line in next_text.splitlines() if line.strip()]
        if any(
            re.fullmatch(r"[1-9]\s+[\u4e00-\u9fffA-Za-z].{0,40}", line)
            or re.fullmatch(r"参考文献|致\s*谢|附录", re.sub(r"\s+", "", line))
            for line in next_lines[:8]
        ):
            return True
    return False
