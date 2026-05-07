from __future__ import annotations

import json
import math
import re
import shutil
from dataclasses import dataclass, field
from pathlib import Path


@dataclass(frozen=True)
class VisionPack:
    audit_dir: Path
    out_dir: Path
    target_sheets: list[Path] = field(default_factory=list)
    template_sheets: list[Path] = field(default_factory=list)
    key_pages: list[Path] = field(default_factory=list)
    prompt: Path | None = None


def build_vision_pack(audit_dir: Path, out_dir: Path, thumb_width: int = 620, pages_per_sheet: int = 6) -> VisionPack:
    target_dir = audit_dir / "png" / "target"
    template_dir = audit_dir / "png" / "template"
    if not target_dir.exists():
        raise FileNotFoundError(f"Target PNG directory not found: {target_dir}")
    out_dir.mkdir(parents=True, exist_ok=True)

    target_sheets = _contact_sheets(target_dir, out_dir / "target", "target", thumb_width, pages_per_sheet)
    template_sheets: list[Path] = []
    if template_dir.exists():
        template_sheets = _contact_sheets(template_dir, out_dir / "template", "template", thumb_width, pages_per_sheet)

    key_pages = _copy_key_pages(audit_dir, target_dir, out_dir / "key-pages")
    prompt = out_dir / "vision_review_prompt.md"
    prompt.write_text(_prompt_text(audit_dir, target_sheets, template_sheets, key_pages), encoding="utf-8")
    return VisionPack(
        audit_dir=audit_dir,
        out_dir=out_dir,
        target_sheets=target_sheets,
        template_sheets=template_sheets,
        key_pages=key_pages,
        prompt=prompt,
    )


def _contact_sheets(image_dir: Path, out_dir: Path, prefix: str, thumb_width: int, pages_per_sheet: int) -> list[Path]:
    try:
        from PIL import Image, ImageDraw, ImageFont
    except ImportError as exc:
        raise RuntimeError("Pillow is required for vision packs") from exc

    pages = sorted(image_dir.glob("page-*.png"), key=_page_sort_key)
    if not pages:
        raise FileNotFoundError(f"No page PNG files found in {image_dir}")
    out_dir.mkdir(parents=True, exist_ok=True)
    columns = 2
    rows = math.ceil(pages_per_sheet / columns)
    label_h = 34
    gap = 18
    margin = 18
    sheets: list[Path] = []
    font = ImageFont.load_default()

    for sheet_idx, start in enumerate(range(0, len(pages), pages_per_sheet), start=1):
        chunk = pages[start:start + pages_per_sheet]
        thumbs: list[tuple[Path, Image.Image]] = []
        for page in chunk:
            image = Image.open(page).convert("RGB")
            scale = thumb_width / image.width
            thumb_h = int(image.height * scale)
            image = image.resize((thumb_width, thumb_h))
            thumbs.append((page, image))

        cell_w = thumb_width
        cell_h = max(image.height for _, image in thumbs) + label_h
        sheet_w = margin * 2 + columns * cell_w + (columns - 1) * gap
        sheet_h = margin * 2 + rows * cell_h + (rows - 1) * gap
        sheet = Image.new("RGB", (sheet_w, sheet_h), "white")
        draw = ImageDraw.Draw(sheet)

        for idx, (page, thumb) in enumerate(thumbs):
            row = idx // columns
            col = idx % columns
            x = margin + col * (cell_w + gap)
            y = margin + row * (cell_h + gap)
            page_no = _page_no(page)
            label = f"{prefix} page {page_no:02d}"
            draw.rectangle((x, y, x + cell_w, y + label_h - 1), fill=(245, 245, 245), outline=(210, 210, 210))
            draw.text((x + 10, y + 10), label, fill=(0, 0, 0), font=font)
            sheet.paste(thumb, (x, y + label_h))
            draw.rectangle((x, y + label_h, x + thumb.width, y + label_h + thumb.height), outline=(180, 180, 180))

        output = out_dir / f"{prefix}-sheet-{sheet_idx:02d}.png"
        sheet.save(output)
        sheets.append(output)
        for _, thumb in thumbs:
            thumb.close()
    return sheets


def _copy_key_pages(audit_dir: Path, target_dir: Path, out_dir: Path) -> list[Path]:
    report = audit_dir / "report.json"
    pages: set[int] = {1, 2, 3, 4, 5}
    if report.exists():
        data = json.loads(report.read_text(encoding="utf-8"))
        visual = data.get("target_visual", {})
        pages.update(int(page) for page in visual.get("blank_pages", []) if str(page).isdigit())
        pages.update(int(page) for page in visual.get("near_blank_pages", []) if str(page).isdigit())
        pages.update(int(page) for page in visual.get("caption_orphan_pages", []) if str(page).isdigit())
        for found_pages in visual.get("heading_pages", {}).values():
            pages.update(int(page) for page in found_pages if str(page).isdigit())
    out_dir.mkdir(parents=True, exist_ok=True)
    copied: list[Path] = []
    for page in sorted(pages):
        src = target_dir / f"page-{page:02d}.png"
        if not src.exists():
            src = target_dir / f"page-{page}.png"
        if not src.exists():
            continue
        dst = out_dir / src.name
        shutil.copy2(src, dst)
        copied.append(dst)
    return copied


def _prompt_text(audit_dir: Path, target_sheets: list[Path], template_sheets: list[Path], key_pages: list[Path]) -> str:
    lines = [
        "# Vision Review Prompt",
        "",
        "你是本科毕业论文格式与质量审阅员。请基于渲染页面图片做视觉级审阅，不要只依赖文字抽取。",
        "",
        "## 审阅目标",
        "",
        "- 对照模板页面，检查目标论文的页边距、字体观感、字号层级、行距、段距、页眉、页脚、页码、目录、正文标题、图表题注、参考文献格式。",
        "- 标出所有明显视觉不一致：空白页、页码体系错误、目录断页异常、标题缩进或字号不一致、图题/表题与对象跨页分离、图表过宽、截图模糊、参考文献排版混乱。",
        "- 内容上按二本/应用型本科工科毕业设计合格线审阅：结构完整、需求分析明确、方案设计可落地、实现与测试充分、结论具体、参考文献规范。",
        "",
        "## 必须输出",
        "",
        "1. `pass/fail` 总结，不能只说大体可以。",
        "2. 按页码列出格式问题和证据。",
        "3. 按章节列出内容质量问题和需要补强的段落。",
        "4. 给出下一轮自动修复应处理的具体动作。",
        "",
        "## 审计目录",
        "",
        f"- `{audit_dir}`",
        "",
        "## Template Sheets",
        "",
    ]
    lines.extend(f"- `{path}`" for path in template_sheets)
    lines.extend(["", "## Target Sheets", ""])
    lines.extend(f"- `{path}`" for path in target_sheets)
    lines.extend(["", "## Key Target Pages", ""])
    lines.extend(f"- `{path}`" for path in key_pages)
    return "\n".join(lines) + "\n"


def _page_sort_key(path: Path) -> tuple[int, str]:
    return (_page_no(path), path.name)


def _page_no(path: Path) -> int:
    match = re.search(r"page-(\d+)\.png$", path.name)
    return int(match.group(1)) if match else 0
