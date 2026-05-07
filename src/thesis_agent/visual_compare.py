from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field
from pathlib import Path

from PIL import Image, ImageChops

from .config import AgentConfig
from .render import render_document
from .tools import Toolchain

VISUAL_MEAN_TOLERANCE = 0.001
VISUAL_CHANGED_PIXEL_TOLERANCE = 0.005


@dataclass(frozen=True)
class PageVisualDiff:
    page: int
    expected: Path
    actual: Path
    mean_abs_diff: float
    changed_pixel_ratio: float
    max_channel_diff: int
    diff_image: Path | None = None

    @property
    def strict_passed(self) -> bool:
        return self.mean_abs_diff == 0 and self.changed_pixel_ratio == 0 and self.max_channel_diff == 0

    @property
    def passed(self) -> bool:
        return (
            self.mean_abs_diff <= VISUAL_MEAN_TOLERANCE
            and self.changed_pixel_ratio <= VISUAL_CHANGED_PIXEL_TOLERANCE
        )


@dataclass(frozen=True)
class TemplateVisualCompareReport:
    cover: Path
    body: Path
    rebuilt: Path
    out_dir: Path
    expected_pages: int
    actual_pages: int
    passed: bool
    page_diffs: list[PageVisualDiff] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    def to_json(self) -> str:
        return json.dumps(_jsonable(asdict(self)), ensure_ascii=False, indent=2)


def compare_standard_template_visual(
    cover_template: Path,
    body_template: Path,
    rebuilt_template: Path,
    out_dir: Path,
    config: AgentConfig,
    toolchain: Toolchain,
) -> TemplateVisualCompareReport:
    """Render cover/body originals and the assembled template, then compare pixels page by page."""
    out_dir.mkdir(parents=True, exist_ok=True)
    pdf_dir = out_dir / "pdf"
    png_dir = out_dir / "png"
    diff_dir = out_dir / "diff"
    diff_dir.mkdir(parents=True, exist_ok=True)

    rendered_cover = render_document(cover_template, pdf_dir, png_dir / "cover", toolchain, config.renderer_dpi)
    rendered_body = render_document(body_template, pdf_dir, png_dir / "body", toolchain, config.renderer_dpi)
    rendered_rebuilt = render_document(rebuilt_template, pdf_dir, png_dir / "rebuilt", toolchain, config.renderer_dpi)

    expected_pages = len(rendered_cover.pages) + len(rendered_body.pages)
    actual_pages = len(rendered_rebuilt.pages)
    warnings: list[str] = []
    if expected_pages != actual_pages:
        warnings.append(f"页数不一致：期望 {expected_pages} 页，实际 {actual_pages} 页。")

    expected = list(rendered_cover.pages) + list(rendered_body.pages)
    paired_count = min(len(expected), len(rendered_rebuilt.pages))
    page_diffs = [
        _compare_page(page, expected[page - 1], rendered_rebuilt.pages[page - 1], diff_dir)
        for page in range(1, paired_count + 1)
    ]
    passed = expected_pages == actual_pages and all(item.passed for item in page_diffs)
    report = TemplateVisualCompareReport(
        cover=cover_template,
        body=body_template,
        rebuilt=rebuilt_template,
        out_dir=out_dir,
        expected_pages=expected_pages,
        actual_pages=actual_pages,
        passed=passed,
        page_diffs=page_diffs,
        warnings=warnings,
    )
    (out_dir / "report.json").write_text(report.to_json() + "\n", encoding="utf-8")
    (out_dir / "report.md").write_text(_markdown(report), encoding="utf-8")
    return report


def _compare_page(page: int, expected: Path, actual: Path, diff_dir: Path) -> PageVisualDiff:
    with Image.open(expected).convert("RGB") as expected_image, Image.open(actual).convert("RGB") as actual_image:
        if expected_image.size != actual_image.size:
            diff_path = diff_dir / f"page-{page:02d}.png"
            _size_mismatch_image(expected_image, actual_image).save(diff_path)
            return PageVisualDiff(
                page=page,
                expected=expected,
                actual=actual,
                mean_abs_diff=1.0,
                changed_pixel_ratio=1.0,
                max_channel_diff=255,
                diff_image=diff_path,
            )
        diff = ImageChops.difference(expected_image, actual_image)
        pixels = list(diff.getdata())
        total = max(1, len(pixels))
        channel_sum = sum(r + g + b for r, g, b in pixels)
        changed = sum(1 for r, g, b in pixels if r or g or b)
        max_channel = max((max(pixel) for pixel in pixels), default=0)
        diff_path: Path | None = None
        if changed:
            diff_path = diff_dir / f"page-{page:02d}.png"
            amplified = diff.point(lambda value: min(255, value * 8))
            amplified.save(diff_path)
        return PageVisualDiff(
            page=page,
            expected=expected,
            actual=actual,
            mean_abs_diff=channel_sum / (total * 3 * 255),
            changed_pixel_ratio=changed / total,
            max_channel_diff=max_channel,
            diff_image=diff_path,
        )


def _size_mismatch_image(expected: Image.Image, actual: Image.Image) -> Image.Image:
    width = expected.width + actual.width
    height = max(expected.height, actual.height)
    canvas = Image.new("RGB", (width, height), "white")
    canvas.paste(expected, (0, 0))
    canvas.paste(actual, (expected.width, 0))
    return canvas


def _markdown(report: TemplateVisualCompareReport) -> str:
    lines = [
        "# Template Visual Compare",
        "",
        f"- Cover: `{report.cover}`",
        f"- Body: `{report.body}`",
        f"- Rebuilt: `{report.rebuilt}`",
        f"- Expected pages: {report.expected_pages}",
        f"- Actual pages: {report.actual_pages}",
        f"- Passed: `{str(report.passed).lower()}`",
        f"- Visual tolerance: mean <= {VISUAL_MEAN_TOLERANCE}, changed pixels <= {VISUAL_CHANGED_PIXEL_TOLERANCE}",
    ]
    if report.warnings:
        lines.append("- Warnings: " + "; ".join(report.warnings))
    lines.extend(["", "## Page Diffs", ""])
    for item in report.page_diffs:
        status = "PASS" if item.passed else "FAIL"
        strict = "exact" if item.strict_passed else "within tolerance" if item.passed else "different"
        diff = f", diff: `{item.diff_image}`" if item.diff_image else ""
        lines.append(
            f"- Page {item.page}: {status} ({strict}), mean={item.mean_abs_diff:.8f}, "
            f"changed={item.changed_pixel_ratio:.8f}, max={item.max_channel_diff}{diff}"
        )
    return "\n".join(lines) + "\n"


def _jsonable(value):
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, dict):
        return {str(key): _jsonable(item) for key, item in value.items()}
    if isinstance(value, list):
        return [_jsonable(item) for item in value]
    return value
