from __future__ import annotations

import re
import shutil
import subprocess
from dataclasses import dataclass
from pathlib import Path

from .tools import Toolchain, run_cmd


@dataclass(frozen=True)
class RenderedDocument:
    source: Path
    pdf: Path
    png_dir: Path
    pages: list[Path]


def render_to_pdf(input_path: Path, out_dir: Path, toolchain: Toolchain) -> Path:
    if not toolchain.soffice:
        raise RuntimeError("LibreOffice executable not found")
    out_dir.mkdir(parents=True, exist_ok=True)
    expected_pdf = out_dir / f"{input_path.stem}.pdf"
    if expected_pdf.exists():
        expected_pdf.unlink()
    try:
        run_cmd(
            [
                toolchain.soffice,
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(out_dir),
                str(input_path),
            ]
        )
    except subprocess.CalledProcessError as exc:
        detail = (exc.stderr or exc.stdout or "").strip()
        raise RuntimeError(f"LibreOffice failed to render {input_path}: {detail}") from exc
    if expected_pdf.exists():
        return expected_pdf
    candidates = sorted(out_dir.glob(f"{input_path.stem}*.pdf"), key=lambda p: p.stat().st_mtime, reverse=True)
    if candidates:
        return candidates[0]
    raise RuntimeError(f"LibreOffice did not create a PDF for {input_path}")


def render_pdf_to_png(pdf_path: Path, out_dir: Path, toolchain: Toolchain, dpi: int = 120) -> list[Path]:
    if not toolchain.pdftoppm:
        raise RuntimeError("pdftoppm not found")
    if out_dir.exists():
        shutil.rmtree(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    prefix = out_dir / "page"
    try:
        run_cmd([toolchain.pdftoppm, "-r", str(dpi), "-png", str(pdf_path), str(prefix)])
    except subprocess.CalledProcessError as exc:
        detail = (exc.stderr or exc.stdout or "").strip()
        raise RuntimeError(f"pdftoppm failed for {pdf_path}: {detail}") from exc
    return sorted(out_dir.glob("page-*.png"), key=_page_sort_key)


def render_document(input_path: Path, pdf_dir: Path, png_dir: Path, toolchain: Toolchain, dpi: int) -> RenderedDocument:
    pdf = render_to_pdf(input_path, pdf_dir, toolchain)
    pages = render_pdf_to_png(pdf, png_dir, toolchain, dpi=dpi)
    return RenderedDocument(source=input_path, pdf=pdf, png_dir=png_dir, pages=pages)


def _page_sort_key(path: Path) -> tuple[int, str]:
    match = re.search(r"-(\d+)\.png$", path.name)
    return (int(match.group(1)) if match else 0, path.name)
