from __future__ import annotations

import re
import subprocess
from dataclasses import dataclass, field
from pathlib import Path

from .tools import Toolchain, run_cmd


@dataclass(frozen=True)
class PdfInfo:
    pages: int
    page_size: str | None
    raw: dict[str, str] = field(default_factory=dict)


def pdf_info(pdf_path: Path, toolchain: Toolchain) -> PdfInfo:
    if not toolchain.pdfinfo:
        raise RuntimeError("pdfinfo not found")
    result = run_cmd([toolchain.pdfinfo, str(pdf_path)])
    raw: dict[str, str] = {}
    for line in result.stdout.splitlines():
        if ":" in line:
            key, value = line.split(":", 1)
            raw[key.strip()] = value.strip()
    return PdfInfo(
        pages=int(raw.get("Pages", "0") or "0"),
        page_size=raw.get("Page size"),
        raw=raw,
    )


def extract_text(pdf_path: Path, toolchain: Toolchain, first_page: int | None = None, last_page: int | None = None) -> str:
    if not toolchain.pdftotext:
        raise RuntimeError("pdftotext not found")
    args = [toolchain.pdftotext, "-layout"]
    if first_page is not None:
        args.extend(["-f", str(first_page)])
    if last_page is not None:
        args.extend(["-l", str(last_page)])
    args.extend([str(pdf_path), "-"])
    try:
        return run_cmd(args).stdout
    except subprocess.CalledProcessError as exc:
        detail = (exc.stderr or exc.stdout or "").strip()
        raise RuntimeError(f"pdftotext failed for {pdf_path}: {detail}") from exc


def extract_page_texts(pdf_path: Path, toolchain: Toolchain, pages: int) -> dict[int, str]:
    return {page: extract_text(pdf_path, toolchain, page, page) for page in range(1, pages + 1)}


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", "", text)
