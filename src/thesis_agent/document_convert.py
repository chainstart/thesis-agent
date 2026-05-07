from __future__ import annotations

import shutil
import subprocess
from dataclasses import dataclass
from pathlib import Path

from .tools import Toolchain, run_cmd


@dataclass(frozen=True)
class PreparedDocument:
    original: Path
    docx: Path
    converted: bool


def ensure_docx(input_path: Path, out_dir: Path, toolchain: Toolchain) -> PreparedDocument:
    if input_path.suffix.lower() == ".docx":
        out_dir.mkdir(parents=True, exist_ok=True)
        target = out_dir / input_path.name
        shutil.copy2(input_path, target)
        return PreparedDocument(original=input_path, docx=target, converted=False)
    if input_path.suffix.lower() != ".doc":
        raise ValueError(f"Unsupported thesis file type: {input_path.suffix}")
    if not toolchain.soffice:
        raise RuntimeError("LibreOffice executable not found; cannot convert .doc to .docx")

    out_dir.mkdir(parents=True, exist_ok=True)
    staged = out_dir / input_path.name
    shutil.copy2(input_path, staged)
    expected = out_dir / f"{staged.stem}.docx"
    if expected.exists():
        expected.unlink()
    try:
        run_cmd([
            toolchain.soffice,
            "--headless",
            "--convert-to",
            "docx",
            "--outdir",
            str(out_dir),
            str(staged),
        ])
    except subprocess.CalledProcessError as exc:
        detail = (exc.stderr or exc.stdout or "").strip()
        raise RuntimeError(f"LibreOffice failed to convert {input_path} to docx: {detail}") from exc
    if not expected.exists():
        candidates = sorted(out_dir.glob(f"{staged.stem}*.docx"), key=lambda p: p.stat().st_mtime, reverse=True)
        if not candidates:
            raise RuntimeError(f"LibreOffice did not create a docx for {input_path}")
        expected = candidates[0]
    return PreparedDocument(original=input_path, docx=expected, converted=True)
