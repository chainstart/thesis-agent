from __future__ import annotations

import shutil
import subprocess
from dataclasses import dataclass
from pathlib import Path


DEFAULT_OFFICECLI = Path("/home/biostar/work/external/bin/officecli")


@dataclass(frozen=True)
class Toolchain:
    soffice: str | None
    pdftoppm: str | None
    pdftotext: str | None
    pdfinfo: str | None
    officecli: str | None

    @classmethod
    def discover(cls) -> "Toolchain":
        officecli = shutil.which("officecli")
        if officecli is None and DEFAULT_OFFICECLI.exists():
            officecli = str(DEFAULT_OFFICECLI)
        return cls(
            soffice=shutil.which("soffice-headless") or shutil.which("soffice") or shutil.which("libreoffice"),
            pdftoppm=shutil.which("pdftoppm"),
            pdftotext=shutil.which("pdftotext"),
            pdfinfo=shutil.which("pdfinfo"),
            officecli=officecli,
        )

    def missing_for_render(self) -> list[str]:
        missing: list[str] = []
        if not self.soffice:
            missing.append("soffice-headless|soffice|libreoffice")
        if not self.pdftoppm:
            missing.append("pdftoppm")
        if not self.pdftotext:
            missing.append("pdftotext")
        if not self.pdfinfo:
            missing.append("pdfinfo")
        return missing

    def as_dict(self) -> dict[str, str | None]:
        return {
            "soffice": self.soffice,
            "pdftoppm": self.pdftoppm,
            "pdftotext": self.pdftotext,
            "pdfinfo": self.pdfinfo,
            "officecli": self.officecli,
        }


def run_cmd(args: list[str], cwd: Path | None = None) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        args,
        cwd=str(cwd) if cwd else None,
        text=True,
        capture_output=True,
        check=True,
    )


def probe_font(name: str) -> str:
    fc_match = shutil.which("fc-match")
    if not fc_match:
        return "fc-match not found"
    try:
        result = run_cmd([fc_match, name])
    except subprocess.CalledProcessError as exc:
        return (exc.stderr or exc.stdout or str(exc)).strip()
    return result.stdout.strip()


def command_version(command: str | None, args: list[str]) -> str | None:
    if not command:
        return None
    try:
        result = run_cmd([command, *args])
    except Exception:
        return None
    return (result.stdout or result.stderr).strip().splitlines()[0] if (result.stdout or result.stderr) else ""
