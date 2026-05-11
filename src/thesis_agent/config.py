from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any


DEFAULT_CONFIG = Path(__file__).resolve().parents[2] / "configs" / "default_format.json"


@dataclass(frozen=True)
class AgentConfig:
    path: Path
    data: dict[str, Any]

    @classmethod
    def load(cls, path: Path | None = None) -> "AgentConfig":
        config_path = path or DEFAULT_CONFIG
        with config_path.open("r", encoding="utf-8") as fh:
            data = json.load(fh)
        return cls(path=config_path, data=data)

    @property
    def renderer_dpi(self) -> int:
        return int(self.data.get("renderer", {}).get("dpi", 120))

    @property
    def blank_ink_ratio(self) -> float:
        return float(self.data.get("renderer", {}).get("blank_ink_ratio", 0.001))

    @property
    def near_blank_ink_ratio(self) -> float:
        return float(self.data.get("renderer", {}).get("near_blank_ink_ratio", 0.006))

    @property
    def main_heading_patterns(self) -> list[str]:
        return list(self.data.get("main_heading_patterns", []))

    @property
    def expected_sections(self) -> list[str]:
        return list(self.data.get("expected_sections", []))

    @property
    def content_quality(self) -> dict[str, Any]:
        return dict(self.data.get("content_quality", {}))
