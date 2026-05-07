from __future__ import annotations

import json
import re
import zipfile
from dataclasses import asdict, dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET

from .docx_inspect import NS, W_NS, _paragraph_text
from .format_fix import _resolve_template_docx


@dataclass(frozen=True)
class TemplateProfile:
    source: Path
    docx_source: Path | None
    front_matter_titles: list[str] = field(default_factory=list)
    required_sections: list[str] = field(default_factory=list)
    heading_style_ids: dict[str, str] = field(default_factory=dict)
    main_heading_patterns: list[str] = field(default_factory=list)
    has_front_matter: bool = False

    def to_json(self) -> str:
        data = asdict(self)
        data["source"] = str(self.source)
        data["docx_source"] = str(self.docx_source) if self.docx_source else None
        return json.dumps(data, ensure_ascii=False, indent=2)


def build_template_profile(template_path: Path) -> TemplateProfile:
    docx = _resolve_template_docx(template_path)
    front_titles: list[str] = []
    required_sections: list[str] = []
    heading_styles: dict[str, str] = {}
    patterns = [
        r"^1\s+.+",
        r"^2\s+.+",
        r"^3\s+.+",
        r"^4\s+.+",
        r"^5\s+.+",
        r"^参考文献",
    ]

    if docx and zipfile.is_zipfile(docx):
        with zipfile.ZipFile(docx) as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
            texts = [_paragraph_text(p) for p in root.findall(".//w:p", NS)]
            compact = "\n".join(texts)
            front_titles = _front_matter_titles(compact)
            required_sections = _required_sections(compact)
            if "word/styles.xml" in zf.namelist():
                styles = ET.fromstring(zf.read("word/styles.xml"))
                heading_styles = _heading_style_ids(styles)

    return TemplateProfile(
        source=template_path,
        docx_source=docx,
        front_matter_titles=front_titles,
        required_sections=required_sections,
        heading_style_ids=heading_styles,
        main_heading_patterns=patterns,
        has_front_matter=bool(front_titles),
    )


def _front_matter_titles(text: str) -> list[str]:
    titles = []
    for title in ["毕业设计（论文）学术诚信声明", "毕业设计（论文）AI使用情况声明", "毕业设计（论文）版权使用授权书"]:
        if re.sub(r"\s+", "", title) in re.sub(r"\s+", "", text):
            titles.append(title)
    return titles


def _required_sections(text: str) -> list[str]:
    compact = re.sub(r"\s+", "", text).lower()
    sections: list[str] = []
    for section in ["摘 要", "Abstract", "目 录", "参考文献"]:
        if re.sub(r"\s+", "", section).lower() in compact:
            sections.append(section)
    return sections


def _heading_style_ids(styles: ET.Element) -> dict[str, str]:
    result: dict[str, str] = {}
    for style in styles.findall("w:style", NS):
        if style.attrib.get(f"{{{W_NS}}}type") != "paragraph":
            continue
        name = style.find("w:name", NS)
        value = name.attrib.get(f"{{{W_NS}}}val", "") if name is not None else ""
        if value.lower().startswith("heading"):
            result[value] = style.attrib.get(f"{{{W_NS}}}styleId", "")
    return result
