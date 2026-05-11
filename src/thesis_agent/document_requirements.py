from __future__ import annotations

import io
import posixpath
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET

from PIL import Image

from .docx_inspect import NS, W_NS, _paragraph_text
from .document_profile import MIDTERM, PROPOSAL, THESIS


R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
REL = {"rel": REL_NS}
R = {"r": R_NS}


@dataclass(frozen=True)
class RequirementInspection:
    required_field_errors: list[str] = field(default_factory=list)
    signature_errors: list[str] = field(default_factory=list)
    opinion_errors: list[str] = field(default_factory=list)
    metadata_warnings: list[str] = field(default_factory=list)

    @property
    def errors(self) -> list[str]:
        return [*self.required_field_errors, *self.signature_errors, *self.opinion_errors]


def inspect_document_requirements(path: Path, kind: str, metadata: dict[str, str] | None = None) -> RequirementInspection:
    metadata = metadata or {}
    if path.suffix.lower() != ".docx" or not zipfile.is_zipfile(path):
        return RequirementInspection(metadata_warnings=["非 docx 文件无法做结构化材料门禁检查。"])
    with zipfile.ZipFile(path) as zf:
        try:
            root = ET.fromstring(zf.read("word/document.xml"))
        except KeyError:
            return RequirementInspection(metadata_warnings=["word/document.xml 不存在，无法做结构化材料门禁检查。"])
        rels = _document_relationships(zf)
        media = {name: zf.read(name) for name in zf.namelist() if name.startswith("word/media/")}

    return RequirementInspection(
        required_field_errors=_required_field_errors(root, kind),
        signature_errors=_signature_errors(root, rels, media),
        opinion_errors=_opinion_errors(root, kind),
        metadata_warnings=_metadata_warnings(metadata),
    )


def _required_field_errors(root: ET.Element, kind: str) -> list[str]:
    labels = ["姓名", "学号", "专业", "班级", "课题名称", "题目", "指导教师"]
    if kind == THESIS:
        labels.extend(["学院", "年级", "论文题目"])
    errors: list[str] = []
    seen: set[str] = set()
    for text in _paragraph_and_cell_texts(root):
        compact = re.sub(r"\s+", "", text)
        for label in labels:
            if label in seen or label not in compact:
                continue
            value = _value_after_label(text, label)
            if value is None:
                continue
            seen.add(label)
            if _is_blank_value(value):
                errors.append(f"{label}: 信息应填写完整，不能留空或只保留下划线。")
    return errors


def _signature_errors(root: ET.Element, rels: dict[str, str], media: dict[str, bytes]) -> list[str]:
    errors: list[str] = []
    checked_cells: set[int] = set()
    for cell in root.findall(".//w:tc", NS):
        text = _element_text(cell)
        if not _looks_like_signature_field(text):
            continue
        checked_cells.add(id(cell))
        images = _image_parts(cell, rels)
        label = _short(text)
        if not images:
            errors.append(f"{label}: 需要签名图片，不能留空。")
            continue
        if not any(_image_has_transparent_signature(media.get(part, b"")) for part in images):
            errors.append(f"{label}: 签名图片必须为透明背景，只保留姓名笔迹。")

    body = root.find("w:body", NS)
    children = list(body) if body is not None else []
    for idx, paragraph in enumerate(children):
        if paragraph.tag != f"{{{W_NS}}}p":
            continue
        text = _paragraph_text(paragraph)
        if not _looks_like_signature_field(text):
            continue
        images: list[str] = []
        for candidate in children[idx : idx + 3]:
            if candidate.tag == f"{{{W_NS}}}p":
                images.extend(_image_parts(candidate, rels))
        if not images:
            errors.append(f"{_short(text)}: 需要签名图片，不能留空。")
        elif not any(_image_has_transparent_signature(media.get(part, b"")) for part in images):
            errors.append(f"{_short(text)}: 签名图片必须为透明背景，只保留姓名笔迹。")
    return _dedupe(errors)


def _opinion_errors(root: ET.Element, kind: str) -> list[str]:
    if kind not in {PROPOSAL, MIDTERM}:
        return []
    texts = _paragraph_and_cell_texts(root)
    opinion_texts = [
        text
        for text in texts
        if re.search(r"(指导|答辩|评阅|检查|小组).{0,10}(意见|评语|结论)", re.sub(r"\s+", "", text))
    ]
    if not opinion_texts:
        label = "开题报告" if kind == PROPOSAL else "中期检查报告"
        return [f"{label}: 末尾指导教师意见、答辩/检查意见等意见区缺失或未识别。"]
    errors: list[str] = []
    for text in opinion_texts:
        content = _opinion_content(text)
        if len(re.findall(r"[\u4e00-\u9fff]", content)) < 12:
            errors.append(f"{_short(text)}: 意见内容应填写完整，不能留空或只写日期/签名。")
    return _dedupe(errors)


def _metadata_warnings(metadata: dict[str, str]) -> list[str]:
    warnings: list[str] = []
    if not metadata.get("student_id") and not metadata.get("student_name"):
        warnings.append("未识别到学号或姓名，无法写入全局学生信息表主键。")
    return warnings


def _looks_like_signature_field(text: str) -> bool:
    compact = re.sub(r"\s+", "", text)
    if "签名" not in compact and "签字" not in compact:
        return False
    if "数字签名" in compact:
        return False
    if re.search(r"(学生|本人|指导教师|教师|组长|成员|答辩小组|负责人|签收).{0,12}(签名|签字)", compact):
        return True
    if re.search(r"(签名|签字)[：:年月日_＿\u00a0\s]*$", text):
        return True
    return len(compact) <= 40 and re.search(r"(签名|签字)", compact) is not None


def _paragraph_and_cell_texts(root: ET.Element) -> list[str]:
    texts = []
    for cell in root.findall(".//w:tc", NS):
        text = _element_text(cell).strip()
        if text:
            texts.append(text)
    for paragraph in root.findall(".//w:p", NS):
        text = _paragraph_text(paragraph).strip()
        if text:
            texts.append(text)
    return texts


def _element_text(element: ET.Element) -> str:
    return "\n".join(_paragraph_text(paragraph) for paragraph in element.findall(".//w:p", NS) if _paragraph_text(paragraph))


def _value_after_label(text: str, label: str) -> str | None:
    compact_label = re.sub(r"\s+", "", label)
    for line in text.splitlines() or [text]:
        compact = re.sub(r"\s+", "", line)
        pos = compact.find(compact_label)
        if pos < 0:
            continue
        match = re.search(r"[：:]", line)
        if match is not None:
            return line[match.end() :]
        tail = re.sub(r"\s+", "", line)[pos + len(compact_label) :]
        return tail
    return None


def _is_blank_value(value: str) -> bool:
    cleaned = re.sub(r"[\s\u00a0_＿—－\-年月日/]+", "", value)
    return not cleaned or cleaned in {"待填", "填写", "无"}


def _opinion_content(text: str) -> str:
    compact = re.sub(r"\s+", "", text)
    compact = re.sub(r"^.*?(意见|评语|结论)[：:]?", "", compact)
    compact = re.sub(r"(签名|日期|年月日).*$", "", compact)
    return compact


def _document_relationships(zf: zipfile.ZipFile) -> dict[str, str]:
    try:
        rels_root = ET.fromstring(zf.read("word/_rels/document.xml.rels"))
    except KeyError:
        return {}
    result: dict[str, str] = {}
    for rel in rels_root.findall("rel:Relationship", REL):
        rid = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        if not rid or not target:
            continue
        if target.startswith("/"):
            part = target.lstrip("/")
        else:
            part = posixpath.normpath(posixpath.join("word", target))
        result[rid] = part
    return result


def _image_parts(element: ET.Element, rels: dict[str, str]) -> list[str]:
    parts: list[str] = []
    for node in element.findall(".//*[@r:embed]", {**NS, **R}):
        rid = node.attrib.get(f"{{{R_NS}}}embed")
        if rid and rels.get(rid, "").startswith("word/media/"):
            parts.append(rels[rid])
    for node in element.findall(".//*[@r:id]", {**NS, **R}):
        rid = node.attrib.get(f"{{{R_NS}}}id")
        if rid and rels.get(rid, "").startswith("word/media/"):
            parts.append(rels[rid])
    return _dedupe(parts)


def _image_has_transparent_signature(data: bytes) -> bool:
    if not data:
        return False
    try:
        image = Image.open(io.BytesIO(data)).convert("RGBA")
    except Exception:
        return False
    alpha = image.getchannel("A")
    values = list(alpha.getdata())
    if not values:
        return False
    transparent = sum(1 for value in values if value < 245)
    opaque = sum(1 for value in values if value > 10)
    total = len(values)
    return transparent / total >= 0.15 and opaque >= 20


def _short(text: str, limit: int = 40) -> str:
    compact = re.sub(r"\s+", "", text)
    return compact[:limit] + ("..." if len(compact) > limit else "")


def _dedupe(items: list[str]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for item in items:
        if item in seen:
            continue
        seen.add(item)
        result.append(item)
    return result
