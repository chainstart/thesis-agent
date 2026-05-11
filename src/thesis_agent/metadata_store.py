from __future__ import annotations

import json
import re
import shutil
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from .docx_inspect import NS, W_NS, _paragraph_text
from .ooxml import serialize_xml


FIELD_LABELS = {
    "学院": "college",
    "院系": "college",
    "专业": "major",
    "班级": "class_name",
    "年级": "grade",
    "学生姓名": "student_name",
    "姓名": "student_name",
    "学生学号": "student_id",
    "学号": "student_id",
    "指导教师": "advisor",
    "教师": "advisor",
    "论文题目": "title",
    "毕业设计（论文）题目": "title",
    "毕业设计题目": "title",
    "课题名称": "title",
    "题目": "title",
}


def default_metadata_store(out_dir: Path) -> Path:
    return out_dir / "student_metadata.json"


def extract_document_metadata(path: Path, text: str = "") -> dict[str, str]:
    metadata: dict[str, str] = {}
    source = f"{path.stem}\n{text[:5000]}"
    student_id_match = re.search(r"(?<!\d)(\d{12})(?!\d)", source)
    if student_id_match:
        metadata["student_id"] = student_id_match.group(1)

    class_match = re.search(r"(物联网|网络)\s*(\d{4})", source)
    if class_match:
        metadata["class_name"] = f"{class_match.group(1)}{class_match.group(2)}"
        metadata.setdefault("major", "物联网工程" if class_match.group(1) == "物联网" else "网络工程")
        metadata.setdefault("grade", f"20{class_match.group(2)[:2]}级")

    id_name = re.search(r"(?<!\d)(\d{12})(?!\d)[-_ 　]*([\u4e00-\u9fff]{2,4})", source)
    if id_name:
        metadata.setdefault("student_name", id_name.group(2))
    else:
        class_name = re.search(r"(?:物联网|网络)\s*\d{4}[-_ 　]+(?:\d{12}[-_ 　]+)?([\u4e00-\u9fff]{2,4})", path.stem)
        if class_name:
            metadata.setdefault("student_name", class_name.group(1))
        else:
            for part in re.split(r"[-_ 　\s]+", path.stem):
                if re.fullmatch(r"[\u4e00-\u9fff]{2,4}", part) and part not in {"论文", "开题", "任务书", "中期", "检查", "报告"}:
                    metadata.setdefault("student_name", part)
                    break

    for line in [item.strip() for item in text.splitlines() if item.strip()][:120]:
        label_match = re.match(r"^(.{1,16}?)[：:]\s*(.+)$", line)
        if not label_match:
            continue
        label = re.sub(r"\s+", "", label_match.group(1))
        key = FIELD_LABELS.get(label)
        if key is None:
            continue
        value = _clean_field_value(label_match.group(2))
        if value:
            metadata[key] = value

    title = _title_from_filename(path) or _title_from_text(text)
    if title:
        metadata.setdefault("title", title)
    return {key: value for key, value in metadata.items() if value}


def update_metadata_store(store_path: Path, metadata: dict[str, str], source: Path) -> dict[str, str]:
    if not metadata:
        return {}
    store = _load_store(store_path)
    records = store.setdefault("records", {})
    key = metadata_key(metadata)
    if key is None:
        return metadata
    existing = dict(records.get(key, {}))
    source_mtime = source.stat().st_mtime if source.exists() else 0.0
    existing_mtime = float(existing.get("source_mtime", 0.0) or 0.0)
    merged = {k: v for k, v in existing.items() if isinstance(v, str) and v}
    if source_mtime >= existing_mtime:
        merged.update(metadata)
    else:
        for field, value in metadata.items():
            merged.setdefault(field, value)
    merged["source"] = str(source)
    merged["source_mtime"] = source_mtime
    records[key] = merged
    if metadata.get("student_id") and metadata.get("student_name"):
        records[f"name:{metadata['student_name']}"] = merged
    store_path.parent.mkdir(parents=True, exist_ok=True)
    store_path.write_text(json.dumps(store, ensure_ascii=False, indent=2), encoding="utf-8")
    return {key: value for key, value in merged.items() if isinstance(value, str) and key not in {"source"}}


def lookup_metadata(store_path: Path, metadata: dict[str, str]) -> dict[str, str]:
    store = _load_store(store_path)
    records = store.get("records", {})
    keys = []
    if metadata.get("student_id"):
        keys.append(f"id:{metadata['student_id']}")
    if metadata.get("student_name"):
        keys.append(f"name:{metadata['student_name']}")
    merged: dict[str, str] = {}
    for key in keys:
        record = records.get(key)
        if isinstance(record, dict):
            merged.update({k: v for k, v in record.items() if isinstance(v, str) and v})
    merged.update(metadata)
    return merged


def metadata_key(metadata: dict[str, str]) -> str | None:
    if metadata.get("student_id"):
        return f"id:{metadata['student_id']}"
    if metadata.get("student_name"):
        return f"name:{metadata['student_name']}"
    return None


def apply_metadata_to_docx(input_path: Path, output_path: Path, metadata: dict[str, str]) -> int:
    if input_path.suffix.lower() != ".docx" or not zipfile.is_zipfile(input_path) or not metadata:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(input_path, output_path)
        return 0
    with zipfile.ZipFile(input_path) as zin:
        try:
            root = ET.fromstring(zin.read("word/document.xml"))
        except KeyError:
            output_path.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(input_path, output_path)
            return 0
        filled = 0
        for paragraph in root.findall(".//w:p", NS):
            text = _paragraph_text(paragraph)
            replacement = _filled_field_text(text, metadata)
            if replacement and replacement != text:
                _set_paragraph_text(paragraph, replacement)
                filled += 1
        document_xml = serialize_xml(root)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = document_xml if item.filename == "word/document.xml" else zin.read(item.filename)
                zout.writestr(item, data)
    return filled


def _load_store(path: Path) -> dict[str, object]:
    if not path.exists():
        return {"records": {}}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {"records": {}}
    return payload if isinstance(payload, dict) else {"records": {}}


def _filled_field_text(text: str, metadata: dict[str, str]) -> str | None:
    compact_text = re.sub(r"\s+", "", text)
    for label, key in sorted(FIELD_LABELS.items(), key=lambda item: len(item[0]), reverse=True):
        value = metadata.get(key)
        if not value:
            continue
        compact_label = re.sub(r"\s+", "", label)
        if compact_label not in compact_text:
            continue
        match = re.search(rf"({re.escape(label)}\s*[：:])(.+)?$", text)
        if match is None:
            match = re.search(rf"({re.escape(label)})(.+)?$", text)
        if match is None:
            continue
        tail = match.group(2) or ""
        if not _field_value_is_blank(tail):
            continue
        return text[: match.start()] + match.group(1) + value
    return None


def _field_value_is_blank(value: str) -> bool:
    cleaned = _clean_field_value(value)
    return not cleaned or cleaned in {"待填", "填写", "无", "年月日"}


def _clean_field_value(value: str) -> str:
    return re.sub(r"[\s\u00a0_＿—－-]+", "", value).strip()


def _set_paragraph_text(paragraph: ET.Element, text: str) -> None:
    ppr = paragraph.find("w:pPr", NS)
    preserved_ppr = ET.fromstring(ET.tostring(ppr)) if ppr is not None else None
    first_rpr = paragraph.find("w:r/w:rPr", NS)
    preserved_rpr = ET.fromstring(ET.tostring(first_rpr)) if first_rpr is not None else None
    for child in list(paragraph):
        paragraph.remove(child)
    if preserved_ppr is not None:
        paragraph.append(preserved_ppr)
    run = ET.SubElement(paragraph, f"{{{W_NS}}}r")
    if preserved_rpr is not None:
        run.append(preserved_rpr)
    text_node = ET.SubElement(run, f"{{{W_NS}}}t")
    text_node.text = text
    if re.search(r"^\s|\s$|\s{2,}", text):
        text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def _title_from_filename(path: Path) -> str | None:
    stem = re.sub(r"(毕业设计（论文）|毕业论文|论文|开题报告|任务书|中期检查报告|中期检查|最终稿|初稿|第二版|第三版|批注待改|待修改|修改|改\d*)", "", path.stem)
    parts = [part.strip(" -_　()（）") for part in re.split(r"[-_]", stem) if part.strip(" -_　()（）")]
    for part in reversed(parts):
        if _looks_like_title(part):
            return part
    return None


def _title_from_text(text: str) -> str | None:
    for line in [item.strip() for item in text.splitlines() if item.strip()][:80]:
        if _looks_like_title(line):
            return line
    return None


def _looks_like_title(value: str) -> bool:
    compact = re.sub(r"\s+", "", value)
    if not 8 <= len(compact) <= 45:
        return False
    if any(token in compact for token in ("姓名", "学号", "专业", "班级", "学院", "指导教师", "课题", "课题名称", "论文题目", "开题报告", "任务书")):
        return False
    if re.search(r"\d{4}年|\d+月|完成|初稿|中期检查|进度|时间|安排", compact):
        return False
    return bool(re.search(r"[\u4e00-\u9fff]", compact)) and bool(re.search(r"(系统|平台|设计|研究|开发|监测|追踪|网络|管理)", compact))
