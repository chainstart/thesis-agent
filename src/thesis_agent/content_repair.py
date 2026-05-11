from __future__ import annotations

import json
import re
import zipfile
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from .content_enhance import _apply_language_cleanup, _augment_test_chapter, _insert_acknowledgements
from .docx_inspect import NS, W_NS, _paragraph_text
from .ooxml import serialize_xml


ET.register_namespace("w", W_NS)


@dataclass(frozen=True)
class ContentRepairReport:
    input: Path
    output: Path
    language_fixes_applied: int = 0
    test_chapter_augmented: bool = False
    acknowledgements_inserted: bool = False
    inserted_paragraphs: int = 0
    actions: list[str] = field(default_factory=list)
    skipped: list[str] = field(default_factory=list)

    def to_json(self) -> str:
        return json.dumps(asdict(self), ensure_ascii=False, indent=2, default=str)


def repair_reviewed_content(input_path: Path, output_path: Path, issues: list[Any]) -> ContentRepairReport:
    if input_path.suffix.lower() != ".docx":
        raise ValueError("content repair currently supports .docx targets only")
    if not zipfile.is_zipfile(input_path):
        raise ValueError(f"Not a valid docx file: {input_path}")

    issue_codes = {_issue_code(issue) for issue in issues}
    issue_messages = [_issue_message(issue) for issue in issues]
    should_repair_test = bool({"thin-test-chapter", "weak-test-method"} & issue_codes)
    should_insert_ack = any("致谢" in message and ("缺少" in message or "未识别" in message) for message in issue_messages)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    actions: list[str] = []
    skipped: list[str] = []
    with zipfile.ZipFile(input_path) as zin:
        root = ET.fromstring(zin.read("word/document.xml"))
        body = root.find("w:body", NS)
        if body is None:
            raise ValueError("word/document.xml does not contain w:body")

        language_fixes = _apply_language_cleanup(body)
        if language_fixes:
            actions.append(f"规范化正文语言、标点或引用格式 {language_fixes} 处")

        full_text = "\n".join(_paragraph_text(p) for p in body.iter(_w("p")))
        test_augmented = False
        test_count = 0
        if should_repair_test:
            test_augmented, test_count = _augment_test_chapter(body, full_text)
            if test_augmented:
                actions.append("根据内容审阅意见补强测试章节")
                full_text = "\n".join(_paragraph_text(p) for p in body.iter(_w("p")))
            else:
                skipped.append("测试章节未找到合适插入点，或已存在测试环境/结果分析内容")

        ack_inserted = False
        ack_count = 0
        if should_insert_ack:
            ack_inserted, ack_count = _insert_acknowledgements(body, full_text)
            if ack_inserted:
                actions.append("根据缺少致谢的审阅意见补齐致谢章节")
            else:
                skipped.append("致谢章节已存在，未重复插入")

        document_xml = serialize_xml(root)
        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = document_xml if item.filename == "word/document.xml" else zin.read(item.filename)
                zout.writestr(item, data)

    return ContentRepairReport(
        input=input_path,
        output=output_path,
        language_fixes_applied=language_fixes,
        test_chapter_augmented=test_augmented,
        acknowledgements_inserted=ack_inserted,
        inserted_paragraphs=test_count + ack_count,
        actions=actions,
        skipped=skipped,
    )


def _issue_code(issue: Any) -> str:
    if isinstance(issue, dict):
        return str(issue.get("code", ""))
    return str(getattr(issue, "code", ""))


def _issue_message(issue: Any) -> str:
    if isinstance(issue, dict):
        return str(issue.get("message", ""))
    return str(getattr(issue, "message", ""))


def _w(local: str) -> str:
    return f"{{{W_NS}}}{local}"
