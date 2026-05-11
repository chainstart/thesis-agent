from __future__ import annotations

import copy
import re
from dataclasses import dataclass
from pathlib import Path

from .config import AgentConfig


THESIS = "thesis"
PROPOSAL = "proposal"
TASK_BOOK = "task_book"
MIDTERM = "midterm"
SCORE_FORM = "score_form"
DEFENSE_RECORD = "defense_record"
GUIDANCE_RECORD = "guidance_record"
UNKNOWN = "unknown"

HEAVY_KINDS = {THESIS}
MEDIUM_KINDS = {PROPOSAL}
LIGHT_KINDS = {TASK_BOOK, MIDTERM, SCORE_FORM, DEFENSE_RECORD, GUIDANCE_RECORD, UNKNOWN}


@dataclass(frozen=True)
class DocumentProfile:
    kind: str
    label: str
    template: Path | None
    pipeline: str


def detect_document_kind(path: Path, text: str = "") -> str:
    path_haystack = _compact(" ".join([str(path), path.stem]))
    if "任务书" in path_haystack:
        return TASK_BOOK
    if "中期检查" in path_haystack or "中期" in path_haystack:
        return MIDTERM
    if "答辩记录" in path_haystack:
        return DEFENSE_RECORD
    if "成绩考核" in path_haystack or "考核表" in path_haystack:
        return SCORE_FORM
    if "指导记录" in path_haystack or "记录本" in path_haystack:
        return GUIDANCE_RECORD
    if "开题" in path_haystack:
        return PROPOSAL

    haystack = _compact(" ".join([path_haystack, text[:3000]]))
    if "毕业设计（论文）任务书" in haystack or "任务书反面" in haystack:
        return TASK_BOOK
    if "毕业设计（论文）开题报告" in haystack:
        return PROPOSAL
    if "任务书" in haystack:
        return TASK_BOOK
    if "开题" in haystack:
        return PROPOSAL
    if "中期检查" in haystack or "中期" in haystack:
        return MIDTERM
    if "答辩记录" in haystack:
        return DEFENSE_RECORD
    if "成绩考核" in haystack or "考核表" in haystack:
        return SCORE_FORM
    if "指导记录" in haystack or "记录本" in haystack:
        return GUIDANCE_RECORD
    if "论文" in haystack or re.search(r"第[一二三四五六]章|参考文献|致谢", haystack):
        return THESIS
    return UNKNOWN


def document_label(kind: str) -> str:
    return {
        THESIS: "毕业论文",
        PROPOSAL: "开题报告",
        TASK_BOOK: "任务书",
        MIDTERM: "中期检查报告",
        SCORE_FORM: "成绩考核表",
        DEFENSE_RECORD: "答辩记录表",
        GUIDANCE_RECORD: "学生工作及教师指导记录本",
        UNKNOWN: "未知材料",
    }.get(kind, "未知材料")


def pipeline_for_kind(kind: str) -> str:
    if kind in HEAVY_KINDS:
        return "thesis-heavy"
    if kind in MEDIUM_KINDS:
        return "report-medium"
    return "report-light"


def find_template_for_kind(kind: str, requested_template: Path | None, target: Path | None = None) -> Path | None:
    if requested_template is not None and kind == THESIS:
        return requested_template
    search_roots = _template_search_roots(requested_template, target)
    tokens = {
        THESIS: ("附件17", "论文格式", "格式规范"),
        PROPOSAL: ("附件13", "开题报告"),
        MIDTERM: ("附件14", "中期检查"),
        TASK_BOOK: ("附件16", "任务书"),
        SCORE_FORM: ("附件21", "成绩考核"),
        DEFENSE_RECORD: ("附件22", "答辩记录"),
        GUIDANCE_RECORD: ("附件23", "记录本", "指导记录"),
    }.get(kind, ())
    for root in search_roots:
        if not root.exists():
            continue
        candidates = [
            path
            for path in sorted(root.rglob("*"))
            if path.is_file()
            and path.suffix.lower() in {".doc", ".docx"}
            and not path.name.startswith("~$")
            and "Zone.Identifier" not in path.name
        ]
        for token in tokens:
            for candidate in candidates:
                if token in candidate.name:
                    return candidate
    return requested_template


def build_document_profile(target: Path, requested_template: Path | None, text: str = "") -> DocumentProfile:
    kind = detect_document_kind(target, text)
    template = find_template_for_kind(kind, requested_template, target)
    return DocumentProfile(
        kind=kind,
        label=document_label(kind),
        template=template,
        pipeline=pipeline_for_kind(kind),
    )


def config_for_document_kind(config: AgentConfig, kind: str) -> AgentConfig:
    if kind == THESIS:
        return config
    data = copy.deepcopy(config.data)
    quality = data.setdefault("content_quality", {})
    if kind == PROPOSAL:
        data["expected_sections"] = ["开题报告", "参考文献"]
        quality.update(
            {
                "min_body_chinese_chars": 2500,
                "recommended_body_chinese_chars": 3500,
                "min_reference_count": 5,
                "min_foreign_reference_count": 0,
                "required_engineering_topics": ["研究意义", "研究内容", "技术路线", "进度安排"],
            }
        )
    elif kind == TASK_BOOK:
        data["expected_sections"] = ["任务书"]
        quality.update(_light_quality(400, ["主要任务", "进度安排"]))
    elif kind == MIDTERM:
        data["expected_sections"] = ["中期检查"]
        quality.update(_light_quality(500, ["已完成", "存在问题", "指导教师意见"]))
    elif kind == DEFENSE_RECORD:
        data["expected_sections"] = ["答辩记录"]
        quality.update(_light_quality(200, ["答辩", "记录"]))
    elif kind == SCORE_FORM:
        data["expected_sections"] = ["成绩考核"]
        quality.update(_light_quality(200, ["成绩", "评定"]))
    elif kind == GUIDANCE_RECORD:
        data["expected_sections"] = ["记录"]
        quality.update(_light_quality(300, ["指导", "记录"]))
    else:
        data["expected_sections"] = []
        quality.update(_light_quality(200, []))
    return AgentConfig(path=config.path, data=data)


def is_supported_input(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() in {".doc", ".docx"} and not path.name.startswith("~$")


def iter_supported_inputs(root: Path) -> list[Path]:
    if root.is_file():
        return [root] if is_supported_input(root) else []
    return [path for path in sorted(root.rglob("*")) if is_supported_input(path)]


def _template_search_roots(requested_template: Path | None, target: Path | None) -> list[Path]:
    roots: list[Path] = []
    if requested_template is not None:
        roots.append(requested_template.parent)
        roots.append(requested_template.parent / "3 毕业设计（论文）材料模板")
    if target is not None:
        current = target.resolve()
        for parent in [current.parent, *current.parents]:
            candidate = parent / "samples" / "templates"
            if candidate.exists():
                roots.append(candidate)
            if parent.name == "samples":
                roots.append(parent / "templates")
    roots.append(Path("samples/templates"))
    roots.append(Path("samples/templates/3 毕业设计（论文）材料模板"))
    deduped: list[Path] = []
    seen: set[Path] = set()
    for root in roots:
        resolved = root.resolve() if root.exists() else root
        if resolved in seen:
            continue
        seen.add(resolved)
        deduped.append(root)
    return deduped


def _light_quality(min_chars: int, topics: list[str]) -> dict[str, object]:
    return {
        "min_body_chinese_chars": min_chars,
        "recommended_body_chinese_chars": min_chars * 2,
        "min_reference_count": 0,
        "min_foreign_reference_count": 0,
        "min_keywords": 0,
        "max_keywords": 20,
        "required_engineering_topics": topics,
    }


def _compact(value: str) -> str:
    return re.sub(r"\s+", "", value)
