from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass(frozen=True)
class GateResult:
    passed: bool
    score: int
    format_score: int
    content_score: int
    hard_blocker_count: int = 0
    hard_blockers: list[str] = field(default_factory=list)
    blockers: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    format_blockers: list[str] = field(default_factory=list)
    format_warnings: list[str] = field(default_factory=list)
    content_blockers: list[str] = field(default_factory=list)
    content_warnings: list[str] = field(default_factory=list)


def evaluate_quality_gate(audit_result: dict[str, Any]) -> GateResult:
    visual = audit_result["target_visual"]
    content = audit_result["content_review"]
    docx = audit_result["docx"]
    requirements = audit_result.get("document_requirements")
    document_profile = audit_result.get("document_profile")
    document_kind = getattr(document_profile, "kind", "thesis") if document_profile is not None else "thesis"
    format_blockers, format_warnings = _format_findings(visual, docx, requirements, document_kind=document_kind)
    content_blockers, content_warnings = _content_findings(content)
    hard_blockers = [*format_blockers, *content_blockers]

    format_score = _score(format_blockers, format_warnings)
    content_score = int(getattr(content, "score", 0))

    if format_score < 80:
        format_blockers.append(f"格式分 {format_score}/100，低于通过线 80 分")
    if content_score < 80:
        content_blockers.append(f"内容分 {content_score}/100，低于通过线 80 分")

    blockers = [*format_blockers, *content_blockers]
    warnings = [*format_warnings, *content_warnings]
    score = min(format_score, content_score)
    return GateResult(
        passed=not hard_blockers and format_score >= 80 and content_score >= 80,
        score=score,
        format_score=format_score,
        content_score=content_score,
        hard_blocker_count=len(hard_blockers),
        hard_blockers=hard_blockers,
        blockers=blockers,
        warnings=warnings,
        format_blockers=format_blockers,
        format_warnings=format_warnings,
        content_blockers=content_blockers,
        content_warnings=content_warnings,
    )


def _format_findings(
    visual: Any,
    docx: Any,
    requirements: Any | None = None,
    document_kind: str = "thesis",
) -> tuple[list[str], list[str]]:
    blockers: list[str] = []
    warnings: list[str] = []
    thesis_like = document_kind == "thesis"
    if visual.blank_pages:
        blockers.append(f"存在空白页：{_fmt(visual.blank_pages)}")
    if visual.near_blank_pages:
        blockers.append(f"存在疑似空白页：{_fmt(visual.near_blank_pages)}")
    if visual.broken_reference_pages:
        blockers.append(f"存在断引用页：{_fmt(visual.broken_reference_pages)}")
    if visual.toc_title_split_pages:
        blockers.append(f"目录标题拆页：{_fmt(visual.toc_title_split_pages)}")
    if thesis_like:
        front_errors = getattr(visual, "front_matter_page_number_errors", {})
        if front_errors:
            formatted = ", ".join(f"{page}:{label}" for page, label in sorted(front_errors.items()))
            blockers.append(f"正文前页码应为罗马数字，检测到阿拉伯数字：{formatted}")
        toc_mismatches = getattr(visual, "toc_page_number_mismatches", [])
        if toc_mismatches:
            blockers.append(f"目录页码与正文页码不一致：{_fmt(toc_mismatches[:8])}")
        front_layout_errors = getattr(visual, "front_matter_layout_errors", [])
        if front_layout_errors:
            blockers.append(f"前置页混页：{_fmt(front_layout_errors[:8])}")
        header_alignment_errors = getattr(visual, "header_page_number_alignment_errors", [])
        if header_alignment_errors:
            blockers.append(f"页眉页码未右顶格：{_fmt(header_alignment_errors[:8])}")
    if visual.caption_orphan_pages:
        blockers.append(f"疑似题注孤页：{_fmt(visual.caption_orphan_pages)}")
    figure_warnings = getattr(visual, "figure_readability_warnings", [])
    if figure_warnings:
        blockers.append(f"图像可读性不足，需要基于原图放大或重绘：{_fmt(figure_warnings[:8])}")
    if docx.supported and docx.empty_paragraph_runs:
        blockers.append("仍存在长空段排版")
    orphan_empty_errors = getattr(docx, "orphan_empty_paragraph_errors", [])
    if docx.supported and orphan_empty_errors:
        blockers.append(f"正文存在多余空段：{_fmt(orphan_empty_errors[:8])}")
    if docx.supported and docx.broken_references:
        blockers.append("DOCX 文本中存在断引用")
    if thesis_like:
        cover_errors = getattr(docx, "cover_format_errors", [])
        if docx.supported and cover_errors:
            blockers.append(f"封面格式不符合模板要求：{_fmt(cover_errors[:8])}")
        abstract_errors = getattr(docx, "abstract_format_errors", [])
        if docx.supported and abstract_errors:
            blockers.append(f"摘要/关键词格式不符合模板红字要求：{_fmt(abstract_errors[:8])}")
        heading_errors = getattr(docx, "main_heading_format_errors", [])
        if docx.supported and heading_errors:
            blockers.append(f"一级标题格式不符合模板红字要求：{_fmt(heading_errors[:8])}")
        sub_heading_errors = getattr(docx, "sub_heading_format_errors", [])
        if docx.supported and sub_heading_errors:
            blockers.append(f"二三级标题格式不符合模板红字要求：{_fmt(sub_heading_errors[:8])}")
        toc_errors = getattr(docx, "toc_format_errors", [])
        if docx.supported and toc_errors:
            blockers.append(f"目录格式不符合模板红字要求：{_fmt(toc_errors[:8])}")
        body_errors = getattr(docx, "body_paragraph_format_errors", [])
        if docx.supported and body_errors:
            blockers.append(f"正文段落格式不符合模板红字要求：{_fmt(body_errors[:8])}")
    table_errors = getattr(docx, "table_format_errors", [])
    if docx.supported and table_errors:
        blockers.append(f"表格格式不符合模板要求：{_fmt(table_errors[:8])}")
    caption_errors = getattr(docx, "caption_format_errors", [])
    if docx.supported and caption_errors:
        blockers.append(f"图表题注格式不符合模板红字要求：{_fmt(caption_errors[:8])}")
    if thesis_like:
        reference_errors = getattr(docx, "reference_format_errors", [])
        if docx.supported and reference_errors:
            blockers.append(f"参考文献格式不符合模板红字要求：{_fmt(reference_errors[:8])}")
        ack_errors = getattr(docx, "acknowledgement_format_errors", [])
        if docx.supported and ack_errors:
            blockers.append(f"致谢格式不符合模板红字要求：{_fmt(ack_errors[:8])}")
    if requirements is not None:
        required_field_errors = getattr(requirements, "required_field_errors", [])
        if required_field_errors:
            blockers.append(f"必填信息不完整：{_fmt(required_field_errors[:8])}")
        signature_errors = getattr(requirements, "signature_errors", [])
        if signature_errors:
            blockers.append(f"签名图片不符合要求：{_fmt(signature_errors[:8])}")
        opinion_errors = getattr(requirements, "opinion_errors", [])
        if opinion_errors:
            blockers.append(f"意见区内容不完整：{_fmt(opinion_errors[:8])}")
        metadata_warnings = getattr(requirements, "metadata_warnings", [])
        if metadata_warnings:
            warnings.extend(metadata_warnings)
    return blockers, warnings


def _content_findings(content: Any) -> tuple[list[str], list[str]]:
    blockers: list[str] = []
    warnings: list[str] = []
    for issue in content.issues:
        if issue.severity == "error":
            blockers.append(issue.message)
        else:
            warnings.append(issue.message)
    return blockers, warnings


def _score(blockers: list[str], warnings: list[str]) -> int:
    return max(0, 100 - 15 * len(blockers) - 4 * len(warnings))


def _fmt(values: list[Any]) -> str:
    return ", ".join(str(v) for v in values)
