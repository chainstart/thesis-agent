from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass(frozen=True)
class GateResult:
    passed: bool
    score: int
    blockers: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def evaluate_quality_gate(audit_result: dict[str, Any]) -> GateResult:
    visual = audit_result["target_visual"]
    content = audit_result["content_review"]
    docx = audit_result["docx"]
    blockers: list[str] = []
    warnings: list[str] = []

    if visual.blank_pages:
        blockers.append(f"存在空白页：{_fmt(visual.blank_pages)}")
    if visual.near_blank_pages:
        blockers.append(f"存在疑似空白页：{_fmt(visual.near_blank_pages)}")
    if visual.broken_reference_pages:
        blockers.append(f"存在断引用页：{_fmt(visual.broken_reference_pages)}")
    if visual.toc_title_split_pages:
        blockers.append(f"目录标题拆页：{_fmt(visual.toc_title_split_pages)}")
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
    for warning in getattr(visual, "figure_readability_warnings", []):
        warnings.append(warning)
    if docx.supported and docx.empty_paragraph_runs:
        blockers.append("仍存在长空段排版")
    orphan_empty_errors = getattr(docx, "orphan_empty_paragraph_errors", [])
    if docx.supported and orphan_empty_errors:
        blockers.append(f"正文存在多余空段：{_fmt(orphan_empty_errors[:8])}")
    if docx.supported and docx.broken_references:
        blockers.append("DOCX 文本中存在断引用")
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
    reference_errors = getattr(docx, "reference_format_errors", [])
    if docx.supported and reference_errors:
        blockers.append(f"参考文献格式不符合模板红字要求：{_fmt(reference_errors[:8])}")
    ack_errors = getattr(docx, "acknowledgement_format_errors", [])
    if docx.supported and ack_errors:
        blockers.append(f"致谢格式不符合模板红字要求：{_fmt(ack_errors[:8])}")

    for issue in content.issues:
        if issue.severity == "error":
            blockers.append(issue.message)
        else:
            warnings.append(issue.message)

    score = max(0, 100 - 15 * len(blockers) - 4 * len(warnings))
    return GateResult(
        passed=not blockers,
        score=score,
        blockers=blockers,
        warnings=warnings,
    )


def _fmt(values: list[Any]) -> str:
    return ", ".join(str(v) for v in values)
