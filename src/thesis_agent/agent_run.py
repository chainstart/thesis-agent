from __future__ import annotations

import json
import re
import shutil
import subprocess
import sys
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

from .annotations import strip_red_annotations_from_docx
from .config import AgentConfig
from .content_enhance import enhance_docx_content
from .content_plan import write_content_plan
from .content_repair import repair_reviewed_content
from .diagram_repair import repair_known_diagrams
from .docx_inspect import inspect_docx
from .document_profile import THESIS, build_document_profile
from .document_convert import ensure_docx
from .format_fix import fix_docx_format
from .metadata_store import (
    apply_metadata_to_docx,
    default_metadata_store,
    extract_document_metadata,
    lookup_metadata,
    update_metadata_store,
)
from .pipeline import _jsonable
from .quality_gate import evaluate_quality_gate
from .report_format import fix_report_template_format
from .slot_fill import fill_standard_template_docx, final_output_filename
from .template_profile import build_template_profile
from .template_checklist import write_red_text_checklist
from .template_rebuild import rebuild_standard_template
from .tools import Toolchain
from .toc_sync import sync_static_toc_from_pdf
from .visual_repair import repair_visual_items
from .vision_pack import build_vision_pack


@dataclass(frozen=True)
class ProcessResult:
    target: Path
    final_docx: Path
    audit_report: Path
    vision_prompt: Path
    content_plan: Path
    revision_checklist: Path
    gate_passed: bool
    gate_score: int
    format_score: int
    content_score: int
    hard_blocker_count: int
    hard_blockers: list[str]
    format_blockers: list[str]
    content_blockers: list[str]
    blockers: list[str]
    document_kind: str = "thesis"
    document_label: str = "毕业论文"
    pipeline: str = "thesis-heavy"


def process_document(
    template: Path,
    target: Path,
    out_dir: Path,
    config: AgentConfig,
    toolchain: Toolchain,
    build_vision: bool = True,
    metadata_store_path: Path | None = None,
) -> ProcessResult:
    profile = build_document_profile(target, template)
    routed_template = profile.template or template
    metadata_store_path = metadata_store_path or default_metadata_store(out_dir.parent)
    if profile.kind == THESIS:
        return process_thesis(
            routed_template,
            target,
            out_dir,
            config,
            toolchain,
            build_vision=build_vision,
            metadata_store_path=metadata_store_path,
        )
    return process_report_document(
        routed_template,
        target,
        out_dir,
        config,
        toolchain,
        build_vision=build_vision,
        metadata_store_path=metadata_store_path,
    )


def process_thesis(
    template: Path,
    target: Path,
    out_dir: Path,
    config: AgentConfig,
    toolchain: Toolchain,
    build_vision: bool = True,
    metadata_store_path: Path | None = None,
) -> ProcessResult:
    out_dir.mkdir(parents=True, exist_ok=True)
    work_dir = out_dir / "work"
    prepared = ensure_docx(target, work_dir, toolchain)
    source_text = inspect_docx(prepared.docx).text
    metadata = _known_metadata(target, source_text, metadata_store_path or default_metadata_store(out_dir.parent))
    profile = build_template_profile(template)
    (out_dir / "template_profile.json").write_text(profile.to_json(), encoding="utf-8")
    write_red_text_checklist(
        template,
        out_dir / "template_red_checklist.md",
        out_dir / "template_red_checklist.json",
    )
    processing_template = _prepare_processing_template(template, work_dir, out_dir)

    rebuilt = work_dir / "rebuilt_from_template.docx"
    rebuild_report = fill_standard_template_docx(processing_template, prepared.docx, rebuilt, metadata_overrides=metadata)
    (out_dir / "slot_fill_report.json").write_text(rebuild_report.to_json(), encoding="utf-8")
    (out_dir / "rebuild_report.json").write_text(rebuild_report.to_json(), encoding="utf-8")

    format_pass1 = work_dir / "format_pass1.docx"
    fix_report_pass1 = fix_docx_format(
        rebuilt,
        format_pass1,
        template_path=processing_template,
        preserve_template_front_matter=True,
    )
    (out_dir / "fix_report_pass1.json").write_text(
        json.dumps(_jsonable(asdict(fix_report_pass1)), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    enhanced = work_dir / "content_enhanced.docx"
    content_enhance_report = enhance_docx_content(format_pass1, enhanced)
    (out_dir / "content_enhance_report.json").write_text(
        json.dumps(_jsonable(asdict(content_enhance_report)), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    fixed = out_dir / final_output_filename(rebuild_report)
    fix_report = fix_docx_format(
        enhanced,
        fixed,
        template_path=processing_template,
        preserve_template_front_matter=True,
    )
    (out_dir / "fix_report.json").write_text(
        json.dumps(_jsonable(asdict(fix_report)), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    repaired_diagrams = work_dir / "diagrams_repaired.docx"
    diagram_report = repair_known_diagrams(fixed, repaired_diagrams)
    if diagram_report.repaired_diagrams:
        shutil.copy2(repaired_diagrams, fixed)
    (out_dir / "diagram_repair_report.json").write_text(
        json.dumps(_jsonable(asdict(diagram_report)), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    stripped = work_dir / "red_annotations_stripped.docx"
    annotation_report = strip_red_annotations_from_docx(fixed, stripped)
    shutil.copy2(stripped, fixed)
    (out_dir / "annotation_strip_report.json").write_text(annotation_report.to_json(), encoding="utf-8")

    _sync_toc_until_stable(processing_template, fixed, work_dir, out_dir, config, toolchain, report_prefix="toc_sync_report")

    audit_dir = out_dir / "audit"
    audit_result = _run_clean_audit_process(processing_template, fixed, audit_dir, config)
    audit_result = _run_review_repair_loop(processing_template, fixed, audit_result, work_dir, out_dir, config, toolchain)
    audit_result = _run_clean_audit_process(processing_template, fixed, audit_dir, config)
    content_plan_path = out_dir / "content_improvement.md"
    write_content_plan(audit_result, fixed, content_plan_path)

    vision_prompt = out_dir / "vision_review_prompt.md"
    if build_vision:
        pack = build_vision_pack(audit_dir, out_dir / "vision")
        if pack.prompt:
            shutil.copy2(pack.prompt, vision_prompt)

    gate = evaluate_quality_gate(audit_result)
    revision_checklist_path = out_dir / "revision_checklist.md"
    hard_blockers = _student_hard_blockers(audit_result)
    _write_revision_checklist(audit_result, gate, fixed, revision_checklist_path, hard_blockers)
    legacy_final = out_dir / "final.docx"
    if fixed.name != legacy_final.name:
        shutil.copy2(fixed, legacy_final)
    result = ProcessResult(
        target=target,
        final_docx=fixed,
        audit_report=audit_dir / "report.md",
        vision_prompt=vision_prompt,
        content_plan=content_plan_path,
        revision_checklist=revision_checklist_path,
        gate_passed=gate.passed,
        gate_score=gate.score,
        format_score=gate.format_score,
        content_score=gate.content_score,
        hard_blocker_count=len(hard_blockers),
        hard_blockers=hard_blockers,
        format_blockers=gate.format_blockers,
        content_blockers=gate.content_blockers,
        blockers=gate.blockers,
        document_kind=THESIS,
        document_label="毕业论文",
        pipeline="thesis-heavy",
    )
    (out_dir / "process_result.json").write_text(
        json.dumps(_jsonable(asdict(result)), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    (out_dir / "process_report.md").write_text(_process_markdown(result, gate.warnings), encoding="utf-8")
    return result


def process_report_document(
    template: Path,
    target: Path,
    out_dir: Path,
    config: AgentConfig,
    toolchain: Toolchain,
    build_vision: bool = True,
    metadata_store_path: Path | None = None,
) -> ProcessResult:
    out_dir.mkdir(parents=True, exist_ok=True)
    work_dir = out_dir / "work"
    prepared = ensure_docx(target, work_dir, toolchain)
    source_text = inspect_docx(prepared.docx).text
    profile = build_document_profile(target, template, source_text)
    processing_template = profile.template or template
    metadata = _known_metadata(target, source_text, metadata_store_path or default_metadata_store(out_dir.parent))
    (out_dir / "document_profile.json").write_text(json.dumps(_jsonable(asdict(profile)), ensure_ascii=False, indent=2), encoding="utf-8")
    (out_dir / "metadata.json").write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")
    template_docx = None
    if processing_template.exists():
        (out_dir / "template_profile.json").write_text(build_template_profile(processing_template).to_json(), encoding="utf-8")
        write_red_text_checklist(
            processing_template,
            out_dir / "template_red_checklist.md",
            out_dir / "template_red_checklist.json",
        )
        template_docx = ensure_docx(processing_template, work_dir / "template", toolchain).docx

    metadata_filled = work_dir / "metadata_filled.docx"
    filled_fields = apply_metadata_to_docx(prepared.docx, metadata_filled, metadata)
    (out_dir / "metadata_fill_report.json").write_text(
        json.dumps({"filled_fields": filled_fields, "metadata": metadata}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    fixed = out_dir / _report_output_filename(target, profile.label, metadata)
    if template_docx is not None:
        fix_report = fix_report_template_format(metadata_filled, fixed, template_docx)
        report_payload = json.dumps(_jsonable(asdict(fix_report)), ensure_ascii=False, indent=2)
        (out_dir / "report_template_format_report.json").write_text(report_payload, encoding="utf-8")
        (out_dir / "fix_report.json").write_text(report_payload, encoding="utf-8")
    else:
        fix_report = fix_docx_format(metadata_filled, fixed, template_path=None, preserve_template_front_matter=True)
        (out_dir / "fix_report.json").write_text(
            json.dumps(_jsonable(asdict(fix_report)), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    stripped = work_dir / "red_annotations_stripped.docx"
    annotation_report = strip_red_annotations_from_docx(fixed, stripped)
    shutil.copy2(stripped, fixed)
    (out_dir / "annotation_strip_report.json").write_text(annotation_report.to_json(), encoding="utf-8")

    audit_dir = out_dir / "audit"
    audit_result = _run_clean_audit_process(processing_template, fixed, audit_dir, config)
    content_plan_path = out_dir / "content_improvement.md"
    write_content_plan(audit_result, fixed, content_plan_path)
    vision_prompt = out_dir / "vision_review_prompt.md"
    if build_vision:
        pack = build_vision_pack(audit_dir, out_dir / "vision")
        if pack.prompt:
            shutil.copy2(pack.prompt, vision_prompt)

    gate = evaluate_quality_gate(audit_result)
    revision_checklist_path = out_dir / "revision_checklist.md"
    hard_blockers = _student_hard_blockers(audit_result)
    _write_revision_checklist(audit_result, gate, fixed, revision_checklist_path, hard_blockers)
    legacy_final = out_dir / "final.docx"
    if fixed.name != legacy_final.name:
        shutil.copy2(fixed, legacy_final)
    result = ProcessResult(
        target=target,
        final_docx=fixed,
        audit_report=audit_dir / "report.md",
        vision_prompt=vision_prompt,
        content_plan=content_plan_path,
        revision_checklist=revision_checklist_path,
        gate_passed=gate.passed,
        gate_score=gate.score,
        format_score=gate.format_score,
        content_score=gate.content_score,
        hard_blocker_count=len(hard_blockers),
        hard_blockers=hard_blockers,
        format_blockers=gate.format_blockers,
        content_blockers=gate.content_blockers,
        blockers=gate.blockers,
        document_kind=profile.kind,
        document_label=profile.label,
        pipeline=profile.pipeline,
    )
    (out_dir / "process_result.json").write_text(
        json.dumps(_jsonable(asdict(result)), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    (out_dir / "process_report.md").write_text(_process_markdown(result, gate.warnings), encoding="utf-8")
    return result


def _run_clean_audit_process(template: Path, target: Path, out_dir: Path, config: AgentConfig) -> dict[str, Any]:
    """Run audit in a fresh Python process so generation state cannot leak into review."""
    cmd = [
        sys.executable,
        "-m",
        "thesis_agent",
        "--config",
        str(config.path),
        "audit",
        "--template",
        str(template),
        "--target",
        str(target),
        "--out",
        str(out_dir),
    ]
    subprocess.run(cmd, check=True, capture_output=True, text=True)
    report_path = out_dir / "report.json"
    with report_path.open("r", encoding="utf-8") as fh:
        payload = json.load(fh)
    return _audit_payload_with_attributes(payload)


def _audit_payload_with_attributes(payload: dict[str, Any]) -> dict[str, Any]:
    converted = dict(payload)
    for key in ("template_visual", "target_visual", "document_profile", "docx", "content_review", "document_requirements"):
        if key in converted:
            converted[key] = _to_namespace(converted[key])
    return converted


def _run_review_repair_loop(
    template: Path,
    fixed: Path,
    audit_result: dict[str, Any],
    work_dir: Path,
    out_dir: Path,
    config: AgentConfig,
    toolchain: Toolchain,
    max_passes: int = 2,
) -> dict[str, Any]:
    loop_reports: list[dict[str, Any]] = []
    current_audit = audit_result
    for pass_idx in range(1, max_passes + 1):
        gate = evaluate_quality_gate(current_audit)
        if gate.passed:
            break

        changed = False
        pass_reports: dict[str, Any] = {"pass": pass_idx}

        visual_output = work_dir / f"review_loop_visual_pass{pass_idx}.docx"
        visual_report = repair_visual_items(fixed, visual_output, current_audit)
        pass_reports["visual_repair"] = _jsonable(asdict(visual_report))
        (out_dir / f"visual_repair_report_pass{pass_idx}.json").write_text(
            json.dumps(_jsonable(asdict(visual_report)), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        if visual_report.enhanced_figures or visual_report.extracted_code_formula_figures or visual_report.renumbered_figures or visual_report.updated_references:
            shutil.copy2(visual_output, fixed)
            changed = True

        content_output = work_dir / f"review_loop_content_pass{pass_idx}.docx"
        content_report = repair_reviewed_content(fixed, content_output, getattr(current_audit["content_review"], "issues", []))
        pass_reports["content_repair"] = _jsonable(asdict(content_report))
        (out_dir / f"content_repair_report_pass{pass_idx}.json").write_text(
            json.dumps(_jsonable(asdict(content_report)), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        if content_report.actions:
            shutil.copy2(content_output, fixed)
            changed = True

        loop_reports.append(pass_reports)
        if not changed:
            break

        formatted = work_dir / f"review_loop_format_pass{pass_idx}.docx"
        format_report = fix_docx_format(
            fixed,
            formatted,
            template_path=template,
            preserve_template_front_matter=True,
        )
        (out_dir / f"repair_loop_format_report_pass{pass_idx}.json").write_text(
            json.dumps(_jsonable(asdict(format_report)), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        shutil.copy2(formatted, fixed)

        stripped = work_dir / f"review_loop_stripped_pass{pass_idx}.docx"
        annotation_report = strip_red_annotations_from_docx(fixed, stripped)
        (out_dir / f"repair_loop_annotation_strip_report_pass{pass_idx}.json").write_text(annotation_report.to_json(), encoding="utf-8")
        shutil.copy2(stripped, fixed)

        _sync_toc_until_stable(
            template,
            fixed,
            work_dir,
            out_dir,
            config,
            toolchain,
            report_prefix=f"repair_loop_toc_sync_report_pass{pass_idx}",
        )
        current_audit = _run_clean_audit_process(template, fixed, work_dir / f"review_loop_audit_pass{pass_idx}", config)

    (out_dir / "repair_loop_report.json").write_text(json.dumps(loop_reports, ensure_ascii=False, indent=2), encoding="utf-8")
    return current_audit


def _sync_toc_until_stable(
    template: Path,
    fixed: Path,
    work_dir: Path,
    out_dir: Path,
    config: AgentConfig,
    toolchain: Toolchain,
    report_prefix: str,
) -> list[Any]:
    toc_reports = []
    for pass_idx in range(1, 4):
        pass_audit_dir = work_dir / f"{report_prefix}_audit{pass_idx}"
        pass_audit = _run_clean_audit_process(template, fixed, pass_audit_dir, config)
        synced = work_dir / f"{report_prefix}_synced{pass_idx}.docx"
        toc_report = sync_static_toc_from_pdf(
            fixed,
            Path(pass_audit["files"]["target_pdf"]),
            synced,
            toolchain,
        )
        toc_reports.append(toc_report)
        report_name = f"{report_prefix}_pass{pass_idx}.json" if report_prefix == "toc_sync_report" else f"{report_prefix}_{pass_idx}.json"
        (out_dir / report_name).write_text(
            json.dumps(_jsonable(asdict(toc_report)), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        if not toc_report.updated_entries and not toc_report.inserted_entries:
            break
        shutil.copy2(synced, fixed)
    (out_dir / f"{report_prefix}.json").write_text(
        json.dumps([_jsonable(asdict(report)) for report in toc_reports], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return toc_reports


def _to_namespace(value: Any) -> Any:
    if isinstance(value, dict):
        return AttrDict({key: _to_namespace(item) for key, item in value.items()})
    if isinstance(value, list):
        return [_to_namespace(item) for item in value]
    return value


class AttrDict(dict):
    def __getattr__(self, name: str) -> Any:
        try:
            return self[name]
        except KeyError as exc:
            raise AttributeError(name) from exc


def _known_metadata(target: Path, text: str, store_path: Path) -> dict[str, str]:
    extracted = extract_document_metadata(target, text)
    known = lookup_metadata(store_path, extracted)
    if extracted or known:
        known = update_metadata_store(store_path, known or extracted, target)
    return known


def _report_output_filename(target: Path, label: str, metadata: dict[str, str]) -> str:
    student_id = _safe_filename_part(metadata.get("student_id"))
    student_name = _safe_filename_part(metadata.get("student_name"))
    label_part = _safe_filename_part(label)
    if not metadata.get("title"):
        return f"{student_id}-{student_name}-{label_part}.docx"
    title = _safe_filename_part(metadata.get("title"))
    if title == label_part:
        return f"{student_id}-{student_name}-{label_part}.docx"
    return f"{student_id}-{student_name}-{label_part}-{title}.docx"


def _safe_filename_part(value: str | None) -> str:
    if not value:
        return "XX"
    safe = re.sub(r"[\\/:*?\"<>|]+", "", str(value)).strip()
    safe = re.sub(r"\s+", "", safe)
    return safe[:40] or "XX"


def _prepare_processing_template(template: Path, work_dir: Path, out_dir: Path) -> Path:
    cover = _find_cover_template(template)
    if cover is None:
        return template
    standard_template = work_dir / "standard_template.docx"
    template_report = rebuild_standard_template(cover, template, standard_template)
    (out_dir / "standard_template_report.json").write_text(template_report.to_json(), encoding="utf-8")
    formal_template = work_dir / "standard_template_formal.docx"
    strip_report = strip_red_annotations_from_docx(standard_template, formal_template)
    (out_dir / "standard_template_annotation_strip_report.json").write_text(strip_report.to_json(), encoding="utf-8")
    return formal_template


def _find_cover_template(template: Path) -> Path | None:
    template_docx = template if template.suffix.lower() == ".docx" else template.with_suffix(".docx")
    if not template.parent.exists():
        return None
    candidates = [
        path
        for path in sorted(template.parent.glob("*.docx"))
        if path != template_docx and "Zone.Identifier" not in path.name and not path.name.startswith("~$")
    ]
    preferred = [path for path in candidates if any(token in path.stem.lower() for token in ("cover", "封面"))]
    return (preferred or candidates)[0] if candidates else None


def _process_markdown(result: ProcessResult, warnings: list[str]) -> str:
    lines = [
        "# Thesis Agent Process Report",
        "",
        f"- Target: `{result.target}`",
        f"- Document type: {result.document_label} (`{result.pipeline}`)",
        f"- Final DOCX: `{result.final_docx}`",
        f"- Gate: {'PASS' if result.gate_passed else 'FAIL'}",
        "- Pass rule: format score >= 80, content score >= 80, and hard blocker count = 0.",
        f"- Gate score: {result.gate_score}/100",
        f"- Format score: {result.format_score}/100",
        f"- Content score: {result.content_score}/100",
        f"- Hard blockers: {result.hard_blocker_count}",
        f"- Audit report: `{result.audit_report}`",
        f"- Vision prompt: `{result.vision_prompt}`",
        f"- Content plan: `{result.content_plan}`",
        f"- Revision checklist: `{result.revision_checklist}`",
        "",
        "## Blockers",
        "",
    ]
    if result.blockers:
        lines.extend(f"- {item}" for item in result.blockers)
    else:
        lines.append("- none")
    lines.extend(["", "## Student Hard Blockers", ""])
    if result.hard_blockers:
        lines.extend(f"- {item}" for item in result.hard_blockers)
    else:
        lines.append("- none")
    lines.extend(["", "## Warnings", ""])
    if warnings:
        lines.extend(f"- {item}" for item in warnings)
    else:
        lines.append("- none")
    return "\n".join(lines) + "\n"


def _write_revision_checklist(
    audit_result: dict[str, Any],
    gate,
    target: Path,
    out_path: Path,
    hard_blockers: list[str],
) -> None:
    visual = audit_result["target_visual"]
    content = audit_result["content_review"]
    docx = audit_result["docx"]
    requirements = audit_result.get("document_requirements")
    document_profile = audit_result.get("document_profile")
    document_kind = getattr(document_profile, "kind", "thesis") if document_profile is not None else "thesis"
    format_deductions = _student_format_deductions(visual, docx, requirements, document_kind=document_kind)
    content_deductions = _student_content_deductions(content)
    lines = [
        "# 论文返修清单",
        "",
        f"- 文件: `{target}`",
        f"- 结论: `{'PASS' if gate.passed else 'FAIL'}`",
        f"- 通过规则: 格式分 >= 80、内容分 >= 80、硬阻塞数 = 0，三项同时满足才通过。",
        f"- 格式分: {gate.format_score}/100（{'通过' if gate.format_score >= 80 else '未通过'}）",
        f"- 内容分: {gate.content_score}/100（{'通过' if gate.content_score >= 80 else '未通过'}）",
        f"- 硬阻塞数: {len(hard_blockers)}（{'通过' if not hard_blockers else '未通过'}）",
        "",
        "## 硬阻塞",
        "",
    ]
    if hard_blockers:
        lines.extend(f"- {item}" for item in hard_blockers)
    else:
        lines.append("- 无。")

    lines.extend(["", "## 格式扣分点", ""])
    lines.append(_format_score_formula(gate))
    format_only_deductions = [item for item in format_deductions if item not in set(hard_blockers)]
    if format_only_deductions:
        lines.extend(f"- {item}" for item in format_only_deductions)
    elif format_deductions:
        lines.append("- 本轮格式扣分点均已列入“硬阻塞”，按上方硬阻塞条目修改。")
    else:
        lines.append("- 未发现明确格式扣分点。")

    lines.extend(["", "## 内容扣分点", ""])
    lines.append("- 内容计分说明: 内容问题按规则扣分，`error` 每项扣 12 分，`warning` 每项扣 5 分。")
    if content_deductions:
        lines.extend(f"- {item}" for item in content_deductions)
    else:
        lines.append("- 未发现明确内容扣分点。")
    out_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def _student_hard_blockers(audit_result: dict[str, Any]) -> list[str]:
    visual = audit_result["target_visual"]
    content = audit_result["content_review"]
    docx = audit_result["docx"]
    requirements = audit_result.get("document_requirements")
    document_profile = audit_result.get("document_profile")
    document_kind = getattr(document_profile, "kind", "thesis") if document_profile is not None else "thesis"
    items = _student_format_deductions(visual, docx, requirements, document_kind=document_kind)
    for issue in getattr(content, "issues", []):
        if getattr(issue, "severity", "") == "error":
            items.append(f"内容硬阻塞：{getattr(issue, 'message', str(issue))}")
    return _dedupe(items)


def _student_format_deductions(
    visual: Any,
    docx: Any,
    requirements: Any | None = None,
    document_kind: str = "thesis",
) -> list[str]:
    items: list[str] = []
    thesis_like = document_kind == "thesis"
    _extend_values(items, "存在空白页", getattr(visual, "blank_pages", []))
    _extend_values(items, "存在疑似空白页", getattr(visual, "near_blank_pages", []))
    _extend_values(items, "存在断引用页", getattr(visual, "broken_reference_pages", []))
    _extend_values(items, "目录标题拆页", getattr(visual, "toc_title_split_pages", []))

    if thesis_like:
        front_errors = getattr(visual, "front_matter_page_number_errors", {})
        if front_errors:
            for page, label in sorted(front_errors.items()):
                items.append(f"正文前页码错误：第 {page} 页显示为 `{label}`，应为罗马数字。")
        _extend_values(items, "目录页码与正文页码不一致", getattr(visual, "toc_page_number_mismatches", []))
        _extend_values(items, "前置页混页", getattr(visual, "front_matter_layout_errors", []))
        _extend_values(items, "页眉页码未右顶格", getattr(visual, "header_page_number_alignment_errors", []))
    _extend_values(items, "疑似题注孤页", getattr(visual, "caption_orphan_pages", []))
    _extend_values(
        items,
        "图像可读性不足，需要学生提供清晰原图或重新截图",
        getattr(visual, "figure_readability_warnings", []),
    )

    if getattr(docx, "supported", False):
        if getattr(docx, "empty_paragraph_runs", []):
            items.append("仍存在长空段排版：应使用分页符或段落间距，不应靠连续空段排版。")
        _extend_values(items, "正文存在多余空段", getattr(docx, "orphan_empty_paragraph_errors", []))
        if getattr(docx, "broken_references", []):
            items.append("DOCX 文本中存在断引用：需要修复交叉引用或删除失效域。")
        attrs = [
            ("cover_format_errors", "封面格式不符合模板要求"),
            ("abstract_format_errors", "摘要/关键词格式不符合模板要求"),
            ("main_heading_format_errors", "一级标题格式不符合模板要求"),
            ("sub_heading_format_errors", "二三级标题格式不符合模板要求"),
            ("toc_format_errors", "目录格式不符合模板要求"),
            ("body_paragraph_format_errors", "正文段落格式不符合模板要求"),
            ("table_format_errors", "表格格式不符合模板要求"),
            ("caption_format_errors", "图表题注格式不符合模板要求"),
            ("reference_format_errors", "参考文献格式不符合模板要求"),
            ("acknowledgement_format_errors", "致谢格式不符合模板要求"),
        ]
        if not thesis_like:
            attrs = [
                ("table_format_errors", "表格格式不符合模板要求"),
                ("caption_format_errors", "图表题注格式不符合模板要求"),
            ]
        for attr, label in attrs:
            for value in getattr(docx, attr, []):
                if attr == "caption_format_errors" and "图名上方应存在对应图片" in str(value):
                    items.append(f"缺少图片，需要学生补充对应原图：{value}")
                else:
                    items.append(f"{label}：{value}")
    if requirements is not None:
        _extend_values(items, "必填信息不完整", getattr(requirements, "required_field_errors", []))
        _extend_values(items, "签名图片不符合要求", getattr(requirements, "signature_errors", []))
        _extend_values(items, "意见区内容不完整", getattr(requirements, "opinion_errors", []))
    return _dedupe(items)


def _student_content_deductions(content: Any) -> list[str]:
    items: list[str] = []
    for issue in getattr(content, "issues", []):
        severity = getattr(issue, "severity", "")
        message = getattr(issue, "message", str(issue))
        if severity == "error":
            items.append(f"`error` 扣 12 分：{message}")
        elif severity == "warning":
            items.append(f"`warning` 扣 5 分：{message}")
        else:
            items.append(message)
    return _dedupe(items)


def _format_score_formula(gate: Any) -> str:
    blocker_count = len([item for item in getattr(gate, "format_blockers", []) if not str(item).startswith("格式分 ")])
    warning_count = len(getattr(gate, "format_warnings", []))
    return (
        f"- 格式计分说明: 100 - 15×格式问题类别({blocker_count}) "
        f"- 4×格式警告类别({warning_count}) = {gate.format_score}/100；具体修改点见本清单条目。"
    )


def _extend_values(items: list[str], label: str, values: Any) -> None:
    if not values:
        return
    for value in values:
        items.append(f"{label}：{value}")


def _dedupe(items: list[str]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for item in items:
        if item in seen:
            continue
        seen.add(item)
        result.append(item)
    return result
