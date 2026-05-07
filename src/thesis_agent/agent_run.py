from __future__ import annotations

import json
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
from .diagram_repair import repair_known_diagrams
from .document_convert import ensure_docx
from .format_fix import fix_docx_format
from .pipeline import _jsonable
from .quality_gate import evaluate_quality_gate
from .slot_fill import fill_standard_template_docx, final_output_filename
from .template_profile import build_template_profile
from .template_checklist import write_red_text_checklist
from .template_rebuild import rebuild_standard_template
from .tools import Toolchain
from .toc_sync import sync_static_toc_from_pdf
from .vision_pack import build_vision_pack


@dataclass(frozen=True)
class ProcessResult:
    target: Path
    final_docx: Path
    audit_report: Path
    vision_prompt: Path
    content_plan: Path
    gate_passed: bool
    gate_score: int
    blockers: list[str]


def process_thesis(
    template: Path,
    target: Path,
    out_dir: Path,
    config: AgentConfig,
    toolchain: Toolchain,
    build_vision: bool = True,
) -> ProcessResult:
    out_dir.mkdir(parents=True, exist_ok=True)
    work_dir = out_dir / "work"
    prepared = ensure_docx(target, work_dir, toolchain)
    profile = build_template_profile(template)
    (out_dir / "template_profile.json").write_text(profile.to_json(), encoding="utf-8")
    write_red_text_checklist(
        template,
        out_dir / "template_red_checklist.md",
        out_dir / "template_red_checklist.json",
    )
    processing_template = _prepare_processing_template(template, work_dir, out_dir)

    rebuilt = work_dir / "rebuilt_from_template.docx"
    rebuild_report = fill_standard_template_docx(processing_template, prepared.docx, rebuilt)
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

    toc_reports = []
    for pass_idx in range(1, 4):
        pass_audit_dir = work_dir / f"audit_pass{pass_idx}"
        pass_audit = _run_clean_audit_process(processing_template, fixed, pass_audit_dir, config)
        synced = work_dir / f"toc_synced_pass{pass_idx}.docx"
        toc_report = sync_static_toc_from_pdf(
            fixed,
            Path(pass_audit["files"]["target_pdf"]),
            synced,
            toolchain,
        )
        toc_reports.append(toc_report)
        (out_dir / f"toc_sync_report_pass{pass_idx}.json").write_text(
            json.dumps(_jsonable(asdict(toc_report)), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        if not toc_report.updated_entries and not toc_report.inserted_entries:
            break
        shutil.copy2(synced, fixed)
    (out_dir / "toc_sync_report.json").write_text(
        json.dumps([_jsonable(asdict(report)) for report in toc_reports], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

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
    legacy_final = out_dir / "final.docx"
    if fixed.name != legacy_final.name:
        shutil.copy2(fixed, legacy_final)
    result = ProcessResult(
        target=target,
        final_docx=fixed,
        audit_report=audit_dir / "report.md",
        vision_prompt=vision_prompt,
        content_plan=content_plan_path,
        gate_passed=gate.passed,
        gate_score=gate.score,
        blockers=gate.blockers,
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
    for key in ("template_visual", "target_visual", "docx", "content_review"):
        if key in converted:
            converted[key] = _to_namespace(converted[key])
    return converted


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


def _prepare_processing_template(template: Path, work_dir: Path, out_dir: Path) -> Path:
    cover = template.parent / "附件15 学士学位论文封面.docx"
    if not cover.exists():
        return template
    standard_template = work_dir / "standard_template.docx"
    template_report = rebuild_standard_template(cover, template, standard_template)
    (out_dir / "standard_template_report.json").write_text(template_report.to_json(), encoding="utf-8")
    formal_template = work_dir / "standard_template_formal.docx"
    strip_report = strip_red_annotations_from_docx(standard_template, formal_template)
    (out_dir / "standard_template_annotation_strip_report.json").write_text(strip_report.to_json(), encoding="utf-8")
    return formal_template


def _process_markdown(result: ProcessResult, warnings: list[str]) -> str:
    lines = [
        "# Thesis Agent Process Report",
        "",
        f"- Target: `{result.target}`",
        f"- Final DOCX: `{result.final_docx}`",
        f"- Gate: {'PASS' if result.gate_passed else 'FAIL'}",
        f"- Gate score: {result.gate_score}/100",
        f"- Audit report: `{result.audit_report}`",
        f"- Vision prompt: `{result.vision_prompt}`",
        f"- Content plan: `{result.content_plan}`",
        "",
        "## Blockers",
        "",
    ]
    if result.blockers:
        lines.extend(f"- {item}" for item in result.blockers)
    else:
        lines.append("- none")
    lines.extend(["", "## Warnings", ""])
    if warnings:
        lines.extend(f"- {item}" for item in warnings)
    else:
        lines.append("- none")
    return "\n".join(lines) + "\n"
