from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

from .agent_run import process_document
from .annotations import strip_red_annotations_from_docx
from .config import AgentConfig
from .format_fix import fix_docx_format
from .document_profile import iter_supported_inputs
from .metadata_store import default_metadata_store
from .pipeline import run_audit
from .rebuild import rebuild_thesis_docx
from .slot_fill import fill_standard_template_docx
from .template_profile import build_template_profile
from .template_checklist import write_red_text_checklist
from .template_rebuild import rebuild_standard_template
from .tools import Toolchain, command_version, probe_font
from .visual_compare import compare_standard_template_visual
from .vision_pack import build_vision_pack
from .web_app import serve as serve_web_app


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="thesis-agent")
    parser.add_argument("--config", type=Path, default=None, help="Path to JSON config.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    subparsers.add_parser("doctor", help="Check renderer and Office tool availability.")

    audit = subparsers.add_parser("audit", help="Render and audit a thesis draft against a template.")
    audit.add_argument("--template", type=Path, required=True)
    audit.add_argument("--target", type=Path, required=True)
    audit.add_argument("--out", type=Path, required=True)

    profile = subparsers.add_parser("profile-template", help="Extract a reusable template profile.")
    profile.add_argument("--template", type=Path, required=True)
    profile.add_argument("--out", type=Path, default=None)

    checklist = subparsers.add_parser("template-checklist", help="Extract red-text requirements from a thesis template.")
    checklist.add_argument("--template", type=Path, required=True)
    checklist.add_argument("--out", type=Path, required=True)

    fix_format = subparsers.add_parser("fix-format", help="Apply conservative DOCX format fixes and optionally audit the result.")
    fix_format.add_argument("--target", type=Path, required=True)
    fix_format.add_argument("--output", type=Path, required=True)
    fix_format.add_argument("--template", type=Path, default=None, help="Template for post-fix audit.")
    fix_format.add_argument("--audit-out", type=Path, default=None, help="Run audit after fixing and write report here.")

    rebuild = subparsers.add_parser("rebuild", help="Extract thesis content and rebuild a clean DOCX from the format template.")
    rebuild.add_argument("--template", type=Path, required=True)
    rebuild.add_argument("--target", type=Path, required=True)
    rebuild.add_argument("--output", type=Path, required=True)

    fill_template = subparsers.add_parser("fill-template", help="Fill a formal standard DOCX template with extracted thesis content.")
    fill_template.add_argument("--template", type=Path, required=True)
    fill_template.add_argument("--target", type=Path, required=True)
    fill_template.add_argument("--output", type=Path, required=True)

    template_rebuild = subparsers.add_parser("template-rebuild", help="Build the canonical DOCX template from official cover and body templates.")
    template_rebuild.add_argument("--cover", type=Path, required=True)
    template_rebuild.add_argument("--body", type=Path, required=True)
    template_rebuild.add_argument("--output", type=Path, required=True)
    template_rebuild.add_argument("--strip-red", action="store_true", help="Remove red instructional annotations from the rebuilt template.")

    template_selftest = subparsers.add_parser("template-selftest", help="Rebuild the canonical template and compare it visually with the source templates.")
    template_selftest.add_argument("--cover", type=Path, required=True)
    template_selftest.add_argument("--body", type=Path, required=True)
    template_selftest.add_argument("--out", type=Path, required=True)

    process = subparsers.add_parser("process", help="Run the full generic thesis-agent workflow.")
    process.add_argument("--template", type=Path, required=True)
    process.add_argument("--target", type=Path, required=True)
    process.add_argument("--out", type=Path, required=True)
    process.add_argument("--no-vision", action="store_true", help="Skip vision review pack generation.")
    process.add_argument("--metadata-store", type=Path, default=None, help="Global student metadata JSON table.")

    batch = subparsers.add_parser("batch-process", help="Run process for every .doc/.docx in a directory.")
    batch.add_argument("--template", type=Path, required=True)
    batch.add_argument("--inputs", type=Path, required=True)
    batch.add_argument("--out", type=Path, required=True)
    batch.add_argument("--no-vision", action="store_true")
    batch.add_argument("--metadata-store", type=Path, default=None, help="Global student metadata JSON table.")

    samples = subparsers.add_parser("list-samples", help="List bundled sample files.")
    samples.add_argument("--root", type=Path, default=Path("samples"))

    vision_pack = subparsers.add_parser("vision-pack", help="Create contact sheets and a prompt for VLM visual review.")
    vision_pack.add_argument("--audit-dir", type=Path, required=True)
    vision_pack.add_argument("--out", type=Path, required=True)
    vision_pack.add_argument("--thumb-width", type=int, default=620)
    vision_pack.add_argument("--pages-per-sheet", type=int, default=6)

    web = subparsers.add_parser("web", help="Start the local chat-style web UI.")
    web.add_argument("--host", default="127.0.0.1")
    web.add_argument("--port", type=int, default=8765)

    args = parser.parse_args(argv)
    config = AgentConfig.load(args.config)
    toolchain = Toolchain.discover()

    if args.command == "doctor":
        return _doctor(toolchain)
    if args.command == "audit":
        result = run_audit(args.template, args.target, args.out, config, toolchain)
        print(json.dumps({"status": result["status"], "report": str(args.out / "report.md")}, ensure_ascii=False))
        return 0
    if args.command == "profile-template":
        profile = build_template_profile(args.template)
        payload = profile.to_json()
        if args.out:
            args.out.parent.mkdir(parents=True, exist_ok=True)
            args.out.write_text(payload + "\n", encoding="utf-8")
            print(json.dumps({"profile": str(args.out)}, ensure_ascii=False))
        else:
            print(payload)
        return 0
    if args.command == "template-checklist":
        items = write_red_text_checklist(args.template, args.out, args.out.with_suffix(".json"))
        print(json.dumps({"items": len(items), "checklist": str(args.out), "json": str(args.out.with_suffix(".json"))}, ensure_ascii=False))
        return 0
    if args.command == "fix-format":
        fix_report = fix_docx_format(args.target, args.output, template_path=args.template)
        payload = {"fix": _jsonable_dataclass(fix_report)}
        if args.template and args.audit_out:
            result = run_audit(args.template, args.output, args.audit_out, config, toolchain)
            payload["audit"] = {"status": result["status"], "report": str(args.audit_out / "report.md")}
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 0
    if args.command == "rebuild":
        report = rebuild_thesis_docx(args.template, args.target, args.output)
        print(report.to_json())
        return 0
    if args.command == "fill-template":
        report = fill_standard_template_docx(args.template, args.target, args.output)
        print(report.to_json())
        return 0
    if args.command == "template-rebuild":
        report = rebuild_standard_template(args.cover, args.body, args.output)
        payload = {"rebuild": _jsonable_dataclass(report)}
        if args.strip_red:
            stripped = args.output.with_name(args.output.stem + "-formal.docx")
            strip_report = strip_red_annotations_from_docx(args.output, stripped)
            payload["strip_red"] = _jsonable_dataclass(strip_report)
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 0
    if args.command == "template-selftest":
        args.out.mkdir(parents=True, exist_ok=True)
        rebuilt = args.out / "standard_template.docx"
        rebuild_report = rebuild_standard_template(args.cover, args.body, rebuilt)
        (args.out / "rebuild_report.json").write_text(rebuild_report.to_json() + "\n", encoding="utf-8")
        visual = compare_standard_template_visual(args.cover, args.body, rebuilt, args.out / "visual_compare", config, toolchain)
        payload = {
            "rebuilt": str(rebuilt),
            "rebuild_report": str(args.out / "rebuild_report.json"),
            "visual_report": str(args.out / "visual_compare" / "report.md"),
            "visual_passed": visual.passed,
        }
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 0 if visual.passed else 1
    if args.command == "process":
        result = process_document(
            args.template,
            args.target,
            args.out,
            config,
            toolchain,
            build_vision=not args.no_vision,
            metadata_store_path=args.metadata_store or default_metadata_store(args.out),
        )
        print(json.dumps(_jsonable_dataclass(result), ensure_ascii=False, indent=2))
        return 0
    if args.command == "batch-process":
        args.out.mkdir(parents=True, exist_ok=True)
        results = []
        metadata_store = args.metadata_store or default_metadata_store(args.out)
        for target in iter_supported_inputs(args.inputs):
            try:
                relative_stem = target.relative_to(args.inputs).with_suffix("")
            except ValueError:
                relative_stem = Path(target.stem)
            out_dir = args.out / _slug(str(relative_stem))
            result = process_document(
                args.template,
                target,
                out_dir,
                config,
                toolchain,
                build_vision=not args.no_vision,
                metadata_store_path=metadata_store,
            )
            results.append(_jsonable_dataclass(result))
        (args.out / "batch_result.json").write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps({"count": len(results), "result": str(args.out / "batch_result.json")}, ensure_ascii=False))
        return 0
    if args.command == "list-samples":
        for path in sorted(args.root.rglob("*")):
            if path.is_file():
                print(path)
        return 0
    if args.command == "vision-pack":
        pack = build_vision_pack(args.audit_dir, args.out, args.thumb_width, args.pages_per_sheet)
        print(json.dumps(_jsonable_dataclass(pack), ensure_ascii=False, indent=2))
        return 0
    if args.command == "web":
        return serve_web_app(args.host, args.port)
    raise AssertionError(args.command)


def _doctor(toolchain: Toolchain) -> int:
    payload = {
        "tools": toolchain.as_dict(),
        "versions": {
            "soffice": command_version(toolchain.soffice, ["--version"]),
            "officecli": command_version(toolchain.officecli, ["--version"]),
        },
        "fonts": {
            "宋体": probe_font("宋体"),
            "黑体": probe_font("黑体"),
            "Times New Roman": probe_font("Times New Roman"),
        },
        "missing_for_render": toolchain.missing_for_render(),
    }
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 1 if payload["missing_for_render"] else 0


def _slug(value: str) -> str:
    slug = re.sub(r"[^\w\u4e00-\u9fff.-]+", "-", value).strip("-")
    return slug[:80] or "thesis"


def _jsonable_dataclass(value):
    if hasattr(value, "__dataclass_fields__"):
        return {key: _jsonable_dataclass(getattr(value, key)) for key in value.__dataclass_fields__}
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, list):
        return [_jsonable_dataclass(item) for item in value]
    return value
