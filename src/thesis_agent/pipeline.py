from __future__ import annotations

import json
import shutil
from dataclasses import asdict, is_dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

from .config import AgentConfig
from .content_review import review_content
from .docx_inspect import inspect_docx
from .pdf_inspect import extract_text
from .render import render_document
from .tools import Toolchain
from .visual_check import inspect_visual


def run_audit(template: Path, target: Path, out_dir: Path, config: AgentConfig, toolchain: Toolchain) -> dict[str, Any]:
    missing = toolchain.missing_for_render()
    if missing:
        raise RuntimeError(f"Missing renderer tools: {', '.join(missing)}")
    if not template.exists():
        raise FileNotFoundError(template)
    if not target.exists():
        raise FileNotFoundError(target)

    out_dir.mkdir(parents=True, exist_ok=True)
    inputs_dir = out_dir / "inputs"
    pdf_dir = out_dir / "pdf"
    png_dir = out_dir / "png"
    inputs_dir.mkdir(exist_ok=True)
    pdf_dir.mkdir(exist_ok=True)
    png_dir.mkdir(exist_ok=True)

    template_input = inputs_dir / f"template_input{template.suffix}"
    target_input = inputs_dir / f"target_input{target.suffix}"
    shutil.copy2(template, template_input)
    shutil.copy2(target, target_input)

    rendered_template = render_document(
        template_input,
        pdf_dir,
        png_dir / "template",
        toolchain,
        dpi=config.renderer_dpi,
    )
    rendered_target = render_document(
        target_input,
        pdf_dir,
        png_dir / "target",
        toolchain,
        dpi=config.renderer_dpi,
    )

    template_visual = inspect_visual(rendered_template.pdf, rendered_template.pages, toolchain, config)
    target_visual = inspect_visual(rendered_target.pdf, rendered_target.pages, toolchain, config)
    docx = inspect_docx(target_input)
    pdf_text = extract_text(rendered_target.pdf, toolchain)
    target_text = docx.text if docx.supported and docx.text.strip() else pdf_text
    content = review_content(target_text, config)

    result: dict[str, Any] = {
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "config": str(config.path),
        "tools": toolchain.as_dict(),
        "files": {
            "template": str(template),
            "target": str(target),
            "template_input": str(template_input),
            "target_input": str(target_input),
            "template_pdf": str(rendered_template.pdf),
            "target_pdf": str(rendered_target.pdf),
        },
        "template_visual": template_visual,
        "target_visual": target_visual,
        "docx": docx,
        "content_review": content,
        "status": _audit_status(target_visual, docx, content),
    }

    (out_dir / "report.json").write_text(
        json.dumps(_jsonable(result), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    (out_dir / "report.md").write_text(_markdown_report(result), encoding="utf-8")
    return result


def _audit_status(target_visual: Any, docx: Any, content: Any) -> str:
    if target_visual.blank_pages or target_visual.near_blank_pages:
        return "needs_fix"
    if target_visual.broken_reference_pages:
        return "needs_fix"
    if target_visual.toc_title_split_pages:
        return "needs_fix"
    if getattr(target_visual, "front_matter_page_number_errors", {}):
        return "needs_fix"
    if getattr(target_visual, "toc_page_number_mismatches", []):
        return "needs_fix"
    if getattr(target_visual, "front_matter_layout_errors", []):
        return "needs_fix"
    if getattr(target_visual, "header_page_number_alignment_errors", []):
        return "needs_fix"
    if target_visual.caption_orphan_pages:
        return "needs_fix"
    if docx.supported and docx.broken_references:
        return "needs_fix"
    if docx.supported and getattr(docx, "orphan_empty_paragraph_errors", []):
        return "needs_fix"
    if docx.supported and getattr(docx, "cover_format_errors", []):
        return "needs_fix"
    if docx.supported and getattr(docx, "abstract_format_errors", []):
        return "needs_fix"
    if docx.supported and getattr(docx, "main_heading_format_errors", []):
        return "needs_fix"
    if docx.supported and getattr(docx, "sub_heading_format_errors", []):
        return "needs_fix"
    if docx.supported and getattr(docx, "toc_format_errors", []):
        return "needs_fix"
    if docx.supported and getattr(docx, "body_paragraph_format_errors", []):
        return "needs_fix"
    if docx.supported and getattr(docx, "table_format_errors", []):
        return "needs_fix"
    if docx.supported and getattr(docx, "caption_format_errors", []):
        return "needs_fix"
    if docx.supported and getattr(docx, "reference_format_errors", []):
        return "needs_fix"
    if docx.supported and getattr(docx, "acknowledgement_format_errors", []):
        return "needs_fix"
    if any(issue.severity == "error" for issue in content.issues):
        return "needs_fix"
    return "review"


def _jsonable(value: Any) -> Any:
    if isinstance(value, Path):
        return str(value)
    if is_dataclass(value):
        return _jsonable(asdict(value))
    if isinstance(value, dict):
        return {str(key): _jsonable(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_jsonable(item) for item in value]
    return value


def _markdown_report(result: dict[str, Any]) -> str:
    template_visual = result["template_visual"]
    target_visual = result["target_visual"]
    docx = result["docx"]
    content = result["content_review"]

    lines: list[str] = []
    lines.append("# Thesis Agent Audit Report")
    lines.append("")
    lines.append(f"- Created: {result['created_at']}")
    lines.append(f"- Status: `{result['status']}`")
    lines.append(f"- Template: `{result['files']['template']}`")
    lines.append(f"- Target: `{result['files']['target']}`")
    lines.append("")
    lines.append("## Render")
    lines.append("")
    lines.append(f"- Template pages: {template_visual.pages}, page size: {template_visual.page_size or 'unknown'}")
    lines.append(f"- Target pages: {target_visual.pages}, page size: {target_visual.page_size or 'unknown'}")
    lines.append(f"- Target blank pages: {_fmt_list(target_visual.blank_pages)}")
    lines.append(f"- Target near blank pages: {_fmt_list(target_visual.near_blank_pages)}")
    lines.append(f"- First page labels: {_fmt_page_labels(target_visual.page_number_labels, limit=8)}")
    lines.append(f"- Broken reference pages: {_fmt_list(target_visual.broken_reference_pages)}")
    lines.append(f"- Split TOC title pages: {_fmt_list(target_visual.toc_title_split_pages)}")
    front_errors = getattr(target_visual, "front_matter_page_number_errors", {})
    if front_errors:
        lines.append("- Front-matter page number errors: " + ", ".join(f"{page}:{label}" for page, label in sorted(front_errors.items())))
    else:
        lines.append("- Front-matter page number errors: none")
    lines.append(f"- TOC page number mismatches: {_fmt_list(getattr(target_visual, 'toc_page_number_mismatches', []))}")
    lines.append(f"- Front-matter layout errors: {_fmt_list(getattr(target_visual, 'front_matter_layout_errors', []))}")
    lines.append(f"- Header page number alignment errors: {_fmt_list(getattr(target_visual, 'header_page_number_alignment_errors', []))}")
    lines.append(f"- Possible caption orphan pages: {_fmt_list(target_visual.caption_orphan_pages)}")
    lines.append(f"- Figure readability warnings: {_fmt_list(getattr(target_visual, 'figure_readability_warnings', []))}")
    lines.append("")
    lines.append("## Heading Page Map")
    lines.append("")
    if target_visual.heading_pages:
        for pattern, pages in target_visual.heading_pages.items():
            labels = [target_visual.page_number_labels.get(page, "?") for page in pages]
            lines.append(f"- `{pattern}`: {_fmt_list(pages)} (labels: {_fmt_list(labels)})")
    else:
        lines.append("- No configured heading patterns were found after TOC pages.")
    lines.append("")
    lines.append("## DOCX Structure")
    lines.append("")
    if docx.supported:
        lines.append(f"- Paragraphs: {len(docx.paragraphs)}")
        lines.append(f"- Headings: {len(docx.headings)}")
        lines.append(f"- Captions: {len(docx.captions)}")
        lines.append(f"- Explicit page breaks: {docx.explicit_page_breaks}")
        lines.append(f"- Section breaks: {docx.section_breaks}")
        lines.append(f"- Empty paragraph runs >= 4: {_fmt_list(docx.empty_paragraph_runs)}")
        lines.append(f"- Orphan empty paragraph errors: {_fmt_list(getattr(docx, 'orphan_empty_paragraph_errors', []))}")
        lines.append(f"- Broken references in DOCX text: {len(docx.broken_references)}")
        lines.append(f"- Cover format errors: {_fmt_list(getattr(docx, 'cover_format_errors', []))}")
        lines.append(f"- Abstract/keyword format errors: {_fmt_list(getattr(docx, 'abstract_format_errors', []))}")
        lines.append(f"- TOC format errors: {_fmt_list(getattr(docx, 'toc_format_errors', []))}")
        lines.append(f"- Main heading format errors: {_fmt_list(getattr(docx, 'main_heading_format_errors', []))}")
        lines.append(f"- Sub heading format errors: {_fmt_list(getattr(docx, 'sub_heading_format_errors', []))}")
        lines.append(f"- Body paragraph format errors: {_fmt_list(getattr(docx, 'body_paragraph_format_errors', []))}")
        lines.append(f"- Table format errors: {_fmt_list(getattr(docx, 'table_format_errors', []))}")
        lines.append(f"- Caption format errors: {_fmt_list(getattr(docx, 'caption_format_errors', []))}")
        lines.append(f"- Reference format errors: {_fmt_list(getattr(docx, 'reference_format_errors', []))}")
        lines.append(f"- Acknowledgement format errors: {_fmt_list(getattr(docx, 'acknowledgement_format_errors', []))}")
    else:
        lines.append("- DOCX inspection skipped because the target is not `.docx`.")
    lines.append("")
    lines.append("## Content Review")
    lines.append("")
    lines.append(f"- Score: {content.score}/100")
    lines.append(f"- Chinese chars: {content.chinese_chars}")
    lines.append(f"- References: {content.reference_count}")
    lines.append(f"- Foreign references: {content.foreign_reference_count}")
    lines.append(f"- Web references: {content.web_reference_count}")
    lines.append(f"- Keywords: {content.keyword_count if content.keyword_count is not None else 'not detected'}")
    if content.chapter_char_counts:
        lines.append("- Chapter Chinese chars: " + ", ".join(f"{key}={value}" for key, value in content.chapter_char_counts.items()))
    lines.append("")
    if content.issues:
        lines.append("### Issues")
        lines.append("")
        for issue in content.issues:
            lines.append(f"- `{issue.severity}` `{issue.code}`: {issue.message}")
    else:
        lines.append("No content issues detected by the rule baseline.")
    lines.append("")
    lines.append("## Next Fix Targets")
    lines.append("")
    next_targets = _next_targets(target_visual, docx, content)
    for item in next_targets:
        lines.append(f"- {item}")
    return "\n".join(lines) + "\n"


def _next_targets(target_visual: Any, docx: Any, content: Any) -> list[str]:
    items: list[str] = []
    if target_visual.near_blank_pages:
        items.append(f"Remove or explain near blank rendered pages: {_fmt_list(target_visual.near_blank_pages)}.")
    if target_visual.broken_reference_pages:
        items.append(f"Repair Word reference fields shown as missing sources on pages: {_fmt_list(target_visual.broken_reference_pages)}.")
    if target_visual.toc_title_split_pages:
        items.append(f"Fix split TOC heading around pages: {_fmt_list(target_visual.toc_title_split_pages)}.")
    front_errors = getattr(target_visual, "front_matter_page_number_errors", {})
    if front_errors:
        items.append("Normalize all pre-body page labels to Roman numerals before `1 绪论`.")
    toc_mismatches = getattr(target_visual, "toc_page_number_mismatches", [])
    if toc_mismatches:
        items.append(f"Refresh static TOC page labels: {_fmt_list(toc_mismatches[:8])}.")
    front_layout_errors = getattr(target_visual, "front_matter_layout_errors", [])
    if front_layout_errors:
        items.append(f"Separate front-matter pages: {_fmt_list(front_layout_errors[:6])}.")
    header_alignment_errors = getattr(target_visual, "header_page_number_alignment_errors", [])
    if header_alignment_errors:
        items.append(f"Normalize header page-number right alignment: {_fmt_list(header_alignment_errors[:6])}.")
    if target_visual.caption_orphan_pages:
        items.append(f"Keep captions with their figure/table objects near pages: {_fmt_list(target_visual.caption_orphan_pages)}.")
    figure_warnings = getattr(target_visual, "figure_readability_warnings", [])
    if figure_warnings:
        items.append(f"Review or redraw small-text figures: {_fmt_list(figure_warnings[:6])}.")
    first_pages = target_visual.heading_pages.get("^1\\s+绪论", [])
    if first_pages:
        first_label = target_visual.page_number_labels.get(first_pages[0])
        if first_label and first_label != "1":
            items.append(f"Reset main-body page numbering: `1 绪论` renders on physical page {first_pages[0]} with page label {first_label}, expected label 1.")
    if docx.supported and docx.empty_paragraph_runs:
        items.append("Replace long empty paragraph runs with controlled page or section breaks.")
    if docx.supported and getattr(docx, "orphan_empty_paragraph_errors", []):
        items.append("Remove isolated empty paragraphs in the main body; spacing must come from paragraph styles, not blank paragraphs.")
    if docx.supported and getattr(docx, "abstract_format_errors", []):
        items.append("Apply red-text abstract and keyword rules: 中英文摘要标题、正文段落、关键词字体和标点。")
    if docx.supported and getattr(docx, "main_heading_format_errors", []):
        items.append("Apply red-text main heading rule: 小二号黑体居中，段前 0 磅，段后 12 磅，每章另起页。")
    if docx.supported and getattr(docx, "sub_heading_format_errors", []):
        items.append("Apply red-text subheading rule: 二级标题宋体四号左对齐，三级标题黑体小四左对齐，段前 12 磅、段后 0 磅。")
    if docx.supported and getattr(docx, "toc_format_errors", []):
        items.append("Apply red-text TOC rule: 标题小二号黑体居中，目录只列一、二级标题，并设置一、二级目录字体。")
    if docx.supported and getattr(docx, "body_paragraph_format_errors", []):
        items.append("Apply red-text body paragraph rule: 中文小四宋体、英文小四 Times New Roman、首行缩进 2 字符、1.25 倍行距。")
    if docx.supported and getattr(docx, "caption_format_errors", []):
        items.append("Apply red-text figure/table caption rule: 图名在图下、表名在表上，宋体小五号加粗，并按章编号。")
    if docx.supported and getattr(docx, "reference_format_errors", []):
        items.append("Apply red-text reference rule: 顺序编号、悬挂缩进 2 字符、小四宋体/Times New Roman、1.25 倍行距。")
    if docx.supported and getattr(docx, "acknowledgement_format_errors", []):
        items.append("Apply red-text acknowledgement rule: 标题小二黑体，正文五号宋体/Times New Roman、首行缩进 2 字符、单倍行距。")
    for issue in content.issues:
        if issue.severity == "error":
            items.append(issue.message)
    for issue in content.issues:
        if issue.severity == "warning" and issue.code in {"thin-test-chapter", "weak-test-method", "low-citation-coverage", "web-reference-heavy"}:
            items.append(issue.message)
    if not items:
        items.append("Move to visual comparison against the template and automated formatting edits.")
    return items


def _fmt_list(values: list[Any]) -> str:
    return ", ".join(str(value) for value in values) if values else "none"


def _fmt_page_labels(labels: dict[int, str], limit: int) -> str:
    if not labels:
        return "none"
    items = sorted(labels.items())[:limit]
    return ", ".join(f"{page}:{label}" for page, label in items)
