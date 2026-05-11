import zipfile
from pathlib import Path
import re
from xml.etree import ElementTree as ET

import pytest

from thesis_agent.annotations import strip_red_annotations_from_docx
from thesis_agent.config import AgentConfig
from thesis_agent.content_enhance import enhance_docx_content
from thesis_agent.content_review import review_content
from thesis_agent.document_profile import PROPOSAL, TASK_BOOK, build_document_profile, detect_document_kind
from thesis_agent.document_requirements import inspect_document_requirements
from thesis_agent.docx_inspect import W_NS, _paragraph_text, inspect_docx
from thesis_agent.format_fix import (
    _apply_caption_keep_rules,
    _normalize_authorization_title_offset,
    _normalize_caption_format,
    fix_docx_format,
)
from thesis_agent.quality_gate import evaluate_quality_gate
from thesis_agent.metadata_store import apply_metadata_to_docx, extract_document_metadata
from thesis_agent.rebuild import rebuild_thesis_docx
from thesis_agent.slot_fill import (
    TemplateSlots,
    _fallback_title_from_document,
    _filled_body_elements,
    _replace_cover_underline_value,
    _unique_reference_items,
    fill_standard_template_docx,
    final_output_filename,
)
from thesis_agent.template_profile import build_template_profile
from thesis_agent.template_checklist import extract_red_text_checklist
from thesis_agent.template_rebuild import rebuild_standard_template, template_text_digest
from thesis_agent.tools import Toolchain
from thesis_agent.visual_check import (
    _detect_front_matter_layout_errors,
    _detect_front_matter_page_number_errors,
    _detect_split_toc_title,
    _detect_toc_page_number_mismatches,
    _infer_page_number_label,
    _is_expected_sparse_page,
)
from thesis_agent.visual_repair import _formula_paragraph, _simple_flowchart_png
from thesis_agent.vision_pack import build_vision_pack


ROOT = Path(__file__).resolve().parents[1]
TEMPLATE_DIR = ROOT / "samples" / "templates"
DRAFT_DIR = ROOT / "samples" / "drafts"
pytestmark = pytest.mark.skipif(
    not (ROOT / "samples").exists(),
    reason="local sample thesis documents are not committed to the public repository",
)


def _format_template(suffix: str = ".doc") -> Path:
    candidates = [
        path
        for path in sorted(TEMPLATE_DIR.glob(f"*{suffix}"))
        if "Zone.Identifier" not in path.name and not path.name.startswith("~$") and not _looks_like_cover_template(path)
    ]
    if not candidates:
        pytest.skip(f"private format template {suffix} is not available")
    return candidates[0]


def _cover_template() -> Path:
    candidates = [
        path
        for path in sorted(TEMPLATE_DIR.glob("*.docx"))
        if "Zone.Identifier" not in path.name and not path.name.startswith("~$") and _looks_like_cover_template(path)
    ]
    if not candidates:
        candidates = [
            path
            for path in sorted(TEMPLATE_DIR.glob("*.docx"))
            if "Zone.Identifier" not in path.name and not path.name.startswith("~$") and path != _format_template(".docx")
        ]
    if not candidates:
        pytest.skip("private cover template is not available")
    return candidates[0]


def _primary_draft() -> Path:
    candidates = _draft_candidates()
    if not candidates:
        pytest.skip("private thesis draft is not available")
    return candidates[0]


def _draft_needing_front_matter() -> Path:
    required = ("学术诚信声明", "AI使用情况声明", "版权使用授权书")
    for path in _draft_candidates():
        compact = _docx_compact_text(path)
        if not all(marker in compact for marker in required):
            return path
    pytest.skip("no private draft without complete front matter is available")


def _draft_candidates() -> list[Path]:
    return [path for path in sorted(DRAFT_DIR.glob("*.docx")) if not path.name.startswith("~$")]


def _docx_compact_text(path: Path) -> str:
    try:
        with zipfile.ZipFile(path) as zf:
            xml = zf.read("word/document.xml").decode("utf-8")
    except Exception:
        return ""
    text = re.sub(r"<[^>]+>", "", xml)
    return re.sub(r"\s+", "", text)


def _looks_like_cover_template(path: Path) -> bool:
    return any(token in path.stem.lower() for token in {"cover", "封面"})


def test_config_loads():
    config = AgentConfig.load(ROOT / "configs" / "default_format.json")
    assert config.renderer_dpi >= 72
    assert "参考文献" in config.expected_sections
    assert "致谢" in config.expected_sections


def test_template_profile_extracts_front_matter():
    profile = build_template_profile(_format_template(".doc"))
    assert profile.docx_source and profile.docx_source.suffix == ".docx"
    assert profile.has_front_matter
    assert "毕业设计（论文）学术诚信声明" in profile.front_matter_titles


def test_template_checklist_extracts_red_requirements():
    items = extract_red_text_checklist(_format_template(".doc"))
    text = "\n".join(item.text for item in items)
    assert "一级标题小二号黑体居中" in text
    assert "每一章另起页" in text


def test_docx_inspection_reads_sample():
    target = _primary_draft()
    inspection = inspect_docx(target)
    assert inspection.supported
    assert len(inspection.paragraphs) > 100


def test_content_review_flags_incomplete_text():
    config = AgentConfig.load(ROOT / "configs" / "default_format.json")
    review = review_content("摘 要\n关键词：物联网\n1 绪论\n参考文献\n[1] Test.", config)
    assert review.score < 100
    assert any(issue.code == "few-references" for issue in review.issues)


def test_content_review_handles_uppercase_abstract_and_toc_entries():
    config = AgentConfig.load(ROOT / "configs" / "default_format.json")
    text = """
    摘 要
    ABSTRACT
    1 绪论....................................................1
    2 系统分析................................................2
    3 系统设计................................................3
    4 系统测试................................................4
    5 总结与展望..............................................5
    参考文献..................................................6
    1 绪论
    2 系统分析
    3 系统设计
    4 系统测试
    5 总结与展望
    参考文献
    [1] Author. MQTT Standard[S]. OASIS, 2014.
    [2] 作者. 物联网系统设计[J]. 电子技术, 2024.
    """
    review = review_content(text, config)
    missing = [issue.message for issue in review.issues if issue.code == "missing-section"]
    assert not any("Abstract" in item or "目 录" in item for item in missing)


def test_content_review_accepts_conclusion_alias():
    config = AgentConfig.load(ROOT / "configs" / "default_format.json")
    text = """
    摘 要
    Abstract
    目 录
    1 绪论
    2 系统分析
    3 系统设计
    4 系统测试
    5 结论
    参考文献
    [1] Author. MQTT Standard[S]. OASIS, 2014.
    """
    review = review_content(text, config)
    missing = [issue.message for issue in review.issues if issue.code == "missing-section"]
    assert not any("5 总结与展望" in item for item in missing)


def test_content_review_counts_unnumbered_references_and_loose_citations():
    config = AgentConfig.load(ROOT / "configs" / "default_format.json")
    refs = "\n".join(
        f"作者{i}. 地下停车场空气监测系统研究[J]. 物联网技术, 2025, ({i:02d}): 1-4."
        for i in range(1, 11)
    )
    text = f"""
    摘 要
    Abstract
    目 录
    1 绪论
    研究背景已有相关工作[1]，[2]，[6,][9]，[10]。
    2 系统总体设计
    需求分析、功能分析、系统架构与器件选型。
    3 系统硬件设计
    传感器、单片机和通信模块设计。
    4 系统测试
    测试环境、测试用例、测试结果与结果分析。
    5 总结与展望
    参考文献
    {refs}
    """
    review = review_content(text, config)
    assert review.reference_count == 10
    assert not any(issue.code == "few-references" for issue in review.issues)
    assert not any(issue.code == "low-citation-coverage" for issue in review.issues)


def test_content_review_finds_test_chapter_by_title_not_number():
    config = AgentConfig.load(ROOT / "configs" / "default_format.json")
    text = """
    摘 要
    Abstract
    目 录
    1 绪论
    2 系统总体设计
    3 系统硬件设计
    4 系统软件设计
    5 系统调试
    测试环境、测试用例、测试结果与结果分析。
    """ + "功能测试结果稳定。" * 120 + """
    6 总结与展望
    参考文献
    作者1. 系统测试方法研究[J]. 物联网技术, 2025, (01): 1-4.
    作者2. 系统测试方法研究[J]. 物联网技术, 2025, (02): 1-4.
    作者3. 系统测试方法研究[J]. 物联网技术, 2025, (03): 1-4.
    作者4. 系统测试方法研究[J]. 物联网技术, 2025, (04): 1-4.
    作者5. 系统测试方法研究[J]. 物联网技术, 2025, (05): 1-4.
    作者6. 系统测试方法研究[J]. 物联网技术, 2025, (06): 1-4.
    作者7. 系统测试方法研究[J]. 物联网技术, 2025, (07): 1-4.
    作者8. 系统测试方法研究[J]. 物联网技术, 2025, (08): 1-4.
    作者9. 系统测试方法研究[J]. 物联网技术, 2025, (09): 1-4.
    作者10. 系统测试方法研究[J]. 物联网技术, 2025, (10): 1-4.
    """
    review = review_content(text, config)
    assert not any(issue.code == "thin-test-chapter" for issue in review.issues)
    assert not any(issue.code == "weak-test-method" for issue in review.issues)


def test_content_review_ignores_tabbed_toc_entries_when_finding_chapters():
    config = AgentConfig.load(ROOT / "configs" / "default_format.json")
    toc = """
    1 绪论\t1
    2 系统总体设计\t5
    3 系统硬件设计\t9
    4 系统软件设计\t13
    5 系统调试\t21
    6 总结与展望\t28
    """
    text = f"""
    摘 要
    Abstract
    目 录
    {toc}
    1 绪论
    2 系统总体设计
    3 系统硬件设计
    4 系统软件设计
    5 系统调试
    测试环境、测试用例、测试结果与结果分析。
    """ + "功能测试结果稳定。" * 120 + """
    6 总结与展望
    参考文献
    作者1. 系统测试方法研究[J]. 物联网技术, 2025, (01): 1-4.
    作者2. 系统测试方法研究[J]. 物联网技术, 2025, (02): 1-4.
    作者3. 系统测试方法研究[J]. 物联网技术, 2025, (03): 1-4.
    作者4. 系统测试方法研究[J]. 物联网技术, 2025, (04): 1-4.
    作者5. 系统测试方法研究[J]. 物联网技术, 2025, (05): 1-4.
    作者6. 系统测试方法研究[J]. 物联网技术, 2025, (06): 1-4.
    作者7. 系统测试方法研究[J]. 物联网技术, 2025, (07): 1-4.
    作者8. 系统测试方法研究[J]. 物联网技术, 2025, (08): 1-4.
    作者9. 系统测试方法研究[J]. 物联网技术, 2025, (09): 1-4.
    作者10. 系统测试方法研究[J]. 物联网技术, 2025, (10): 1-4.
    """
    review = review_content(text, config)
    assert not any(issue.code == "thin-test-chapter" for issue in review.issues)


def test_toolchain_discovery_is_stable():
    toolchain = Toolchain.discover()
    assert set(toolchain.as_dict()) == {"soffice", "pdftoppm", "pdftotext", "pdfinfo", "officecli"}


def test_fix_format_writes_docx_and_reports_changes(tmp_path):
    target = _primary_draft()
    output = tmp_path / "fixed.docx"
    report = fix_docx_format(target, output)
    assert output.exists()
    assert report.removed_trailing_empty_paragraphs + report.removed_trailing_section_paragraphs > 0
    assert report.page_number_restart_applied
    assert report.heading_styles_applied >= 3
    fixed = inspect_docx(output)
    assert fixed.supported
    assert not fixed.main_heading_format_errors


def test_fix_format_can_insert_template_front_matter(tmp_path):
    target = _draft_needing_front_matter()
    template = _format_template(".doc")
    output = tmp_path / "fixed-with-frontmatter.docx"
    report = fix_docx_format(target, output, template_path=template)
    assert output.exists()
    assert report.inserted_front_matter_paragraphs >= 40
    assert report.toc_headings_fixed >= 1
    fixed = inspect_docx(output)
    text = fixed.text.replace(" ", "")
    assert "毕业设计（论文）学术诚信声明" in text
    assert "毕业设计（论文）AI使用情况声明" in text
    assert "毕业设计（论文）版权使用授权书" in text


def test_rebuild_uses_template_package_and_source_content(tmp_path):
    target = _primary_draft()
    template = _format_template(".doc")
    output = tmp_path / "rebuilt.docx"
    report = rebuild_thesis_docx(template, target, output)
    assert output.exists()
    assert report.cover_elements > 0
    assert report.abstract_elements > 0
    assert report.body_elements > 0
    assert report.toc_entries >= 8
    rebuilt = inspect_docx(output)
    text = "".join(rebuilt.text.split())
    assert "毕业设计（论文）学术诚信声明" in text
    assert "毕业设计（论文）学术诚信声明" in text
    assert "1绪论" in text
    with zipfile.ZipFile(output) as zf:
        document_xml = zf.read("word/document.xml").decode("utf-8", errors="ignore")
    assert 'TOC \\o "1-2" \\u' not in document_xml


def test_standard_template_rebuild_combines_cover_and_body_templates(tmp_path):
    cover = _cover_template()
    body = _format_template(".docx")
    output = tmp_path / "standard-template.docx"
    report = rebuild_standard_template(cover, body, output)
    assert output.exists()
    assert report.cover_elements > 0
    assert report.body_elements > 0
    assert report.imported_relationships > 0
    text = template_text_digest(output)
    assert "学士学位论文" in text
    assert "论文题目" in text
    assert "毕业设计（论文）AI使用情况声明" in text
    assert "毕业设计（论文）版权使用授权书" in text
    with zipfile.ZipFile(output) as zf:
        document_xml = zf.read("word/document.xml").decode("utf-8", errors="ignore")
    root_tag = re.search(r"<w:document[^>]+>", document_xml).group(0)
    ignorable = re.search(r":Ignorable=\"([^\"]+)\"", root_tag)
    assert ignorable
    assert all(f"xmlns:{prefix}=" in root_tag for prefix in ignorable.group(1).split())


def test_red_annotations_can_be_removed_from_standard_template(tmp_path):
    cover = _cover_template()
    body = _format_template(".docx")
    output = tmp_path / "standard-template.docx"
    formal = tmp_path / "standard-template-formal.docx"
    rebuild_standard_template(cover, body, output)
    report = strip_red_annotations_from_docx(output, formal)
    assert report.red_shapes_removed > 0
    with zipfile.ZipFile(formal) as zf:
        document_xml = zf.read("word/document.xml").decode("utf-8", errors="ignore")
    assert 'strokecolor="red"' not in document_xml
    assert "FF0000" not in document_xml


def test_slot_fill_uses_formal_template_slots_and_source_content(tmp_path):
    cover = _cover_template()
    body = _format_template(".docx")
    standard = tmp_path / "standard-template.docx"
    formal = tmp_path / "standard-template-formal.docx"
    rebuild_standard_template(cover, body, standard)
    strip_red_annotations_from_docx(standard, formal)

    target = _primary_draft()
    output = tmp_path / "slot-filled.docx"
    report = fill_standard_template_docx(formal, target, output)
    assert output.exists()
    assert final_output_filename(report).endswith(".docx")
    assert len(final_output_filename(report).removesuffix(".docx").split("-")) == 3
    assert final_output_filename(type("R", (), {"student_id": None, "student_name": None, "title": None})()) == "XX-XX-XX.docx"
    assert report.toc_entries >= 8
    assert report.body_elements > 50
    inspection = inspect_docx(output)
    text = inspection.text
    assert not inspection.cover_format_errors
    assert "学士学位论文" in text
    assert report.title
    assert "论文题目\n关键词" not in text
    assert "学士学位论文\n关键词" not in text
    assert "均质充量压缩着火" not in text
    assert "毕业设计（论文）AI使用情况声明" in text
    assert "毕业设计（论文）版权使用授权书" in text


def test_template_preserving_format_fix_keeps_front_matter_and_single_chapter_breaks(tmp_path):
    cover = _cover_template()
    body = _format_template(".docx")
    standard = tmp_path / "standard-template.docx"
    formal = tmp_path / "standard-template-formal.docx"
    rebuild_standard_template(cover, body, standard)
    strip_red_annotations_from_docx(standard, formal)

    target = _primary_draft()
    filled = tmp_path / "slot-filled.docx"
    fixed = tmp_path / "slot-fixed.docx"
    fill_standard_template_docx(formal, target, filled)
    fix_docx_format(filled, fixed, template_path=formal, preserve_template_front_matter=True)
    inspection = inspect_docx(fixed)
    assert not inspection.cover_format_errors
    assert not any("显式分页符和 pageBreakBefore" in item for item in inspection.main_heading_format_errors)
    assert not any("图名上方应存在对应图片" in item for item in inspection.caption_format_errors)
    assert not inspection.table_format_errors
    assert not inspection.orphan_empty_paragraph_errors

    assert "学生姓名：" in inspection.text
    assert "专    业：" in inspection.text
    assert "日期：     年   月   日" in inspection.text


def test_reference_blank_lines_are_hard_gated_and_fixed(tmp_path):
    source = tmp_path / "references-with-gap.docx"
    fixed = tmp_path / "references-fixed.docx"
    _write_minimal_docx(
        source,
        [
            "参考文献",
            "[1] 第一条参考文献[J]. 测试, 2024.",
            "",
            "[2] Second reference[EB/OL]. 2024.",
            "致谢",
        ],
    )
    before = inspect_docx(source)
    assert any("空段" in error for error in before.reference_format_errors)
    report = fix_docx_format(source, fixed)
    assert report.reference_empty_paragraphs_removed == 1
    after = inspect_docx(fixed)
    assert not after.reference_format_errors


def test_orphan_empty_body_paragraphs_are_hard_gated_and_removed(tmp_path):
    source = tmp_path / "orphan-empty-body.docx"
    fixed = tmp_path / "orphan-empty-body-fixed.docx"
    _write_raw_document_docx(
        source,
        """
        <w:p><w:r><w:t>1 绪论</w:t></w:r></w:p>
        <w:p>
          <w:bookmarkStart w:id="1" w:name="_TocBad"/>
          <w:bookmarkEnd w:id="1"/>
        </w:p>
        <w:p><w:r><w:t>1.1 研究背景</w:t></w:r></w:p>
        <w:p></w:p>
        <w:p><w:r><w:t>正文内容。</w:t></w:r></w:p>
        <w:p><w:r><w:t>参考文献</w:t></w:r></w:p>
        """,
    )
    before = inspect_docx(source)
    assert before.orphan_empty_paragraph_errors
    fix_docx_format(source, fixed)
    after = inspect_docx(fixed)
    assert not after.orphan_empty_paragraph_errors


def test_reference_field_codes_are_hard_gated_and_flattened(tmp_path):
    source = tmp_path / "references-with-field.docx"
    fixed = tmp_path / "references-field-fixed.docx"
    _write_raw_document_docx(
        source,
        """
        <w:p><w:r><w:t>参考文献</w:t></w:r></w:p>
        <w:p>
          <w:bookmarkStart w:id="9" w:name="_RefTest"/>
          <w:r><w:t>[1] Web reference. </w:t></w:r>
          <w:r><w:fldChar w:fldCharType="begin"/></w:r>
          <w:r><w:instrText xml:space="preserve"> HYPERLINK "https://example.com" </w:instrText></w:r>
          <w:r><w:fldChar w:fldCharType="separate"/></w:r>
          <w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>https://example.com</w:t></w:r>
          <w:r><w:fldChar w:fldCharType="end"/></w:r>
          <w:bookmarkEnd w:id="9"/>
        </w:p>
        <w:p><w:r><w:t>致谢</w:t></w:r></w:p>
        """,
    )
    before = inspect_docx(source)
    assert any("超链接域" in error for error in before.reference_format_errors)
    report = fix_docx_format(source, fixed)
    assert report.reference_field_codes_removed == 1
    after = inspect_docx(fixed)
    assert not after.reference_format_errors
    with zipfile.ZipFile(fixed) as zf:
        document_xml = zf.read("word/document.xml").decode("utf-8")
    assert "fldChar" not in document_xml
    assert "instrText" not in document_xml
    assert "rStyle" not in document_xml
    assert "bookmarkStart" in document_xml
    assert "bookmarkEnd" in document_xml


def test_reference_punctuation_spacing_is_hard_gated_and_fixed(tmp_path):
    source = tmp_path / "references-tight-spacing.docx"
    fixed = tmp_path / "references-spacing-fixed.docx"
    _write_minimal_docx(
        source,
        [
            "参考文献",
            "[1] 于一帆.RFID技术在危险化学品安全管理中的应用研究[J].安全、健康和环境,2021,21(08):50-53.",
            "[2] 李庚泽,王卫华,杨帅栋,等.基于RFID的智能仓储管理系统设计[J].自动化与仪表,2025,40(09):149-154+159.DOI:10.19557/j.cnki.1001-9944.2025.09.029.",
            "致谢",
        ],
    )
    before = inspect_docx(source)
    assert any("著录标点后应按模板保留必要空格" in error for error in before.reference_format_errors)
    fix_docx_format(source, fixed)
    after = inspect_docx(fixed)
    assert not after.reference_format_errors
    assert "于一帆. RFID技术" in after.text
    assert "[J]. 安全、健康和环境, 2021" in after.text
    assert ". DOI:10.19557" in after.text


def test_reference_urls_and_bad_punctuation_are_hard_gated_and_fixed(tmp_path):
    source = tmp_path / "references-with-url.docx"
    fixed = tmp_path / "references-url-fixed.docx"
    _write_minimal_docx(
        source,
        [
            "参考文献",
            "[1] Z丶learn. mqtt协议简介[OL]. 2021.10.13. https://example.com/very/long/path",
            "致谢",
        ],
    )
    before = inspect_docx(source)
    assert any("裸 URL" in error for error in before.reference_format_errors)
    assert any("异常昵称" in error for error in before.reference_format_errors)
    fix_docx_format(source, fixed)
    after = inspect_docx(fixed)
    assert not after.reference_format_errors
    assert "https://" not in after.text
    assert "Z丶learn" not in after.text
    assert "OASIS. MQTT Version 5.0[S]. 2019." in after.text


def test_reference_number_gap_is_hard_gated_and_fixed(tmp_path):
    source = tmp_path / "references-wide-gap.docx"
    fixed = tmp_path / "references-wide-gap-fixed.docx"
    _write_minimal_docx(
        source,
        [
            "参考文献",
            "[1]  第一条参考文献[J]. 测试, 2024.",
            "[2]\tSecond reference[Z]. 2024.",
            "致谢",
        ],
    )
    before = inspect_docx(source)
    assert any("编号后应按模板保留一个半角空格" in error for error in before.reference_format_errors)
    fix_docx_format(source, fixed)
    after = inspect_docx(fixed)
    assert not after.reference_format_errors
    assert "[1] 第一条参考文献" in after.text
    assert "[2] Second reference" in after.text


def test_known_component_references_are_rewritten_to_canonical_sources(tmp_path):
    source = tmp_path / "references-components.docx"
    fixed = tmp_path / "references-components-fixed.docx"
    _write_minimal_docx(
        source,
        [
            "参考文献",
            "[1] GeeksMan. 温湿度传感器-DHT11[EB/OL]. (2025-07-27). https://example.com",
            "[2] SGP30 Datasheet[EB/OL]. Sensirion, (2024).",
            "致谢",
        ],
    )
    fix_docx_format(source, fixed)
    after = inspect_docx(fixed)
    assert "Aosong Electronics Co., Ltd. DHT11 temperature and humidity sensor datasheet[Z]." in after.text
    assert "Sensirion AG. SGP30 multi-pixel gas sensor datasheet[Z]." in after.text
    assert not after.reference_format_errors


def test_reference_normalization_is_idempotent_and_deduplicates():
    items = _unique_reference_items(
        [
            "MQTT Version 3.1.1 Specification[EB/OL]. OASIS Standard, (2014).",
            "OASIS. MQTT Version 5.0[S]. 2019.",
            "scaleway. IoT Hub: A Quick Introduction to the MQTT Protocol[EB/OL]. [2022-04-20].",
            "PCBMaY. DHT11与DHT22传感器温度和湿度教程[EB/OL]. (2024-01-09).",
            "IC Components. DHT11与DHT22：温度和湿度传感器的全面比较[EB/OL]. (2024-11-04).",
        ]
    )
    assert items == [
        "OASIS. MQTT Version 3.1.1[S]. 2014.",
        "OASIS. MQTT Version 5.0[S]. 2019.",
        "Aosong Electronics Co., Ltd. DHT11 and AM2302/DHT22 digital temperature and humidity sensor datasheets[Z]. Guangzhou: Aosong Electronics Co., Ltd., 2015.",
    ]


def test_cover_underline_value_is_centered():
    paragraph = ET.fromstring(
        f'<w:p xmlns:w="{W_NS}"><w:r><w:t>学生姓名：____________________</w:t></w:r></w:p>'
    )
    assert _replace_cover_underline_value(paragraph, "杨钰婷")
    text = "".join(node.text or "" for node in paragraph.findall(".//w:t", {"w": W_NS}))
    assert text.startswith("学生姓名：")
    assert "杨钰婷" in text
    assert "_" not in text
    underlined_runs = [
        run
        for run in paragraph.findall(".//w:r", {"w": W_NS})
        if run.find("w:rPr/w:u", {"w": W_NS}) is not None
    ]
    assert len(underlined_runs) == 1
    underlined_text = "".join(node.text or "" for node in underlined_runs[0].findall(".//w:t", {"w": W_NS}))
    assert underlined_text == "\u00A0" * 14 + "杨钰婷" + "\u00A0" * 14


def test_cover_blank_underline_uses_generated_underlined_spaces():
    paragraph = ET.fromstring(
        f'<w:p xmlns:w="{W_NS}"><w:r><w:t>学生学号：____________________</w:t></w:r></w:p>'
    )
    assert _replace_cover_underline_value(paragraph, "")
    text = "".join(node.text or "" for node in paragraph.findall(".//w:t", {"w": W_NS}))
    assert text.startswith("学生学号：")
    assert "_" not in text
    underlined_runs = [
        run
        for run in paragraph.findall(".//w:r", {"w": W_NS})
        if run.find("w:rPr/w:u", {"w": W_NS}) is not None
    ]
    assert len(underlined_runs) == 1
    assert "".join(node.text or "" for node in underlined_runs[0].findall(".//w:t", {"w": W_NS})) == "\u00A0" * 40


def test_figure_caption_does_not_keep_with_following_heading():
    body = ET.fromstring(
        f"""
        <w:body xmlns:w="{W_NS}">
          <w:p><w:r><w:drawing /></w:r></w:p>
          <w:p>
            <w:pPr><w:keepNext /></w:pPr>
            <w:r><w:t>图 3-3 系统架构流程图</w:t></w:r>
          </w:p>
          <w:p><w:r><w:t>3.3 系统硬件设计</w:t></w:r></w:p>
        </w:body>
        """
    )
    image_paragraph, caption_paragraph, _heading = list(body)

    _normalize_caption_format(body)
    _apply_caption_keep_rules(body)

    assert image_paragraph.find("w:pPr/w:keepNext", {"w": W_NS}) is not None
    assert caption_paragraph.find("w:pPr/w:keepLines", {"w": W_NS}) is not None
    assert caption_paragraph.find("w:pPr/w:keepNext", {"w": W_NS}) is None
    assert image_paragraph.find("w:pPr/w:jc", {"w": W_NS}).attrib.get(f"{{{W_NS}}}val") == "center"


def test_authorization_title_offset_keeps_two_template_blank_lines():
    body = ET.fromstring(
        f"""
        <w:body xmlns:w="{W_NS}">
          <w:p><w:pPr><w:spacing w:line="240" /></w:pPr></w:p>
          <w:p><w:r><w:t>上海电机学院</w:t></w:r></w:p>
          <w:p><w:r><w:t>毕业设计（论文）版权使用授权书</w:t></w:r></w:p>
        </w:body>
        """
    )

    _normalize_authorization_title_offset(body)

    children = list(body)
    school_idx = next(idx for idx, child in enumerate(children) if _paragraph_text(child) == "上海电机学院")
    blank_before_school = [child for child in children[:school_idx] if child.tag == f"{{{W_NS}}}p" and not _paragraph_text(child)]
    assert len(blank_before_school) == 2
    assert all(child.find("w:pPr/w:spacing", {"w": W_NS}).attrib.get(f"{{{W_NS}}}line") == "480" for child in blank_before_school)


def test_document_kind_routes_non_thesis_templates():
    proposal = Path("samples/drafts/3-开题/物联网2212-221003710619-蔡宇璐-开题报告第二版.docx")
    task_book = Path("samples/drafts/2-任务书/物联网2212-221003710619-蔡宇璐-任务书-面向冷链物流的温控追踪系统设计.docx")
    requested = Path("samples/templates/论文格式.doc")

    assert detect_document_kind(proposal) == PROPOSAL
    assert detect_document_kind(task_book) == TASK_BOOK
    assert "附件13" in build_document_profile(proposal, requested).template.name
    assert "附件16" in build_document_profile(task_book, requested).template.name


def test_metadata_store_can_fill_blank_material_fields(tmp_path):
    source = tmp_path / "blank-fields.docx"
    fixed = tmp_path / "filled-fields.docx"
    _write_minimal_docx(
        source,
        [
            "学生姓名：________",
            "学号：________",
            "课题名称：________",
        ],
    )

    filled = apply_metadata_to_docx(
        source,
        fixed,
        {
            "student_name": "蔡宇璐",
            "student_id": "221003710619",
            "title": "面向冷链物流的温控追踪系统设计",
        },
    )

    assert filled == 3
    text = inspect_docx(fixed).text
    assert "学生姓名：蔡宇璐" in text
    assert "学号：221003710619" in text
    assert "课题名称：面向冷链物流的温控追踪系统设计" in text
    assert extract_document_metadata(Path("蔡宇璐-开题报告.docx")).get("student_name") == "蔡宇璐"


def test_document_requirements_gate_blocks_signature_and_empty_opinion(tmp_path):
    source = tmp_path / "proposal-missing-signature.docx"
    _write_minimal_docx(
        source,
        [
            "毕业设计（论文）开题报告",
            "学生签名：",
            "指导教师意见：",
        ],
    )

    requirements = inspect_document_requirements(source, PROPOSAL, {"student_name": "蔡宇璐"})

    assert requirements.signature_errors
    assert requirements.opinion_errors


def test_formula_paragraph_uses_subscript_and_superscript_runs():
    paragraph = _formula_paragraph("ppm = 11.5428 × (R0 / RS)^0.6549", "3", 4)
    texts = [node.text or "" for node in paragraph.findall(".//w:t", {"w": W_NS})]
    assert "0" in texts
    assert "S" in texts
    assert "0.6549" in texts
    subscript_texts = [
        "".join(node.text or "" for node in run.findall(".//w:t", {"w": W_NS}))
        for run in paragraph.findall(".//w:r", {"w": W_NS})
        if (align := run.find("w:rPr/w:vertAlign", {"w": W_NS})) is not None
        and align.attrib.get(f"{{{W_NS}}}val") == "subscript"
    ]
    superscript_texts = [
        "".join(node.text or "" for node in run.findall(".//w:t", {"w": W_NS}))
        for run in paragraph.findall(".//w:r", {"w": W_NS})
        if (align := run.find("w:rPr/w:vertAlign", {"w": W_NS})) is not None
        and align.attrib.get(f"{{{W_NS}}}val") == "superscript"
    ]
    assert {"0", "S"}.issubset(set(subscript_texts))
    assert "0.6549" in superscript_texts
    assert paragraph.findall(".//w:tab", {"w": W_NS})


def test_redrawn_simple_flowchart_is_trimmed():
    from PIL import Image
    import io

    image = Image.open(io.BytesIO(_simple_flowchart_png("测试需求图", [("上", "下"), ("左", "右"), ("前", "后")]))).convert("RGB")
    assert image.height < 560


def test_malformed_inline_citations_are_hard_gated_and_fixed(tmp_path):
    source = tmp_path / "bad-citations.docx"
    fixed = tmp_path / "bad-citations-fixed.docx"
    _write_minimal_docx(
        source,
        [
            "1 绪论",
            "系统以STM32F103C8T6单片机作为主控核心[2,[7]]，并通过Wi-Fi模块完成联网。",
            "参考文献",
            "[1] 作者. 文献[J]. 期刊, 2024.",
        ],
    )
    before = inspect_docx(source)
    assert any("正文引用编号格式异常" in error for error in before.body_paragraph_format_errors)
    enhance_docx_content(source, fixed)
    after = inspect_docx(fixed)
    assert "[2,7]" in after.text
    assert not any("正文引用编号格式异常" in error for error in after.body_paragraph_format_errors)


def test_caption_residual_indent_is_hard_gated_and_fixed(tmp_path):
    source = tmp_path / "caption-with-indent.docx"
    fixed = tmp_path / "caption-indent-fixed.docx"
    _write_raw_document_docx(
        source,
        """
        <w:p>
          <w:pPr><w:ind w:left="2520" w:firstLine="420"/><w:jc w:val="center"/></w:pPr>
          <w:r><w:t>表 3-1 常见温湿度传感器对比</w:t></w:r>
        </w:p>
        <w:tbl><w:tr><w:tc><w:p><w:r><w:t>型号</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
        """,
    )
    before = inspect_docx(source)
    assert any("不应残留左右缩进" in error for error in before.caption_format_errors)
    fix_docx_format(source, fixed)
    after = inspect_docx(fixed)
    assert not after.caption_format_errors


def test_figure_visual_paragraph_is_centered_with_caption(tmp_path):
    source = tmp_path / "uncentered-figure.docx"
    fixed = tmp_path / "centered-figure.docx"
    _write_raw_document_docx(
        source,
        """
        <w:p>
          <w:pPr><w:jc w:val="left"/></w:pPr>
          <w:r><w:drawing /></w:r>
        </w:p>
        <w:p>
          <w:pPr><w:jc w:val="center"/></w:pPr>
          <w:r><w:t>图 2-2 系统功能结构图</w:t></w:r>
        </w:p>
        """,
    )

    before = inspect_docx(source)
    assert any("图名上方图片应居中" in error for error in before.caption_format_errors)
    fix_docx_format(source, fixed)
    after = inspect_docx(fixed)
    assert not any("图名上方图片应居中" in error for error in after.caption_format_errors)


def test_title_fallback_does_not_use_abstract_sentence():
    text = (
        "摘要\n"
        "本文设计并实现了一套基于物联网的实验室危化品智能监管系统。"
        "系统以STM32为核心控制器。"
    )
    title = _fallback_title_from_document(text, Path("物联网2212-学生-毕业论文初稿.docx"))
    assert title == "基于物联网的实验室危化品智能监管系统设计"


def test_slot_fill_preserves_visual_when_figure_caption_shares_paragraph():
    def p(text: str) -> ET.Element:
        return ET.fromstring(f'<w:p xmlns:w="{W_NS}"><w:r><w:t>{text}</w:t></w:r></w:p>')

    body_template = p("正文内容")
    caption_template = p("图 1-1 示例图")
    slots = TemplateSlots(
        front=[],
        zh_abstract_heading=p("摘 要"),
        zh_abstract_body=body_template,
        zh_keywords=p("关键词：测试"),
        en_abstract_heading=p("ABSTRACT"),
        en_abstract_body=body_template,
        en_keywords=p("Key words: test"),
        toc_heading=p("目  录"),
        toc_entry_level1=p("1 绪论"),
        toc_entry_level2=p("1.1 背景"),
        toc_section_break=None,
        first_main_heading=p("1 绪论"),
        later_main_heading=p("2 设计"),
        second_heading=p("1.1 背景"),
        third_heading=p("1.1.1 背景"),
        body_paragraph=body_template,
        figure_caption=caption_template,
        table_caption=p("表 1-1 示例表"),
        reference_heading=p("参考文献"),
        reference_item=p("[1] 文献."),
        acknowledgement_heading=p("致  谢"),
        acknowledgement_body=body_template,
        final_section=None,
    )
    source = ET.fromstring(
        f"""
        <w:p xmlns:w="{W_NS}">
          <w:r><w:drawing /></w:r>
          <w:r><w:t>图3-1单片机最小系统板</w:t></w:r>
        </w:p>
        """
    )
    elements = _filled_body_elements(slots, [source])
    assert len(elements) == 2
    assert elements[0].find(".//w:drawing", {"w": W_NS}) is not None
    assert not _paragraph_text(elements[0]).strip()
    assert _paragraph_text(elements[1]).strip() == "图 3-1 单片机最小系统板"


def test_slot_fill_preserves_orphan_figure_caption_without_generating_visual():
    def p(text: str) -> ET.Element:
        return ET.fromstring(f'<w:p xmlns:w="{W_NS}"><w:r><w:t>{text}</w:t></w:r></w:p>')

    body_template = p("正文内容")
    slots = TemplateSlots(
        front=[],
        zh_abstract_heading=p("摘 要"),
        zh_abstract_body=body_template,
        zh_keywords=p("关键词：测试"),
        en_abstract_heading=p("ABSTRACT"),
        en_abstract_body=body_template,
        en_keywords=p("Key words: test"),
        toc_heading=p("目  录"),
        toc_entry_level1=p("1 绪论"),
        toc_entry_level2=p("1.1 背景"),
        toc_section_break=None,
        first_main_heading=p("1 绪论"),
        later_main_heading=p("2 设计"),
        second_heading=p("1.1 背景"),
        third_heading=p("1.1.1 背景"),
        body_paragraph=body_template,
        figure_caption=p("图 1-1 示例图"),
        table_caption=p("表 1-1 示例表"),
        reference_heading=p("参考文献"),
        reference_item=p("[1] 文献."),
        acknowledgement_heading=p("致  谢"),
        acknowledgement_body=body_template,
        final_section=None,
    )
    elements = _filled_body_elements(slots, [p("图3-1单片机最小系统板")])
    assert len(elements) == 1
    assert elements[0].find(".//w:drawing", {"w": W_NS}) is None
    assert _paragraph_text(elements[0]).strip() == "图 3-1 单片机最小系统板"


def test_toc_heading_field_codes_are_hard_gated_and_removed(tmp_path):
    source = tmp_path / "toc-field.docx"
    fixed = tmp_path / "toc-field-fixed.docx"
    _write_raw_document_docx(
        source,
        """
        <w:p><w:r><w:t>摘  要</w:t></w:r></w:p>
        <w:p>
          <w:r><w:t>目  录</w:t></w:r>
          <w:r><w:fldChar w:fldCharType="begin"/></w:r>
          <w:r><w:instrText xml:space="preserve"> TOC \\o "1-2" \\u </w:instrText></w:r>
          <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        </w:p>
        <w:p><w:r><w:t>1 绪论</w:t></w:r></w:p>
        """,
    )
    before = inspect_docx(source)
    assert any("TOC 域代码" in error for error in before.toc_format_errors)
    report = fix_docx_format(source, fixed)
    assert report.front_matter_page_breaks_inserted >= 1
    after = inspect_docx(fixed)
    assert not any("TOC 域代码" in error for error in after.toc_format_errors)
    with zipfile.ZipFile(fixed) as zf:
        document_xml = zf.read("word/document.xml").decode("utf-8")
    assert "TOC" not in document_xml
    assert "instrText" not in document_xml


def test_toc_entries_require_number_spacing_and_right_tab(tmp_path):
    source = tmp_path / "toc-bad-layout.docx"
    _write_raw_document_docx(
        source,
        """
        <w:p><w:r><w:t>目  录</w:t></w:r></w:p>
        <w:p>
          <w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8300"/></w:tabs></w:pPr>
          <w:r><w:t>1.1标题</w:t><w:tab/><w:t>1</w:t></w:r>
        </w:p>
        <w:p><w:r><w:t>1 绪论</w:t></w:r></w:p>
        """,
    )
    errors = inspect_docx(source).toc_format_errors
    assert any("编号和标题之间应有空格" in item for item in errors)
    assert any("版心右侧" in item for item in errors)


def test_caption_anchor_and_table_body_format_are_hard_gated(tmp_path):
    source = tmp_path / "bad-caption-table.docx"
    _write_raw_document_docx(
        source,
        """
        <w:p><w:r><w:t>正文段落</w:t></w:r></w:p>
        <w:p><w:r><w:t>图 3-1 缺图题注</w:t></w:r></w:p>
        <w:p><w:r><w:t>表 3-1 表题</w:t></w:r></w:p>
        <w:tbl>
          <w:tblPr><w:jc w:val="left"/></w:tblPr>
          <w:tr><w:tc><w:p><w:r><w:rPr><w:sz w:val="14"/></w:rPr><w:t>型号</w:t></w:r></w:p></w:tc></w:tr>
        </w:tbl>
        """,
    )
    inspection = inspect_docx(source)
    assert any("图名上方应存在对应图片" in item for item in inspection.caption_format_errors)
    assert any("表格本体应居中" in item for item in inspection.table_format_errors)
    assert any("不应继承学生原稿字号" in item for item in inspection.table_format_errors)


def test_front_matter_page_boundaries_are_inserted(tmp_path):
    source = tmp_path / "mixed-front.docx"
    fixed = tmp_path / "mixed-front-fixed.docx"
    _write_minimal_docx(
        source,
        [
            "学士学位论文",
            "测试系统设计",
            "某某大学",
            "毕业设计（论文）学术诚信声明",
            "本人郑重声明。",
            "某某大学",
            "毕业设计（论文）AI使用情况声明",
            "本人承诺。",
            "摘  要",
            "摘要正文",
            "目  录",
            "1 绪论",
        ],
    )
    report = fix_docx_format(source, fixed)
    assert report.front_matter_page_breaks_inserted >= 4


def test_content_enhance_does_not_insert_unlinked_generic_paragraphs(tmp_path):
    target = _draft_needing_front_matter()
    formatted = tmp_path / "formatted.docx"
    enhanced = tmp_path / "enhanced.docx"
    fix_docx_format(target, formatted, template_path=_format_template(".doc"))
    report = enhance_docx_content(formatted, enhanced)
    assert not report.acknowledgements_inserted
    assert not report.test_chapter_augmented
    assert report.inserted_paragraphs == 0
    text = inspect_docx(enhanced).text.replace(" ", "")
    assert "测试环境与结果分析" not in text
    assert "该图展示了" not in text


def _write_minimal_docx(path: Path, paragraphs: list[str]) -> None:
    body = []
    for text in paragraphs:
        if text:
            body.append(f"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>")
        else:
            body.append("<w:p/>")
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W_NS}"><w:body>{"".join(body)}<w:sectPr/></w:body></w:document>'
    )
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("word/document.xml", document)


def _write_raw_document_docx(path: Path, body_xml: str) -> None:
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W_NS}"><w:body>{body_xml}<w:sectPr/></w:body></w:document>'
    )
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("word/document.xml", document)


def test_page_label_detection_ignores_student_ids():
    assert _infer_page_number_label("专业班级-000000000000\n论文题目") is None
    assert _infer_page_number_label("论文题目                                 12") == "12"
    assert _infer_page_number_label("系统设计                                                                  33") == "33"


def test_split_toc_title_detection():
    assert _detect_split_toc_title({1: "摘要\n目", 2: "录\n1 绪论....1"}) == [1]


def test_visual_sparse_chapter_tail_before_references_is_not_blank_page():
    text = "面向冷链物流的温控追踪系统设计         25\n智能化方向发展贡献一份更加坚实的技术力量。"
    next_text = "面向冷链物流的温控追踪系统设计         26\n参考文献\n[1] 作者. 文献[J]. 期刊, 2024."
    assert _is_expected_sparse_page(text, next_text)


def test_visual_checks_block_front_digits_and_stale_toc():
    page_texts = {
        1: "摘 要\nI",
        2: "目 录\n1 绪论........................4\nII",
        3: "1 绪论\n1",
    }
    labels = {1: "I", 2: "II", 3: "1"}
    assert _detect_front_matter_page_number_errors({**page_texts, 2: "目 录\n1 绪论........................4\n2"}, {1: "I", 2: "2", 3: "1"}) == {2: "2"}
    mismatches = _detect_toc_page_number_mismatches(page_texts, labels)
    assert any("1 绪论" in item and "目录 4" in item and "正文 1" in item for item in mismatches)


def test_visual_front_layout_flags_compressed_academic_date_page():
    errors = _detect_front_matter_layout_errors({1: "毕业设计（论文）学术诚信声明\n作者签名：\n日期： 年 月 日"})
    assert any("学术诚信声明日期页位置异常" in item for item in errors)


def test_visual_toc_mismatch_detects_entries_without_dot_leaders():
    page_texts = {
        1: "目 录\n5.2 移动端 App 设计43\n6 系统测试36\nVIII",
        2: "5.2 移动端 App 设计\n33",
        3: "6 系统测试\n36",
    }
    labels = {1: "VIII", 2: "33", 3: "36"}
    mismatches = _detect_toc_page_number_mismatches(page_texts, labels)
    assert any("5.2 移动端 App 设计" in item and "目录 43" in item and "正文 33" in item for item in mismatches)


def test_vision_pack_builds_contact_sheet(tmp_path):
    from PIL import Image

    audit = tmp_path / "audit"
    target = audit / "png" / "target"
    template = audit / "png" / "template"
    target.mkdir(parents=True)
    template.mkdir(parents=True)
    for idx in range(1, 4):
        Image.new("RGB", (200, 280), "white").save(target / f"page-{idx:02d}.png")
    Image.new("RGB", (200, 280), "white").save(template / "page-01.png")
    (audit / "report.json").write_text(
        '{"target_visual":{"caption_orphan_pages":[2],"heading_pages":{"^1":[3]}}}',
        encoding="utf-8",
    )
    pack = build_vision_pack(audit, tmp_path / "vision", thumb_width=120, pages_per_sheet=2)
    assert pack.target_sheets
    assert pack.template_sheets
    assert pack.prompt and pack.prompt.exists()
    assert any(path.name == "page-02.png" for path in pack.key_pages)


def test_quality_gate_blocks_visual_and_content_errors():
    class Visual:
        blank_pages = []
        near_blank_pages = []
        broken_reference_pages = []
        toc_title_split_pages = [8]
        caption_orphan_pages = []

    class Docx:
        supported = True
        empty_paragraph_runs = []
        broken_references = []

    class Issue:
        severity = "warning"
        message = "测试章节偏薄"

    class Content:
        score = 90
        issues = [Issue()]

    result = evaluate_quality_gate({"target_visual": Visual(), "docx": Docx(), "content_review": Content()})
    assert not result.passed
    assert any("目录标题拆页" in blocker for blocker in result.blockers)


def test_quality_gate_passes_only_when_warning_score_stays_above_threshold():
    class Visual:
        blank_pages = []
        near_blank_pages = []
        broken_reference_pages = []
        toc_title_split_pages = []
        caption_orphan_pages = []

    class Docx:
        supported = True
        empty_paragraph_runs = []
        broken_references = []

    class Issue:
        severity = "warning"
        message = "测试章节偏薄"

    class Content:
        score = 90
        issues = [Issue()]

    result = evaluate_quality_gate({"target_visual": Visual(), "docx": Docx(), "content_review": Content()})
    assert result.passed
    assert result.format_score == 100
    assert result.content_score == 90
    assert result.score == 90
    assert not result.blockers
    assert "测试章节偏薄" in result.warnings


def test_quality_gate_blocks_low_warning_score():
    class Visual:
        blank_pages = []
        near_blank_pages = []
        broken_reference_pages = []
        toc_title_split_pages = []
        caption_orphan_pages = []
        figure_readability_warnings = ["图片文字偏小"] * 4

    class Docx:
        supported = True
        empty_paragraph_runs = []
        broken_references = []

    class Issue:
        severity = "warning"
        message = "图表说明不足"

    class Content:
        score = 75
        issues = [Issue()]

    result = evaluate_quality_gate({"target_visual": Visual(), "docx": Docx(), "content_review": Content()})
    assert not result.passed
    assert any("图像可读性不足" in blocker for blocker in result.blockers)
    assert any("内容分 75/100" in blocker for blocker in result.blockers)


def test_quality_gate_blocks_material_requirement_errors():
    class Visual:
        blank_pages = []
        near_blank_pages = []
        broken_reference_pages = []
        toc_title_split_pages = []
        caption_orphan_pages = []
        figure_readability_warnings = []

    class Docx:
        supported = True
        empty_paragraph_runs = []
        broken_references = []

    class Content:
        score = 100
        issues = []

    class Requirements:
        required_field_errors = ["学生姓名: 信息应填写完整"]
        signature_errors = ["学生签名: 需要签名图片"]
        opinion_errors = ["指导教师意见: 意见内容应填写完整"]
        metadata_warnings = []

    result = evaluate_quality_gate(
        {"target_visual": Visual(), "docx": Docx(), "content_review": Content(), "document_requirements": Requirements()}
    )

    assert not result.passed
    assert any("必填信息不完整" in blocker for blocker in result.blockers)
    assert any("签名图片不符合要求" in blocker for blocker in result.blockers)
    assert any("意见区内容不完整" in blocker for blocker in result.blockers)


def test_quality_gate_requires_separate_format_and_content_scores():
    class Visual:
        blank_pages = []
        near_blank_pages = []
        broken_reference_pages = []
        toc_title_split_pages = []
        caption_orphan_pages = []
        figure_readability_warnings = []

    class Docx:
        supported = True
        empty_paragraph_runs = []
        broken_references = []
        toc_format_errors = ["目录混入正文"]

    class Content:
        score = 95
        issues = []

    result = evaluate_quality_gate({"target_visual": Visual(), "docx": Docx(), "content_review": Content()})
    assert not result.passed
    assert result.format_score == 85
    assert result.content_score == 95
    assert any("目录格式不符合模板" in blocker for blocker in result.format_blockers)


def test_content_review_blocks_auto_text_and_reference_pollution():
    config = AgentConfig(
        path=Path("test"),
        data={
            "expected_sections": [],
            "content_quality": {
                "min_body_chinese_chars": 0,
                "recommended_body_chinese_chars": 0,
                "min_reference_count": 0,
                "min_foreign_reference_count": 0,
                "min_keywords": 0,
                "max_keywords": 10,
                "required_engineering_topics": [],
                "forbidden_markers": [],
            },
        },
    )
    text = "\n".join(
        [
            "目  录",
            "1 绪论\t1",
            "为进一步说明安全监管系统的验证过程，本章测试不仅关注单个功能是否能够运行，还需要关注模块之间的数据传递是否连续。",
            "1 绪论",
            "图 3-9 rc522相关代码",
            "参考文献",
            "[1] 综合全文来看，论文还需要把需求、设计、实现和测试之间的对应关系表达得更加清楚。",
        ]
    )
    review = review_content(text, config)
    codes = {issue.code for issue in review.issues}
    assert {"review-text-pollution", "toc-pollution", "reference-pollution", "code-formula-as-image"} <= codes
    assert review.score < 80


def test_docx_inspect_blocks_toc_pollution_and_code_formula_image(tmp_path):
    docx_path = tmp_path / "polluted.docx"
    _write_raw_document_docx(
        docx_path,
        "".join(
            [
                '<w:p><w:r><w:t>目  录</w:t></w:r></w:p>',
                '<w:p><w:r><w:t>为进一步说明安全监管系统的验证过程，本章测试不仅关注单个功能是否能够运行，还需要关注模块之间的数据传递是否连续。</w:t></w:r></w:p>',
                '<w:p><w:r><w:t>1 绪论</w:t></w:r></w:p>',
                '<w:p><w:r><w:drawing/></w:r></w:p>',
                '<w:p><w:r><w:t>图 3-9 rc522相关代码</w:t></w:r></w:p>',
                '<w:p><w:r><w:t>参考文献</w:t></w:r></w:p>',
                '<w:p><w:r><w:t>[1] 综合全文来看，论文还需要把需求、设计、实现和测试之间的对应关系表达得更加清楚。</w:t></w:r></w:p>',
            ]
        ),
    )
    inspection = inspect_docx(docx_path)
    assert any("目录区域疑似混入正文" in item for item in inspection.toc_format_errors)
    assert any("代码或公式不应作为图片保留" in item for item in inspection.caption_format_errors)
    assert any("参考文献区域疑似混入正文" in item for item in inspection.reference_format_errors)
