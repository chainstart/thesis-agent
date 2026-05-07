from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


@dataclass(frozen=True)
class ParagraphInfo:
    text: str
    style: str | None
    page_breaks: int = 0
    section_break: bool = False


@dataclass(frozen=True)
class DocxInspection:
    path: Path
    supported: bool
    paragraphs: list[ParagraphInfo] = field(default_factory=list)
    headings: list[ParagraphInfo] = field(default_factory=list)
    captions: list[str] = field(default_factory=list)
    broken_references: list[str] = field(default_factory=list)
    explicit_page_breaks: int = 0
    section_breaks: int = 0
    empty_paragraph_runs: list[int] = field(default_factory=list)
    orphan_empty_paragraph_errors: list[str] = field(default_factory=list)
    abstract_format_errors: list[str] = field(default_factory=list)
    cover_format_errors: list[str] = field(default_factory=list)
    toc_format_errors: list[str] = field(default_factory=list)
    main_heading_format_errors: list[str] = field(default_factory=list)
    sub_heading_format_errors: list[str] = field(default_factory=list)
    body_paragraph_format_errors: list[str] = field(default_factory=list)
    table_format_errors: list[str] = field(default_factory=list)
    caption_format_errors: list[str] = field(default_factory=list)
    reference_format_errors: list[str] = field(default_factory=list)
    acknowledgement_format_errors: list[str] = field(default_factory=list)
    red_rule_errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def text(self) -> str:
        return "\n".join(p.text for p in self.paragraphs if p.text)


def inspect_docx(path: Path) -> DocxInspection:
    if path.suffix.lower() != ".docx":
        return DocxInspection(path=path, supported=False, warnings=["Only .docx OOXML inspection is supported."])
    if not zipfile.is_zipfile(path):
        return DocxInspection(path=path, supported=False, warnings=["File is not a valid zip-based docx."])

    with zipfile.ZipFile(path) as zf:
        try:
            xml = zf.read("word/document.xml")
        except KeyError:
            return DocxInspection(path=path, supported=False, warnings=["word/document.xml not found."])

    root = ET.fromstring(xml)
    paragraphs: list[ParagraphInfo] = []
    captions: list[str] = []
    broken_references: list[str] = []
    empty_runs: list[int] = []
    current_empty_run = 0
    seen_main_body = False
    explicit_page_breaks = 0
    section_breaks = 0

    for p in root.findall(".//w:p", NS):
        text = _paragraph_text(p)
        style = _paragraph_style(p)
        page_breaks = len(p.findall(".//w:br[@w:type='page']", NS))
        has_section = p.find(".//w:sectPr", NS) is not None
        explicit_page_breaks += page_breaks
        section_breaks += int(has_section)

        info = ParagraphInfo(text=text, style=style, page_breaks=page_breaks, section_break=has_section)
        paragraphs.append(info)

        if re.match(r"^1\s+绪论\s*$", text.strip()):
            seen_main_body = True

        if not seen_main_body:
            current_empty_run = 0
        elif text.strip():
            if current_empty_run >= 4:
                empty_runs.append(current_empty_run)
            current_empty_run = 0
        else:
            current_empty_run += 1

        if _is_caption(text):
            captions.append(text.strip())
        if "Error: Reference source not found" in text or "错误!未找到引用源" in text:
            broken_references.append(text.strip())

    if current_empty_run >= 4:
        empty_runs.append(current_empty_run)

    headings = [p for p in paragraphs if _is_heading(p)]
    cover_errors = _cover_format_errors(root)
    abstract_errors = _abstract_format_errors(root)
    toc_errors = _toc_format_errors(root)
    heading_errors = _main_heading_format_errors(root)
    sub_heading_errors = _sub_heading_format_errors(root)
    body_errors = _body_paragraph_format_errors(root)
    table_errors = _table_format_errors(root)
    caption_errors = _caption_format_errors(root)
    reference_errors = _reference_format_errors(root)
    ack_errors = _acknowledgement_format_errors(root)
    orphan_empty_errors = _orphan_empty_paragraph_errors(root)
    red_errors = (
        cover_errors
        + abstract_errors
        + toc_errors
        + heading_errors
        + sub_heading_errors
        + body_errors
        + table_errors
        + caption_errors
        + reference_errors
        + ack_errors
        + orphan_empty_errors
    )
    return DocxInspection(
        path=path,
        supported=True,
        paragraphs=paragraphs,
        headings=headings,
        captions=captions,
        broken_references=broken_references,
        explicit_page_breaks=explicit_page_breaks,
        section_breaks=section_breaks,
        empty_paragraph_runs=empty_runs,
        orphan_empty_paragraph_errors=orphan_empty_errors,
        cover_format_errors=cover_errors,
        abstract_format_errors=abstract_errors,
        toc_format_errors=toc_errors,
        main_heading_format_errors=heading_errors,
        sub_heading_format_errors=sub_heading_errors,
        body_paragraph_format_errors=body_errors,
        table_format_errors=table_errors,
        caption_format_errors=caption_errors,
        reference_format_errors=reference_errors,
        acknowledgement_format_errors=ack_errors,
        red_rule_errors=red_errors,
    )


def _paragraph_text(p: ET.Element) -> str:
    pieces: list[str] = []
    for node in p.iter():
        if node.tag == f"{{{W_NS}}}t" and node.text:
            pieces.append(node.text)
        elif node.tag == f"{{{W_NS}}}tab":
            pieces.append("\t")
    return "".join(pieces).strip()


def _cover_format_errors(root: ET.Element) -> list[str]:
    body = root.find("w:body", NS)
    if body is None:
        return []
    errors: list[str] = []
    for child in list(body):
        if child.tag != f"{{{W_NS}}}p":
            continue
        if child.find(".//w:sectPr", NS) is not None:
            break
        text = _paragraph_text(child)
        compact = re.sub(r"\s+", "", text)
        if not any(compact.startswith(prefix) for prefix in ("学生姓名：", "学生姓名:", "学生学号：", "学生学号:", "专业：", "专业:", "指导教师：", "指导教师:", "学院：", "学院:")):
            continue
        if "：" in text:
            value = text.split("：", 1)[1]
        elif ":" in text:
            value = text.split(":", 1)[1]
        else:
            continue
        normalized_value = value.strip()
        if not normalized_value or re.fullmatch(r"_+", normalized_value):
            continue
        if "_" not in value and not _has_underlined_cover_value(child):
            label = compact.split("：", 1)[0].split(":", 1)[0]
            errors.append(f"封面 {label}: 已填字段应保留模板下划线")
    return errors


def _has_underlined_cover_value(paragraph: ET.Element) -> bool:
    seen_colon = False
    for run in paragraph.findall("w:r", NS):
        text = _run_text(run)
        if not seen_colon:
            if "：" in text or ":" in text:
                seen_colon = True
                text = re.split(r"[：:]", text, maxsplit=1)[1]
            else:
                continue
        if not text.strip("_ "):
            continue
        rpr = run.find("w:rPr", NS)
        underline = rpr.find("w:u", NS) if rpr is not None else None
        if underline is not None and underline.attrib.get(f"{{{W_NS}}}val") not in {"0", "none", "false", "False"}:
            return True
    return False


def _paragraph_style(p: ET.Element) -> str | None:
    style = p.find("./w:pPr/w:pStyle", NS)
    if style is None:
        return None
    return style.attrib.get(f"{{{W_NS}}}val")


def _is_heading(p: ParagraphInfo) -> bool:
    style = (p.style or "").lower()
    text = p.text.strip()
    return (
        style.startswith("heading")
        or style.startswith("title")
        or bool(re.match(r"^[1-9]\s+[\u4e00-\u9fffA-Za-z]", text))
        or text in {"摘 要", "摘要", "Abstract", "目 录", "参考文献", "致谢"}
    )


def _is_caption(text: str) -> bool:
    stripped = text.strip()
    return bool(re.match(r"^(图|表)\s*\d+\s*[-－]\s*\d+", stripped))


def _main_heading_format_errors(root: ET.Element) -> list[str]:
    body = root.find("w:body", NS)
    if body is None:
        return []
    errors: list[str] = []
    children = list(body)
    for idx, child in enumerate(children):
        if child.tag != f"{{{W_NS}}}p":
            continue
        text = _paragraph_text(child).strip()
        if not _is_main_chapter_heading(text):
            continue
        compact = re.sub(r"\s+", "", text)
        ppr = child.find("w:pPr", NS)
        if _ppr_value(ppr, "jc", "val") != "center":
            errors.append(f"{text}: 一级标题应居中")
        if _spacing_value(ppr, "before") not in {"0", None}:
            errors.append(f"{text}: 一级标题段前应为 0 磅")
        if _spacing_value(ppr, "after") != "240":
            errors.append(f"{text}: 一级标题段后应为 12 磅")
        if not _heading_looks_black_xiaoer(child):
            errors.append(f"{text}: 一级标题应为小二号黑体")
        if _has_ppr_child(ppr, "pageBreakBefore") and child.find(".//w:br[@w:type='page']", NS) is not None:
            errors.append(f"{text}: 一级标题不应同时含显式分页符和 pageBreakBefore")
        if not compact.startswith("1绪论") and not _starts_new_page(children, idx, child):
            errors.append(f"{text}: 每一章应另起页")
    return errors


def _toc_format_errors(root: ET.Element) -> list[str]:
    errors: list[str] = []
    paragraphs = list(root.iter(f"{{{W_NS}}}p"))
    first_body_pos = _first_main_body_paragraph_pos(paragraphs)
    search_paragraphs = paragraphs[:first_body_pos] if first_body_pos is not None else paragraphs
    in_toc = False
    seen_entries = 0
    for paragraph in search_paragraphs:
        text = _paragraph_text(paragraph).strip()
        compact = re.sub(r"\s+", "", text)
        if compact in {"目录", "目錄"}:
            in_toc = True
            errors.extend(_toc_heading_errors(paragraph, text))
            continue
        parsed = _parse_toc_entry_text(text)
        if parsed is None:
            continue
        in_toc = True
        title, _label = parsed
        seen_entries += 1
        level = _toc_entry_level(title)
        if level > 2:
            errors.append(f"R022 {title}: 目录只允许列出一、二级标题")
            continue
        if re.match(r"^[1-9](?:\.\d+)+\S", title):
            errors.append(f"R022 {title}: 目录编号和标题之间应有空格")
        if not _toc_page_number_tab_ok(paragraph):
            errors.append(f"R022 {title}: 目录页码应使用右对齐制表位到版心右侧")
        expected_size = "28" if level == 1 else "24"
        if not _first_visible_run_matches(paragraph, east_asia="黑体", size=expected_size):
            size_name = "四号" if level == 1 else "小四号"
            errors.append(f"R022 {title}: 目录{level}级标题应为黑体{size_name}")
    if in_toc and seen_entries == 0:
        errors.append("R022 目录未检测到可校验的一、二级目录项")
    return errors[:30]


def _abstract_format_errors(root: ET.Element) -> list[str]:
    body = root.find("w:body", NS)
    if body is None:
        return []
    errors: list[str] = []
    section: str | None = None
    for child in list(body):
        if child.tag != f"{{{W_NS}}}p":
            continue
        text = _paragraph_text(child).strip()
        compact = re.sub(r"\s+", "", text)
        if compact in {"目录", "目錄"} or re.fullmatch(r"1\s+绪论", text):
            break
        if compact == "摘要":
            section = "zh_abstract"
            if _ppr_value(child.find("w:pPr", NS), "jc", "val") != "center" or not _first_visible_run_matches(child, east_asia="黑体", size="36"):
                errors.append("R005 摘要标题应为小二号黑体居中")
            continue
        if compact == "ABSTRACT":
            section = "en_abstract"
            if _ppr_value(child.find("w:pPr", NS), "jc", "val") != "center" or not _first_visible_run_matches(child, east_asia="Times New Roman", ascii_font="Times New Roman", size="36"):
                errors.append("R014/R015 ABSTRACT 标题应为小二号 Times New Roman 居中加黑")
            continue
        if re.match(r"^关键词[:：]", text):
            if not _keyword_text_ok(text, english=False):
                errors.append("R006/R007 中文关键词应逗号分开，最后一个关键词后无标点")
            if not _first_visible_run_matches(child, east_asia="宋体", ascii_font="Times New Roman", size="24"):
                errors.append("R006 中文关键词应为小四号宋体")
            section = None
            continue
        if re.match(r"^Keywords?[:：]", text, re.IGNORECASE):
            if not _keyword_text_ok(text, english=True):
                errors.append("R018/R019 英文关键词应逗号分开，逗号后加空格")
            if not _first_visible_run_matches(child, east_asia="Times New Roman", ascii_font="Times New Roman", size="24"):
                errors.append("R017 英文关键词应为小四号 Times New Roman")
            section = None
            continue
        if not text:
            continue
        if section == "zh_abstract":
            ppr = child.find("w:pPr", NS)
            if not _front_text_paragraph_ok(child, ppr, east_asia="宋体"):
                errors.append(f"R002/R003 {_short_text(text)}: 中文摘要正文应小四宋体、首行缩进 2 字符、段前段后 12 磅、1.25 倍行距")
        elif section == "en_abstract":
            ppr = child.find("w:pPr", NS)
            if not _front_text_paragraph_ok(child, ppr, east_asia="Times New Roman"):
                errors.append(f"R011/R012 {_short_text(text)}: 英文摘要正文应小四 Times New Roman、首行缩进 2 字符、段前段后 12 磅、1.25 倍行距")
        if len(errors) >= 20:
            break
    return errors


def _front_text_paragraph_ok(paragraph: ET.Element, ppr: ET.Element | None, east_asia: str) -> bool:
    return (
        _spacing_value(ppr, "before") == "240"
        and _spacing_value(ppr, "after") == "240"
        and _spacing_value(ppr, "line") == "300"
        and (_indent_value(ppr, "firstLine") == "480" or _indent_value(ppr, "firstLineChars") == "200")
        and _first_visible_run_matches(paragraph, east_asia=east_asia, ascii_font="Times New Roman", size="24")
    )


def _keyword_text_ok(text: str, english: bool) -> bool:
    if english:
        match = re.match(r"^Keywords?[:：]\s*(.+)$", text, re.IGNORECASE)
        if not match:
            return False
        body = match.group(1).strip()
        return bool(body) and not re.search(r"[.;；。]\s*$", body) and not re.search(r",\S", body)
    match = re.match(r"^关键词[:：]\s*(.+)$", text)
    if not match:
        return False
    body = match.group(1).strip()
    return bool(body) and not re.search(r"[.;；。，,]\s*$", body)


def _toc_page_number_tab_ok(paragraph: ET.Element) -> bool:
    tabs = paragraph.find("w:pPr/w:tabs", NS)
    if tabs is None:
        return False
    for tab in tabs.findall("w:tab", NS):
        if tab.attrib.get(f"{{{W_NS}}}val") != "right":
            continue
        try:
            return int(tab.attrib.get(f"{{{W_NS}}}pos", "0")) >= 8800
        except ValueError:
            return False
    return False


def _toc_heading_errors(paragraph: ET.Element, text: str) -> list[str]:
    errors: list[str] = []
    ppr = paragraph.find("w:pPr", NS)
    if _ppr_value(ppr, "jc", "val") != "center":
        errors.append(f"R021 {text}: 目录标题应居中")
    if _spacing_value(ppr, "before") not in {"0", None}:
        errors.append(f"R020 {text}: 目录标题段前应为 0 磅")
    if _spacing_value(ppr, "after") != "240":
        errors.append(f"R020 {text}: 目录标题段后应为 12 磅")
    if not _first_visible_run_matches(paragraph, east_asia="黑体", size="36"):
        errors.append(f"R021 {text}: 目录标题应为小二号黑体")
    if _has_reference_field_or_inline_style(paragraph):
        errors.append(f"R021 {text}: 目录标题不应残留 TOC 域代码或字符样式")
    return errors


def _body_paragraph_format_errors(root: ET.Element) -> list[str]:
    body = root.find("w:body", NS)
    if body is None:
        return []
    children = list(body)
    first_body_idx = _first_main_body_index(children)
    if first_body_idx is None:
        return []
    errors: list[str] = []
    for child in children[first_body_idx + 1:]:
        if child.tag != f"{{{W_NS}}}p":
            continue
        text = _paragraph_text(child).strip()
        compact = re.sub(r"\s+", "", text)
        if compact in {"参考文献", "致谢", "附录"}:
            break
        if not _is_normal_body_text(text):
            continue
        label = _short_text(text)
        ppr = child.find("w:pPr", NS)
        if _has_ppr_child(ppr, "numPr") or _has_ppr_child(ppr, "keepNext") or _has_ppr_child(ppr, "keepLines") or _has_ppr_child(ppr, "pageBreakBefore"):
            errors.append(f"R027/R028 {label}: 正文段落不应带列表、与下段同页或另起页控制")
        if _spacing_value(ppr, "before") not in {"0", None} or _spacing_value(ppr, "after") not in {"0", None}:
            errors.append(f"R027/R028 {label}: 正文段前段后应为 0 磅")
        if _spacing_value(ppr, "line") != "300":
            errors.append(f"R028 {label}: 正文应为 1.25 倍行距")
        if _indent_value(ppr, "firstLine") != "480" and _indent_value(ppr, "firstLineChars") != "200":
            errors.append(f"R028 {label}: 正文应首行缩进 2 字符")
        if _ppr_value(ppr, "jc", "val") not in {"both", None}:
            errors.append(f"R028 {label}: 正文不应居中或右对齐")
        if not _first_visible_run_matches(child, east_asia="宋体", ascii_font="Times New Roman", size="24"):
            errors.append(f"R027 {label}: 正文应为中文小四号宋体、英文小四号 Times New Roman")
        if len(errors) >= 30:
            break
    return errors


def _sub_heading_format_errors(root: ET.Element) -> list[str]:
    body = root.find("w:body", NS)
    if body is None:
        return []
    children = list(body)
    first_body_idx = _first_main_body_index(children)
    if first_body_idx is None:
        return []
    errors: list[str] = []
    for child in children[first_body_idx + 1:]:
        if child.tag != f"{{{W_NS}}}p":
            continue
        text = _paragraph_text(child).strip()
        compact = re.sub(r"\s+", "", text)
        if compact in {"参考文献", "致谢", "附录"}:
            break
        if "\t" in text or "..." in text or "…" in text:
            continue
        if re.match(r"^[1-9]\.\d+\.\d+\s*[\u4e00-\u9fffA-Za-z]", text):
            errors.extend(_single_sub_heading_errors(child, text, level=3))
        elif re.match(r"^[1-9]\.\d+\s*[\u4e00-\u9fffA-Za-z]", text):
            errors.extend(_single_sub_heading_errors(child, text, level=2))
        if len(errors) >= 30:
            break
    return errors


def _single_sub_heading_errors(paragraph: ET.Element, text: str, level: int) -> list[str]:
    errors: list[str] = []
    label = _short_text(text)
    ppr = paragraph.find("w:pPr", NS)
    if _ppr_value(ppr, "jc", "val") not in {"left", None}:
        errors.append(f"R031/R033 {label}: 二、三级标题应左对齐")
    if _spacing_value(ppr, "before") != "240" or _spacing_value(ppr, "after") not in {"0", None}:
        errors.append(f"R031/R033 {label}: 二、三级标题段前应 12 磅、段后应 0 磅")
    if level == 2 and not _first_visible_run_matches(paragraph, east_asia="宋体", ascii_font="Times New Roman", size="28"):
        errors.append(f"R031 {label}: 二级标题应为宋体四号")
    if level == 3 and not _first_visible_run_matches(paragraph, east_asia="黑体", size="24"):
        errors.append(f"R032 {label}: 三级标题应为黑体小四号")
    return errors


def _acknowledgement_format_errors(root: ET.Element) -> list[str]:
    body = root.find("w:body", NS)
    if body is None:
        return []
    errors: list[str] = []
    in_ack = False
    body_count = 0
    for child in list(body):
        if child.tag != f"{{{W_NS}}}p":
            continue
        text = _paragraph_text(child).strip()
        compact = re.sub(r"\s+", "", text)
        if compact == "致谢":
            in_ack = True
            ppr = child.find("w:pPr", NS)
            if _ppr_value(ppr, "jc", "val") != "center":
                errors.append("R058 致谢: 标题应居中")
            if _spacing_value(ppr, "before") not in {"0", None} or _spacing_value(ppr, "after") != "240":
                errors.append("R059 致谢: 标题段前应 0 磅、段后应 12 磅")
            if not _first_visible_run_matches(child, east_asia="黑体", size="36"):
                errors.append("R058 致谢: 标题应为小二号黑体")
            continue
        if not in_ack:
            continue
        if not text:
            continue
        if compact in {"附录"} or _is_main_chapter_heading(text):
            break
        body_count += 1
        label = _short_text(text)
        ppr = child.find("w:pPr", NS)
        if _spacing_value(ppr, "line") != "240":
            errors.append(f"R061 {label}: 致谢正文应为单倍行距")
        if _indent_value(ppr, "firstLine") != "480" and _indent_value(ppr, "firstLineChars") != "200":
            errors.append(f"R061 {label}: 致谢正文应首行缩进 2 字符")
        if _ppr_value(ppr, "jc", "val") not in {"both", None}:
            errors.append(f"R061 {label}: 致谢正文不应居中或右对齐")
        if not _first_visible_run_matches(child, east_asia="宋体", ascii_font="Times New Roman", size="21"):
            errors.append(f"R060 {label}: 致谢正文应为中文五号宋体、英文五号 Times New Roman")
        if len(errors) >= 20:
            break
    if not in_ack:
        errors.append("R058/R062 缺少致谢章节")
    elif body_count == 0:
        errors.append("R062 致谢内容不能为空")
    return errors


def _orphan_empty_paragraph_errors(root: ET.Element) -> list[str]:
    body = root.find("w:body", NS)
    if body is None:
        return []
    children = list(body)
    errors: list[str] = []
    in_body = False
    last_text = ""
    for idx, child in enumerate(children):
        if child.tag != f"{{{W_NS}}}p":
            continue
        text = _paragraph_text(child).strip()
        compact = re.sub(r"\s+", "", text)
        if re.fullmatch(r"1\s+绪论", text):
            in_body = True
        if in_body and compact in {"参考文献", "致谢", "附录"}:
            break
        if not in_body:
            continue
        if text:
            last_text = _short_text(text)
            continue
        if _empty_paragraph_is_layout_content(child):
            continue
        previous_text = last_text or "正文起始"
        next_text = _next_visible_text(children, idx)
        errors.append(f"正文存在多余空段：位于 `{previous_text}` 与 `{next_text or '后续内容'}` 之间")
        if len(errors) >= 20:
            break
    return errors


def _empty_paragraph_is_layout_content(paragraph: ET.Element) -> bool:
    return (
        paragraph.find(".//w:drawing", NS) is not None
        or paragraph.find(".//w:pict", NS) is not None
        or paragraph.find(".//w:object", NS) is not None
        or paragraph.find(".//w:br[@w:type='page']", NS) is not None
        or paragraph.find(".//w:sectPr", NS) is not None
    )


def _next_visible_text(children: list[ET.Element], start_idx: int) -> str | None:
    for child in children[start_idx + 1 :]:
        if child.tag != f"{{{W_NS}}}p":
            return "表格"
        text = _paragraph_text(child).strip()
        if text:
            return _short_text(text)
        if _empty_paragraph_is_layout_content(child):
            return "图片或分页对象"
    return None


def _table_format_errors(root: ET.Element) -> list[str]:
    errors: list[str] = []
    for table_idx, table in enumerate(root.iter(f"{{{W_NS}}}tbl"), start=1):
        tbl_pr = table.find("w:tblPr", NS)
        if _ppr_value(tbl_pr, "jc", "val") != "center":
            errors.append(f"表格 {table_idx}: 表格本体应居中")
        if _has_ppr_child(tbl_pr, "tblStyle") or _has_ppr_child(tbl_pr, "tblLayout"):
            errors.append(f"表格 {table_idx}: 表格不应继承学生原稿表格样式或自适应布局")
        if _table_border_value(tbl_pr, "top") != "single" or _table_border_value(tbl_pr, "bottom") != "single":
            errors.append(f"表格 {table_idx}: 表格线型应套用模板三线表样式")
        for paragraph in table.iter(f"{{{W_NS}}}p"):
            text = _paragraph_text(paragraph).strip()
            if not text:
                continue
            label = _short_text(text)
            ppr = paragraph.find("w:pPr", NS)
            if _spacing_value(ppr, "line") != "240":
                errors.append(f"表格 {table_idx} {label}: 表内文字应使用模板化单倍行距")
            allowed_sizes = {"21", "24"} if _table_max_columns(table) >= 6 else {"24"}
            if not _first_visible_run_matches_any_size(paragraph, east_asia="宋体", ascii_font="Times New Roman", sizes=allowed_sizes):
                errors.append(f"表格 {table_idx} {label}: 表内文字不应继承学生原稿字号")
            break
        if len(errors) >= 20:
            break
    return errors


def _table_max_columns(table: ET.Element) -> int:
    max_cols = 0
    for row in table.findall("w:tr", NS):
        cols = 0
        for cell in row.findall("w:tc", NS):
            span = cell.find("w:tcPr/w:gridSpan", NS)
            try:
                cols += int(span.get(f"{{{W_NS}}}val") if span is not None else 1)
            except (TypeError, ValueError):
                cols += 1
        max_cols = max(max_cols, cols)
    return max_cols


def _first_visible_run_matches_any_size(
    paragraph: ET.Element,
    *,
    east_asia: str,
    ascii_font: str | None,
    sizes: set[str],
) -> bool:
    return any(_first_visible_run_matches(paragraph, east_asia=east_asia, ascii_font=ascii_font, size=size) for size in sizes)


def _table_border_value(tbl_pr: ET.Element | None, name: str) -> str | None:
    if tbl_pr is None:
        return None
    border = tbl_pr.find(f"w:tblBorders/w:{name}", NS)
    return border.attrib.get(f"{{{W_NS}}}val") if border is not None else None


def _caption_format_errors(root: ET.Element) -> list[str]:
    body = root.find("w:body", NS)
    children = list(body) if body is not None else []
    errors: list[str] = []
    for idx, paragraph in enumerate(children):
        if paragraph.tag != f"{{{W_NS}}}p":
            continue
        text = _paragraph_text(paragraph).strip()
        if not re.match(r"^(图|表)\s*\d+\s*[-－]\s*\d+", text):
            continue
        label = _short_text(text)
        ppr = paragraph.find("w:pPr", NS)
        if _ppr_value(ppr, "jc", "val") != "center":
            errors.append(f"R039/R041 {label}: 图表题注应居中")
        if not _first_visible_run_matches(paragraph, east_asia="宋体", ascii_font="Times New Roman", size="18"):
            errors.append(f"R039/R041 {label}: 图表题注应为宋体小五号")
        if not _first_visible_run_bold(paragraph):
            errors.append(f"R039/R041 {label}: 图表题注应加粗")
        if text.startswith("图") and not _previous_content_has_visual(children, idx):
            errors.append(f"R041 {label}: 图名上方应存在对应图片，不能只有图名")
        if text.startswith("表") and not _next_content_is_table(children, idx):
            errors.append(f"R039 {label}: 表名应位于对应表格正上方")
        if len(errors) >= 30:
            break
    return errors


def _previous_content_has_visual(children: list[ET.Element], idx: int) -> bool:
    for previous in reversed(children[:idx]):
        if previous.tag == f"{{{W_NS}}}tbl":
            return False
        if previous.tag != f"{{{W_NS}}}p":
            continue
        if previous.find(".//w:drawing", NS) is not None or previous.find(".//w:pict", NS) is not None or previous.find(".//w:object", NS) is not None:
            return True
        if _paragraph_text(previous).strip():
            return False
    return False


def _next_content_is_table(children: list[ET.Element], idx: int) -> bool:
    for child in children[idx + 1:]:
        if child.tag == f"{{{W_NS}}}tbl":
            return True
        if child.tag == f"{{{W_NS}}}p" and not _paragraph_text(child).strip() and not (
            child.find(".//w:drawing", NS) is not None or child.find(".//w:pict", NS) is not None or child.find(".//w:object", NS) is not None
        ):
            continue
        return False
    return False


def _reference_format_errors(root: ET.Element) -> list[str]:
    body = root.find("w:body", NS)
    if body is None:
        return []
    errors: list[str] = []
    in_references = False
    expected = 1
    seen_reference = False
    for child in list(body):
        if child.tag != f"{{{W_NS}}}p":
            continue
        text = _paragraph_text(child).strip()
        compact = re.sub(r"\s+", "", text)
        if compact == "参考文献":
            in_references = True
            continue
        if in_references and compact in {"致谢", "附录"}:
            break
        if not in_references:
            continue
        if not text:
            if seen_reference and not _contains_explicit_page_or_section_break(child):
                errors.append("R047 参考文献列表中不应存在空段或多余空行")
            continue
        match = re.match(r"^\[(\d+)\]", text)
        if not match:
            errors.append(f"R047 {_short_text(text)}: 参考文献条目应以 [n] 顺序编号开头")
            if len(errors) >= 30:
                break
            continue
        seen_reference = True
        number = int(match.group(1))
        label = _short_text(text)
        if number != expected:
            errors.append(f"R047 {label}: 参考文献编号应按出现次序连续编号，期望 [{expected}]")
            expected = number
        expected += 1
        ppr = child.find("w:pPr", NS)
        if (
            _has_ppr_child(ppr, "pStyle")
            or _has_ppr_child(ppr, "numPr")
            or _has_ppr_child(ppr, "keepNext")
            or _has_ppr_child(ppr, "pageBreakBefore")
        ):
            errors.append(f"R047 {label}: 参考文献条目不应残留段落样式、列表或分页控制")
        if _has_reference_field_or_inline_style(child):
            errors.append(f"R047 {label}: 参考文献条目不应残留超链接域、域代码或字符样式")
        if _reference_spacing_issue(text):
            errors.append(f"R047 {label}: 参考文献著录标点后应按模板保留必要空格")
        if _spacing_value(ppr, "before") not in {"0", None} or _spacing_value(ppr, "after") not in {"0", None}:
            errors.append(f"R047 {label}: 参考文献段前段后应为 0 磅")
        if _spacing_value(ppr, "line") != "300":
            errors.append(f"R047 {label}: 参考文献应为 1.25 倍行距")
        if _indent_value(ppr, "left") != "480" or (
            _indent_value(ppr, "hanging") != "480" and _indent_value(ppr, "hangingChars") != "200"
        ):
            errors.append(f"R047 {label}: 参考文献应左缩进 2 字符并悬挂缩进 2 字符")
        if not _has_ppr_child(ppr, "keepLines"):
            errors.append(f"R047 {label}: 参考文献长条目应段中不分页，避免跨页孤行")
        if _indent_value(ppr, "firstLine") is not None or _indent_value(ppr, "firstLineChars") is not None:
            errors.append(f"R047 {label}: 参考文献不应保留首行缩进")
        if _ppr_value(ppr, "jc", "val") not in {"left", None}:
            errors.append(f"R047 {label}: 参考文献应左对齐")
        if not _all_visible_runs_match(child, east_asia="宋体", ascii_font="Times New Roman", size="24"):
            errors.append(f"R047 {label}: 参考文献所有文字应为中文小四宋体、英文小四 Times New Roman")
        if len(errors) >= 30:
            break
    return errors


def _reference_spacing_issue(text: str) -> bool:
    body = re.sub(r"^\[\d+\]\s*", "", text.strip(), count=1)
    return bool(
        re.search(r"(?<=[\u4e00-\u9fff\]\)])\.(?=[A-Za-z\u4e00-\u9fff\[])", body)
        or re.search(r"(?<=\])\.(?=\S)", body)
        or re.search(r",(?=\d{4})", body)
        or re.search(r"(?<=\d{4}),(?=[\d(（])", body)
        or re.search(r"\.(?=DOI:)", body, flags=re.IGNORECASE)
    )


def _has_reference_field_or_inline_style(paragraph: ET.Element) -> bool:
    return (
        paragraph.find(".//w:fldChar", NS) is not None
        or paragraph.find(".//w:instrText", NS) is not None
        or paragraph.find(".//w:hyperlink", NS) is not None
        or paragraph.find(".//w:rStyle", NS) is not None
    )


def _contains_explicit_page_or_section_break(paragraph: ET.Element) -> bool:
    return (
        paragraph.find(".//w:br[@w:type='page']", NS) is not None
        or paragraph.find(".//w:sectPr", NS) is not None
    )


def _first_main_body_index(children: list[ET.Element]) -> int | None:
    for idx, child in enumerate(children):
        if child.tag == f"{{{W_NS}}}p" and re.fullmatch(r"1\s+绪论", _paragraph_text(child).strip()):
            return idx
    return None


def _first_main_body_paragraph_pos(paragraphs: list[ET.Element]) -> int | None:
    for idx, paragraph in enumerate(paragraphs):
        if re.fullmatch(r"1\s+绪论", _paragraph_text(paragraph).strip()):
            return idx
    return None


def _is_main_chapter_heading(text: str) -> bool:
    stripped = text.strip()
    if "\t" in stripped or "..." in stripped or "…" in stripped:
        return False
    return bool(re.fullmatch(r"[1-9]\s*[\u4e00-\u9fffA-Za-z][^\n]{0,40}", stripped))


def _is_normal_body_text(text: str) -> bool:
    if not text or len(text) < 6:
        return False
    if _is_main_chapter_heading(text) or _is_sub_heading_text(text):
        return False
    if "\t" in text or "..." in text or "…" in text:
        return False
    if re.match(r"^(图|表)\s*\d+\s*[-－]\s*\d+", text):
        return False
    if re.match(r"^(\[\d+\]|\d+[\.\u3001、])", text):
        return False
    return bool(re.search(r"[\u4e00-\u9fff]", text))


def _is_sub_heading_text(text: str) -> bool:
    stripped = text.strip()
    if "\t" in stripped or "..." in stripped or "…" in stripped:
        return False
    return bool(re.match(r"^[1-9]\.\d+(?:\.\d+)?\s+[\u4e00-\u9fffA-Za-z].{0,60}$", stripped))


def _parse_toc_entry_text(text: str) -> tuple[str, str] | None:
    stripped = text.strip()
    if not stripped or "目录" in re.sub(r"\s+", "", stripped):
        return None
    if not re.match(r"^([1-9](?:\.\d+)*\s*|参考文献|致\s*谢|附录)", stripped):
        return None
    match = re.match(r"^(?P<title>.+?)(?:\.{2,}|…{2,}|\t+|\s{2,})(?P<label>\d+)\s*$", stripped)
    if not match:
        return None
    return match.group("title").strip(" .\t…"), match.group("label")


def _toc_entry_level(title: str) -> int:
    stripped = title.strip()
    if re.match(r"^[1-9]\.\d+\.\d+", stripped):
        return 3
    if re.match(r"^[1-9]\.\d+", stripped):
        return 2
    return 1


def _first_visible_run_matches(
    paragraph: ET.Element,
    *,
    east_asia: str,
    size: str,
    ascii_font: str | None = None,
) -> bool:
    ppr_rpr = paragraph.find("w:pPr/w:rPr", NS)
    for run in paragraph.findall(".//w:r", NS):
        if not _run_text(run).strip():
            continue
        rpr = run.find("w:rPr", NS)
        fonts = _fonts_from_rpr(ppr_rpr) | _fonts_from_rpr(rpr)
        run_size = _size_from_rpr(rpr) or _size_from_rpr(ppr_rpr)
        if run_size != size:
            return False
        if fonts.get("eastAsia") != east_asia:
            return False
        if ascii_font is not None and fonts.get("ascii") != ascii_font and fonts.get("hAnsi") != ascii_font:
            return False
        return True
    fonts = _fonts_from_rpr(ppr_rpr)
    run_size = _size_from_rpr(ppr_rpr)
    if run_size != size or fonts.get("eastAsia") != east_asia:
        return False
    return ascii_font is None or fonts.get("ascii") == ascii_font or fonts.get("hAnsi") == ascii_font


def _all_visible_runs_match(
    paragraph: ET.Element,
    *,
    east_asia: str,
    size: str,
    ascii_font: str | None = None,
) -> bool:
    ppr_rpr = paragraph.find("w:pPr/w:rPr", NS)
    seen = False
    for run in paragraph.findall(".//w:r", NS):
        if not _run_text(run).strip():
            continue
        seen = True
        rpr = run.find("w:rPr", NS)
        fonts = _fonts_from_rpr(ppr_rpr) | _fonts_from_rpr(rpr)
        run_size = _size_from_rpr(rpr) or _size_from_rpr(ppr_rpr)
        if run_size != size:
            return False
        if fonts.get("eastAsia") != east_asia:
            return False
        if ascii_font is not None and fonts.get("ascii") != ascii_font and fonts.get("hAnsi") != ascii_font:
            return False
    if seen:
        return True
    return _first_visible_run_matches(paragraph, east_asia=east_asia, ascii_font=ascii_font, size=size)


def _first_visible_run_bold(paragraph: ET.Element) -> bool:
    ppr_rpr = paragraph.find("w:pPr/w:rPr", NS)
    for run in paragraph.findall(".//w:r", NS):
        if not _run_text(run).strip():
            continue
        return _rpr_bold(run.find("w:rPr", NS)) or _rpr_bold(ppr_rpr)
    return _rpr_bold(ppr_rpr)


def _rpr_bold(rpr: ET.Element | None) -> bool:
    if rpr is None:
        return False
    node = rpr.find("w:b", NS)
    if node is None:
        return False
    return node.attrib.get(f"{{{W_NS}}}val") not in {"0", "false", "False"}


def _run_text(run: ET.Element) -> str:
    return "".join(node.text or "" for node in run.findall("w:t", NS))


def _has_ppr_child(ppr: ET.Element | None, name: str) -> bool:
    return ppr is not None and ppr.find(f"w:{name}", NS) is not None


def _indent_value(ppr: ET.Element | None, attr: str) -> str | None:
    if ppr is None:
        return None
    ind = ppr.find("w:ind", NS)
    return ind.attrib.get(f"{{{W_NS}}}{attr}") if ind is not None else None


def _short_text(text: str) -> str:
    compact = re.sub(r"\s+", "", text)
    return compact[:28] + ("..." if len(compact) > 28 else "")


def _starts_new_page(children: list[ET.Element], idx: int, paragraph: ET.Element) -> bool:
    ppr = paragraph.find("w:pPr", NS)
    page_break = ppr.find("w:pageBreakBefore", NS) if ppr is not None else None
    if page_break is not None and page_break.attrib.get(f"{{{W_NS}}}val") != "0":
        return True
    for previous in reversed(children[:idx]):
        if previous.tag != f"{{{W_NS}}}p":
            continue
        if previous.find(".//w:br[@w:type='page']", NS) is not None:
            return True
        if previous.find(".//w:sectPr", NS) is not None:
            return True
        if _paragraph_text(previous).strip():
            return False
    return False


def _heading_looks_black_xiaoer(paragraph: ET.Element) -> bool:
    for run in paragraph.findall("w:r", NS):
        if not "".join(t.text or "" for t in run.findall("w:t", NS)).strip():
            continue
        rpr = run.find("w:rPr", NS)
        ppr_rpr = paragraph.find("w:pPr/w:rPr", NS)
        fonts = _fonts_from_rpr(ppr_rpr) | _fonts_from_rpr(rpr)
        size = _size_from_rpr(rpr) or _size_from_rpr(ppr_rpr)
        return bool(any(value == "黑体" for value in fonts.values()) and size == "36")
    ppr_rpr = paragraph.find("w:pPr/w:rPr", NS)
    fonts = _fonts_from_rpr(ppr_rpr)
    size = _size_from_rpr(ppr_rpr)
    return bool(any(value == "黑体" for value in fonts.values()) and size == "36")


def _fonts_from_rpr(rpr: ET.Element | None) -> dict[str, str]:
    if rpr is None:
        return {}
    fonts = rpr.find("w:rFonts", NS)
    if fonts is None:
        return {}
    return {key.rsplit("}", 1)[-1]: value for key, value in fonts.attrib.items()}


def _size_from_rpr(rpr: ET.Element | None) -> str | None:
    if rpr is None:
        return None
    size = rpr.find("w:sz", NS)
    return size.attrib.get(f"{{{W_NS}}}val") if size is not None else None


def _ppr_value(ppr: ET.Element | None, child_name: str, attr: str) -> str | None:
    if ppr is None:
        return None
    child = ppr.find(f"w:{child_name}", NS)
    return child.attrib.get(f"{{{W_NS}}}{attr}") if child is not None else None


def _spacing_value(ppr: ET.Element | None, attr: str) -> str | None:
    if ppr is None:
        return None
    spacing = ppr.find("w:spacing", NS)
    return spacing.attrib.get(f"{{{W_NS}}}{attr}") if spacing is not None else None
