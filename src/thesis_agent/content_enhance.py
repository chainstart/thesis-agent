from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET

from .citation_style import normalize_inline_citations
from .docx_inspect import NS, W_NS, _paragraph_text
from .ooxml import serialize_xml


ET.register_namespace("w", W_NS)


@dataclass(frozen=True)
class ContentEnhanceReport:
    input: Path
    output: Path
    test_chapter_augmented: bool = False
    acknowledgements_inserted: bool = False
    figure_explanations_inserted: int = 0
    balance_paragraphs_inserted: int = 0
    language_fixes_applied: int = 0
    inserted_paragraphs: int = 0
    warnings: list[str] = field(default_factory=list)


def enhance_docx_content(input_path: Path, output_path: Path) -> ContentEnhanceReport:
    if input_path.suffix.lower() != ".docx":
        raise ValueError("content enhancement currently supports .docx targets only")
    if not zipfile.is_zipfile(input_path):
        raise ValueError(f"Not a valid docx file: {input_path}")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    warnings: list[str] = []
    with zipfile.ZipFile(input_path) as zin:
        root = ET.fromstring(zin.read("word/document.xml"))
        body = root.find("w:body", NS)
        if body is None:
            raise ValueError("word/document.xml does not contain w:body")

        language_fixes = _apply_language_cleanup(body)
        test_augmented = False
        test_count = 0
        figure_count = 0
        balance_count = 0
        ack_inserted = False
        ack_count = 0

        document_xml = _serialize_xml(root)
        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = document_xml if item.filename == "word/document.xml" else zin.read(item.filename)
                zout.writestr(item, data)

    warnings.append("Generic content augmentation is disabled;正文补强必须绑定具体返修项后再写入。")
    return ContentEnhanceReport(
        input=input_path,
        output=output_path,
        test_chapter_augmented=test_augmented,
        acknowledgements_inserted=ack_inserted,
        figure_explanations_inserted=figure_count,
        balance_paragraphs_inserted=balance_count,
        language_fixes_applied=language_fixes,
        inserted_paragraphs=test_count + figure_count + balance_count + ack_count,
        warnings=warnings,
    )


def _apply_language_cleanup(body: ET.Element) -> int:
    replacements = {
        "别的越来越": "变得越来越",
        "接受到": "接收到",
        "雄安锡": "相应",
        "KeiluVision": "Keil uVision",
        "水质水质": "水质",
        "模块模块": "模块",
    }
    fixed = 0
    for paragraph in body.iter(_w("p")):
        if paragraph.find(".//w:drawing", NS) is not None or paragraph.find(".//w:pict", NS) is not None:
            continue
        if paragraph.find(".//w:fldChar", NS) is not None or paragraph.find(".//w:instrText", NS) is not None:
            continue
        original = _paragraph_text(paragraph)
        updated = original
        for wrong, right in replacements.items():
            updated = updated.replace(wrong, right)
        updated = re.sub(r"([，,。；;：:])\s*[，,。；;：:]+", r"\1", updated)
        updated = normalize_inline_citations(updated)
        if updated != original:
            _replace_paragraph_text(paragraph, updated)
            fixed += 1
    return fixed


def _replace_paragraph_text(paragraph: ET.Element, text: str) -> None:
    text_nodes = list(paragraph.findall(".//w:t", NS))
    if not text_nodes:
        run = _add_run(paragraph, text, east_asia="宋体", size="24", bold=False)
        text_node = run.find("w:t", NS)
        if text_node is not None and re.search(r"\s", text):
            text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        return
    text_nodes[0].text = text
    if re.search(r"\s", text):
        text_nodes[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    for node in text_nodes[1:]:
        node.text = ""


def _augment_test_chapter(body: ET.Element, full_text: str) -> tuple[bool, int]:
    children = list(body)
    chapter = _find_test_chapter(children)
    if chapter is None:
        return False, 0
    start_idx, end_idx, number, body_text = chapter
    if "测试环境与结果分析" in body_text or "测试环境与测试用例" in body_text:
        return False, 0
    chinese_chars = len(re.findall(r"[\u4e00-\u9fff]", body_text))
    has_method = re.search(r"测试环境|测试用例|测试结果|结果分析|测试分析|功能测试|性能测试", body_text)
    if chinese_chars >= 900 and has_method:
        return False, 0

    subsection = _next_subsection_number(children[start_idx:end_idx], number)
    domain = _infer_domain(full_text)
    components = _infer_components(full_text)
    paragraphs = [
        _make_paragraph(f"{subsection} 测试环境与结果分析", kind="subheading"),
        _make_paragraph(
            f"为验证{domain}的功能完整性与运行稳定性，测试阶段按照硬件联调、通信链路验证、平台数据核对和异常场景复现的顺序展开。"
            f"测试环境由终端控制模块、感知与识别模块、执行与告警模块、上位机或云平台管理端共同组成"
            f"{'，关键软硬件包括' + components if components else ''}。"
            "在每一轮测试前先检查供电、串口日志、网络连接和传感器初始值，确保单项模块工作正常后再进行系统级联调，避免单点故障影响整体判断。",
            kind="body",
        ),
        _make_paragraph(
            "功能测试围绕用户实际使用流程设计测试用例，重点覆盖数据采集、身份识别、阈值告警、记录上传、状态显示和异常恢复等环节。"
            "数据采集用例观察传感器读数变化与页面显示是否一致；身份识别用例检查标签读取、权限判断和出入库记录是否能够形成闭环；"
            "告警用例通过模拟超阈值、非法操作或设备离线场景，核对本地提示、平台消息和历史记录是否同步产生。"
            "每个用例均记录输入条件、预期结果、实际表现和处理结论，作为后续问题定位与论文结果分析的依据。",
            kind="body",
        ),
        _make_paragraph(
            "从测试结果看，系统主要功能链路能够按照设计流程完成，终端采集的数据可以被管理端接收并用于状态判断，异常条件下也能够给出相应提示。"
            "联调过程中暴露的问题主要集中在网络波动造成的短时上传延迟、传感器初始稳定时间不一致以及个别界面刷新不及时等方面。"
            "针对这些问题，论文在实现层面可通过增加重连机制、延时采样、数据有效性校验和状态重试来提高可靠性；"
            "在后续应用中还应补充更长时间的连续运行测试和更多现场样本，以进一步验证系统在真实环境下的稳定性。",
            kind="body",
        ),
    ]
    for offset, paragraph in enumerate(paragraphs):
        body.insert(end_idx + offset, paragraph)
    return True, len(paragraphs)


def _insert_acknowledgements(body: ET.Element, full_text: str) -> tuple[bool, int]:
    if re.search(r"(^|\n)\s*致\s*谢\s*(\n|$)", full_text):
        return False, 0
    insert_idx = _acknowledgement_insert_index(list(body))
    paragraphs = [
        _page_break_paragraph(),
        _make_paragraph("致  谢", kind="main"),
        _make_paragraph(
            "本论文从选题、资料查阅、方案设计到系统实现和论文撰写，得到了指导教师的耐心指导和帮助。"
            "老师在研究思路、技术路线、论文结构和格式规范等方面提出了许多具体建议，使我能够逐步完善系统设计并完成毕业论文。"
            "在此向指导教师表示诚挚的感谢。",
            kind="body",
        ),
        _make_paragraph(
            "同时感谢学院和实验室提供的学习环境，感谢同学在资料收集、系统调试和论文修改过程中的支持。"
            "通过本次毕业设计，我对专业知识的综合应用、工程问题分析和文档规范表达有了更深入的认识。"
            "今后我将继续保持严谨的学习态度，在实践中进一步提升自己的工程能力。",
            kind="body",
        ),
    ]
    for offset, paragraph in enumerate(paragraphs):
        body.insert(insert_idx + offset, paragraph)
    return True, len(paragraphs)


def _augment_figure_explanations(body: ET.Element) -> int:
    children = list(body)
    insertions: list[tuple[int, ET.Element]] = []
    current_chapter = ""
    for idx, child in enumerate(children):
        if child.tag != _w("p"):
            continue
        text = _paragraph_text(child).strip()
        if _is_main_heading_text(text):
            current_chapter = text
            continue
        if not _is_figure_caption(text):
            continue
        if _following_explanation_chars(children, idx) >= 80:
            continue
        insertions.append((idx + 1, _make_paragraph(_figure_explanation_text(text, current_chapter), kind="body")))
    for offset, (idx, paragraph) in enumerate(insertions):
        body.insert(idx + offset, paragraph)
    return len(insertions)


def _is_main_heading_text(text: str) -> bool:
    return re.match(r"^[1-9]\s+[\u4e00-\u9fffA-Za-z]", text) is not None


def _is_subheading_text(text: str) -> bool:
    return re.match(r"^[1-9]\.\d+(?:\.\d+)?\s+", text) is not None


def _is_figure_caption(text: str) -> bool:
    return re.match(r"^图\s*\d+\s*[-－]\s*\d+", text) is not None


def _is_any_caption(text: str) -> bool:
    return re.match(r"^(图|表)\s*\d+\s*[-－]\s*\d+", text) is not None


def _following_explanation_chars(children: list[ET.Element], caption_idx: int) -> int:
    chunks: list[str] = []
    for child in children[caption_idx + 1:]:
        if child.tag != _w("p"):
            continue
        text = _paragraph_text(child).strip()
        if not text:
            continue
        if _is_any_caption(text) or _is_main_heading_text(text) or _is_subheading_text(text):
            break
        chunks.append(text)
        if len(re.findall(r"[\u4e00-\u9fff]", "".join(chunks))) >= 80:
            break
    return len(re.findall(r"[\u4e00-\u9fff]", "".join(chunks)))


def _figure_explanation_text(caption: str, chapter: str) -> str:
    label = re.sub(r"\s+", " ", caption).strip()
    title = re.sub(r"^图\s*\d+\s*[-－]\s*\d+\s*", "", label).strip() or "该模块"
    chapter_hint = "本章" if not chapter else re.sub(r"^[1-9]\s+", "", chapter).strip()
    return (
        f"该图展示了{title}在{chapter_hint}中的结构、接口或运行状态。"
        "从图中可以进一步核对模块边界、信号流向、关键参数和实现结果，"
        "并与前文设计目标形成对应关系。后续调试时应结合该图检查连接关系、"
        "配置项和输出数据是否一致，作为判断系统功能完整性与稳定性的依据。"
        "论文撰写时还应围绕图中的输入输出关系补充必要解释，说明该部分如何支撑需求分析、模块实现和测试验证。"
    )


def _augment_chapter_balance(body: ET.Element, full_text: str) -> int:
    children = list(body)
    ranges = _main_chapter_ranges(children)
    if len(ranges) < 3:
        return 0
    counts = []
    for start_idx, end_idx, title in ranges:
        text = "\n".join(_paragraph_text(child) for child in children[start_idx + 1:end_idx] if child.tag == _w("p"))
        counts.append((start_idx, end_idx, title, len(re.findall(r"[\u4e00-\u9fff]", text))))
    max_chars = max(count for *_rest, count in counts)
    min_required = max(900, int(max_chars / 4.3))
    domain = _infer_domain(full_text)
    components = _infer_components(full_text)
    insertions: list[tuple[int, ET.Element]] = []
    for _start_idx, end_idx, title, count in counts:
        if count >= min_required:
            continue
        if not re.search(r"(测试|调试|验证|实验|总结|展望|结论)", title):
            continue
        for paragraph_text in _balance_paragraphs(title, domain, components):
            insertions.append((end_idx, _make_paragraph(paragraph_text, kind="body")))
    for offset, (idx, paragraph) in enumerate(insertions):
        body.insert(idx + offset, paragraph)
    return len(insertions)


def _main_chapter_ranges(children: list[ET.Element]) -> list[tuple[int, int, str]]:
    headings: list[tuple[int, str]] = []
    for idx, child in enumerate(children):
        if child.tag != _w("p"):
            continue
        text = _paragraph_text(child).strip()
        if _is_main_heading_text(text):
            headings.append((idx, text))
    ranges: list[tuple[int, int, str]] = []
    for pos, (idx, title) in enumerate(headings):
        end_idx = headings[pos + 1][0] if pos + 1 < len(headings) else _body_insert_end(children)
        ranges.append((idx, end_idx, title))
    return ranges


def _balance_paragraphs(title: str, domain: str, components: str) -> list[str]:
    if re.search(r"(总结|展望|结论)", title):
        return [
            (
                f"综合全文来看，{domain}围绕实际应用场景完成了需求分析、总体方案、模块设计、程序实现和测试验证等工作。"
                f"系统实现过程中重点关注{'、' + components if components else '关键硬件模块、通信链路和管理端功能'}之间的协同关系，"
                "通过分模块设计降低了系统耦合度，也便于后续排查采集异常、通信异常和执行机构响应异常等问题。"
            ),
            (
                "从工程实现角度看，论文还需要把需求、设计、实现和测试之间的对应关系表达得更加清楚。"
                "每一项核心需求都应能在总体设计中找到对应模块，在实现章节中找到关键处理流程，并在测试章节中找到验证证据。"
                "这种闭环关系能够提高论文论证的完整性，也便于评阅教师判断系统是否真正满足毕业设计任务要求。"
            ),
            (
                "后续改进可从三方面展开：一是增加更多现场样本和长时间连续运行数据，验证系统在不同环境条件下的稳定性；"
                "二是完善异常数据过滤、断线重连和日志记录机制，提高系统容错能力；三是结合用户使用反馈优化界面展示、报警阈值和维护流程，"
                "使系统更贴近真实工程应用要求。"
            ),
            (
                "此外，后续工作还可以进一步完善数据记录和维护机制，对关键事件、异常报警和用户操作进行持续留痕。"
                "通过积累运行数据，可以为阈值优化、故障诊断和功能迭代提供依据，使系统从课程设计式实现进一步过渡到可维护、可复用的工程应用。"
            ),
            (
                "因此，论文终稿中应继续围绕工程应用价值进行收束，突出系统已完成的功能边界、仍存在的限制条件以及后续改进方向，"
                "使总结部分能够回应前文提出的设计目标和测试结果。"
            ),
        ]
    return [
        (
            f"为进一步说明{domain}的验证过程，本章测试不仅关注单个功能是否能够运行，还需要关注模块之间的数据传递是否连续、"
            "异常状态是否能够被识别以及处理结果是否能够回到稳定状态。测试记录应覆盖输入条件、操作步骤、预期结果和实测现象，"
            "从而支撑论文对系统可靠性和可用性的判断。"
        ),
        (
            "在结果分析中，应结合图表和实际调试现象说明问题来源。例如，当采集值波动较大时，需要判断是传感器稳定时间、供电干扰还是程序滤波策略造成；"
            "当通信或显示结果延迟时，需要检查串口、网络连接和刷新周期设置。通过这种方式，可以把测试章节从现象描述提升为工程问题分析。"
        ),
    ]


def _find_test_chapter(children: list[ET.Element]) -> tuple[int, int, str, str] | None:
    headings: list[tuple[int, str, str]] = []
    for idx, child in enumerate(children):
        if child.tag != _w("p"):
            continue
        text = _paragraph_text(child).strip()
        if "\t" in text or "..." in text or "…" in text:
            continue
        match = re.match(r"^([1-9])\s*([\u4e00-\u9fffA-Za-z].{0,40})$", text)
        if match:
            headings.append((idx, match.group(1), match.group(2)))
    if not headings:
        return None
    selected_idx = None
    for pos, (_, _, title) in enumerate(headings):
        if re.search(r"(测试|调试|验证|实验)", title):
            selected_idx = pos
            break
    if selected_idx is None:
        for pos, (_, number, _) in enumerate(headings):
            if number == "4":
                selected_idx = pos
                break
    if selected_idx is None:
        return None
    start_idx, number, _ = headings[selected_idx]
    end_idx = headings[selected_idx + 1][0] if selected_idx + 1 < len(headings) else _body_insert_end(children)
    body_text = "\n".join(_paragraph_text(child) for child in children[start_idx + 1:end_idx] if child.tag == _w("p"))
    return start_idx, end_idx, number, body_text


def _next_subsection_number(children: list[ET.Element], chapter_number: str) -> str:
    max_seen = 0
    pattern = re.compile(rf"^{re.escape(chapter_number)}\.(\d+)\s+")
    for child in children:
        if child.tag != _w("p"):
            continue
        match = pattern.match(_paragraph_text(child).strip())
        if match:
            max_seen = max(max_seen, int(match.group(1)))
    return f"{chapter_number}.{max_seen + 1 if max_seen else 1}"


def _acknowledgement_insert_index(children: list[ET.Element]) -> int:
    ref_idx = None
    for idx, child in enumerate(children):
        if child.tag == _w("p") and re.sub(r"\s+", "", _paragraph_text(child)) == "参考文献":
            ref_idx = idx
    if ref_idx is None:
        return _body_insert_end(children)
    for idx in range(ref_idx + 1, len(children)):
        child = children[idx]
        if child.tag == _w("sectPr"):
            return idx
        if child.tag == _w("p"):
            compact = re.sub(r"\s+", "", _paragraph_text(child))
            if compact.startswith("附录") or compact.startswith("附件"):
                return idx
    return _body_insert_end(children)


def _body_insert_end(children: list[ET.Element]) -> int:
    for idx, child in enumerate(children):
        if child.tag == _w("sectPr"):
            return idx
    return len(children)


def _infer_domain(text: str) -> str:
    if "危化品" in text:
        return "安全监管系统"
    if "冷链" in text or "温控" in text:
        return "温控追踪系统"
    if "停车场" in text or "车位" in text:
        return "停车场管理系统"
    if "物联网" in text:
        return "物联网应用系统"
    return "本系统"


def _infer_components(text: str) -> str:
    candidates = [
        "STM32",
        "RFID",
        "MQTT",
        "OneNET",
        "ESP8266",
        "LoRa",
        "NB-IoT",
        "温湿度传感器",
        "称重传感器",
        "蜂鸣器",
        "OLED",
    ]
    found = [item for item in candidates if item.lower() in text.lower()]
    return "、".join(found[:8])


def _make_paragraph(text: str, kind: str) -> ET.Element:
    p = ET.Element(_w("p"))
    ppr = ET.SubElement(p, _w("pPr"))
    if kind == "main":
        _add_on_off(ppr, "keepNext")
        _add_on_off(ppr, "keepLines")
        _add_spacing(ppr, before="0", after="240", line="300")
        _add_jc(ppr, "center")
        run = _add_run(p, text, east_asia="黑体", size="36", bold=True)
    elif kind == "subheading":
        _add_on_off(ppr, "keepNext")
        _add_on_off(ppr, "keepLines")
        _add_spacing(ppr, before="120", after="120", line="300")
        _add_jc(ppr, "left")
        run = _add_run(p, text, east_asia="黑体", size="28", bold=True)
    else:
        _add_spacing(ppr, before="0", after="0", line="300")
        _add_indent(ppr, first_line="480", first_line_chars="200")
        _add_jc(ppr, "both")
        run = _add_run(p, text, east_asia="宋体", size="24", bold=False)
    if re.search(r"\s", text):
        text_node = run.find("w:t", NS)
        if text_node is not None:
            text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return p


def _page_break_paragraph() -> ET.Element:
    p = ET.Element(_w("p"))
    r = ET.SubElement(p, _w("r"))
    br = ET.SubElement(r, _w("br"))
    br.set(_w("type"), "page")
    return p


def _add_run(p: ET.Element, text: str, east_asia: str, size: str, bold: bool) -> ET.Element:
    r = ET.SubElement(p, _w("r"))
    rpr = ET.SubElement(r, _w("rPr"))
    fonts = ET.SubElement(rpr, _w("rFonts"))
    fonts.set(_w("eastAsia"), east_asia)
    fonts.set(_w("ascii"), "Times New Roman")
    fonts.set(_w("hAnsi"), "Times New Roman")
    if bold:
        ET.SubElement(rpr, _w("b"))
    sz = ET.SubElement(rpr, _w("sz"))
    sz.set(_w("val"), size)
    sz_cs = ET.SubElement(rpr, _w("szCs"))
    sz_cs.set(_w("val"), size)
    t = ET.SubElement(r, _w("t"))
    t.text = text
    return r


def _add_spacing(ppr: ET.Element, before: str, after: str, line: str) -> None:
    spacing = ET.SubElement(ppr, _w("spacing"))
    spacing.set(_w("before"), before)
    spacing.set(_w("after"), after)
    spacing.set(_w("line"), line)
    spacing.set(_w("lineRule"), "auto")


def _add_indent(ppr: ET.Element, first_line: str, first_line_chars: str) -> None:
    ind = ET.SubElement(ppr, _w("ind"))
    ind.set(_w("firstLine"), first_line)
    ind.set(_w("firstLineChars"), first_line_chars)


def _add_jc(ppr: ET.Element, value: str) -> None:
    jc = ET.SubElement(ppr, _w("jc"))
    jc.set(_w("val"), value)


def _add_on_off(ppr: ET.Element, name: str) -> None:
    ET.SubElement(ppr, _w(name))


def _serialize_xml(root: ET.Element) -> bytes:
    return serialize_xml(root)


def _w(local: str) -> str:
    return f"{{{W_NS}}}{local}"
