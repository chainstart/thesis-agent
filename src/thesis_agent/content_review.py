from __future__ import annotations

import re
from dataclasses import dataclass, field

from .config import AgentConfig


@dataclass(frozen=True)
class ReviewIssue:
    severity: str
    code: str
    message: str


@dataclass(frozen=True)
class ContentReview:
    score: int
    chinese_chars: int
    reference_count: int
    foreign_reference_count: int
    web_reference_count: int
    keyword_count: int | None
    chapter_char_counts: dict[str, int]
    issues: list[ReviewIssue] = field(default_factory=list)


def review_content(text: str, config: AgentConfig) -> ContentReview:
    quality = config.content_quality
    issues: list[ReviewIssue] = []
    normalized = re.sub(r"\s+", "", text)
    normalized_lower = normalized.lower()
    chinese_chars = len(re.findall(r"[\u4e00-\u9fff]", text))
    references = _extract_references(text)
    foreign_refs = [ref for ref in references if re.search(r"[A-Za-z]{3,}", ref)]
    web_refs = [ref for ref in references if "EB/OL" in ref or "http" in ref.lower()]
    keyword_count = _keyword_count(text)
    chapters = _chapter_texts(text)
    chapter_char_counts = {
        chapter: len(re.findall(r"[\u4e00-\u9fff]", body))
        for chapter, body in chapters.items()
    }

    for section in config.expected_sections:
        if not _section_present(section, text, normalized, normalized_lower):
            issues.append(ReviewIssue("error", "missing-section", f"缺少或未识别必要结构：{section}"))

    if chinese_chars < int(quality.get("min_body_chinese_chars", 7000)):
        issues.append(ReviewIssue("error", "body-too-short", f"正文中文字符量偏少：约 {chinese_chars} 字。"))
    elif chinese_chars < int(quality.get("recommended_body_chinese_chars", 10000)):
        issues.append(ReviewIssue("warning", "body-below-recommended", f"中文字符量约 {chinese_chars} 字，低于建议值。"))

    ref_min = int(quality.get("min_reference_count", 10))
    if len(references) < ref_min:
        issues.append(ReviewIssue("error", "few-references", f"参考文献数量 {len(references)} 篇，少于 {ref_min} 篇。"))

    foreign_min = int(quality.get("min_foreign_reference_count", 1))
    if len(foreign_refs) < foreign_min:
        issues.append(ReviewIssue("warning", "few-foreign-references", "未识别到足够外文参考文献。"))
    if references and len(web_refs) / len(references) > 0.60:
        issues.append(ReviewIssue("warning", "web-reference-heavy", "网页类参考文献比例偏高，建议补充期刊、标准、手册或专著类资料。"))

    cited_numbers = _body_citation_numbers(text)
    if references and len(cited_numbers) < min(len(references), max(5, int(len(references) * 0.40))):
        issues.append(ReviewIssue("warning", "low-citation-coverage", f"正文中可识别引用编号 {len(cited_numbers)} 个，参考文献与正文引用对应关系需要复核。"))

    if keyword_count is None:
        issues.append(ReviewIssue("warning", "missing-keywords", "未识别到关键词行。"))
    elif not int(quality.get("min_keywords", 3)) <= keyword_count <= int(quality.get("max_keywords", 5)):
        issues.append(ReviewIssue("warning", "keyword-count", f"关键词数量疑似为 {keyword_count}，建议 3-5 个。"))

    for topic in quality.get("required_engineering_topics", []):
        if topic not in normalized:
            issues.append(ReviewIssue("warning", "missing-engineering-topic", f"未识别到工程论文常见内容：{topic}"))
    test_chapter = _find_test_chapter(chapters)
    if test_chapter:
        test_chars = len(re.findall(r"[\u4e00-\u9fff]", test_chapter))
        if test_chars < 800:
            issues.append(ReviewIssue("warning", "thin-test-chapter", f"测试章节中文内容约 {test_chars} 字，支撑本科工科设计偏薄。"))
        if not re.search(r"测试环境|测试用例|测试结果|结果分析|测试分析|功能测试|性能测试", test_chapter):
            issues.append(ReviewIssue("warning", "weak-test-method", "测试章节未明显覆盖测试环境、测试用例、测试结果或结果分析。"))

    for marker in quality.get("forbidden_markers", []):
        if marker in text:
            issues.append(ReviewIssue("error", "forbidden-marker", f"发现不应保留的标记：{marker}"))

    issues.extend(_formal_language_issues(text))
    issues.extend(_chapter_quality_issues(chapters, chapter_char_counts))
    issues.extend(_figure_explanation_issues(text))

    score = 100
    for issue in issues:
        score -= 12 if issue.severity == "error" else 5
    score = max(score, 0)
    return ContentReview(
        score=score,
        chinese_chars=chinese_chars,
        reference_count=len(references),
        foreign_reference_count=len(foreign_refs),
        web_reference_count=len(web_refs),
        keyword_count=keyword_count,
        chapter_char_counts=chapter_char_counts,
        issues=issues,
    )


def _formal_language_issues(text: str) -> list[ReviewIssue]:
    issues: list[ReviewIssue] = []
    replacements = {
        "别的越来越": "变得越来越",
        "接受到": "接收到",
        "雄安锡": "相应",
        "KeiluVision": "Keil uVision",
        "水质水质": "水质",
    }
    for wrong, right in replacements.items():
        if wrong in text:
            issues.append(ReviewIssue("warning", "language-typo", f"疑似错别字或术语错误：`{wrong}`，建议改为 `{right}`。"))
    repeated: list[str] = []
    for match in re.finditer(r"([\u4e00-\u9fff]{2,4})\1", text):
        token = match.group(1)
        if token not in repeated and token not in {"研究研究", "系统系统"}:
            repeated.append(token)
        if len(repeated) >= 5:
            break
    for token in repeated:
        issues.append(ReviewIssue("warning", "repeated-word", f"疑似重复词：`{token}{token}`，需要人工复核。"))
    if re.search(r"[，,。；;：:]\s*[，,。；;：:]", text):
        issues.append(ReviewIssue("warning", "punctuation", "发现连续标点或异常标点组合，建议做形式校对。"))
    return issues[:10]


def _chapter_quality_issues(chapters: dict[str, str], chapter_char_counts: dict[str, int]) -> list[ReviewIssue]:
    issues: list[ReviewIssue] = []
    body_chapters = {chapter: count for chapter, count in chapter_char_counts.items() if re.match(r"^[1-9]\s*", chapter)}
    for chapter, count in body_chapters.items():
        if count < 500:
            issues.append(ReviewIssue("warning", "thin-chapter", f"{chapter} 中文内容约 {count} 字，章节内容偏薄。"))
    substantial = [count for count in body_chapters.values() if count >= 300]
    if len(substantial) >= 3 and max(substantial) / max(min(substantial), 1) >= 4.5:
        issues.append(ReviewIssue("warning", "chapter-imbalance", "各章篇幅差异过大，建议检查是否存在图多文少或技术说明不足的章节。"))
    for chapter, body in chapters.items():
        figure_count = len(re.findall(r"图\s*\d+\s*[-－]\s*\d+", body))
        chars = len(re.findall(r"[\u4e00-\u9fff]", body))
        if figure_count >= 3 and chars < figure_count * 260:
            issues.append(ReviewIssue("warning", "figure-heavy-chapter", f"{chapter} 图较多但文字说明偏少，建议补充图前引导、图后分析和参数说明。"))
    return issues[:12]


def _figure_explanation_issues(text: str) -> list[ReviewIssue]:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    issues: list[ReviewIssue] = []
    previous_caption_idx: int | None = None
    previous_caption: str | None = None
    for idx, line in enumerate(lines):
        if not re.match(r"^图\s*\d+\s*[-－]\s*\d+", line):
            continue
        if previous_caption_idx is not None:
            between = "".join(lines[previous_caption_idx + 1: idx])
            chinese_between = len(re.findall(r"[\u4e00-\u9fff]", between))
            if chinese_between < 80:
                issues.append(
                    ReviewIssue(
                        "warning",
                        "figure-without-explanation",
                        f"{previous_caption or '上一幅图'} 与 {line} 之间文字说明不足，建议补充图中模块、参数或测试结果分析。",
                    )
                )
                if len(issues) >= 8:
                    break
        previous_caption_idx = idx
        previous_caption = line
    return issues


def _extract_references(text: str) -> list[str]:
    lines = [line.strip() for line in text.splitlines()]
    refs: list[str] = []
    heading_indices = [
        idx for idx, line in enumerate(lines)
        if re.sub(r"\s+", "", line) == "参考文献"
    ]
    if not heading_indices:
        heading_indices = [
            idx for idx, line in enumerate(lines)
            if re.match(r"^\s*参考文献\s*$", line)
        ]
    start = heading_indices[-1] + 1 if heading_indices else 0
    for line in lines[start:]:
        if not line:
            continue
        if re.match(r"^(致谢|附录|附 件|Appendix)\b", line):
            break
        if re.match(r"^(\[\d+\]|\d+[\.\u3001、])", line) or _looks_like_reference_line(line):
            refs.append(line)
    if refs:
        return refs
    return [line for line in lines if re.match(r"^\[\d+\]", line) or _looks_like_reference_line(line)]


def _looks_like_reference_line(line: str) -> bool:
    compact = line.strip()
    if len(compact) < 12:
        return False
    if re.match(r"^[1-9]\s+[\u4e00-\u9fffA-Za-z]", compact):
        return False
    return bool(
        re.search(r"\[(J|D|S|M|C|N|P|R|EB/OL|OL)\]", compact, re.IGNORECASE)
        or re.search(r"DOI[:：]?\s*10\.", compact, re.IGNORECASE)
        or re.search(r"https?://", compact, re.IGNORECASE)
        or re.search(r"(19|20)\d{2}\s*[,，(（]", compact)
    )


def _keyword_count(text: str) -> int | None:
    for line in text.splitlines():
        compact = line.strip()
        if re.match(r"^(关键词|关键字|Key words|Keywords)\s*[:：]", compact, re.IGNORECASE):
            payload = re.split(r"[:：]", compact, maxsplit=1)[1]
            parts = [part.strip() for part in re.split(r"[；;，,、\s]+", payload) if part.strip()]
            return len(parts)
    return None


def _section_present(section: str, text: str, normalized: str, normalized_lower: str) -> bool:
    needle = re.sub(r"\s+", "", section)
    if not needle:
        return True
    if needle.lower() in normalized_lower:
        return True
    if needle == "目錄" or needle == "目录":
        return _has_toc_entries(text)
    if needle.lower() == "abstract":
        return "abstract" in normalized_lower
    if "总结" in needle or "结论" in needle or "展望" in needle:
        return bool(re.search(r"(总结|结论|展望)", normalized))
    return False


def _has_toc_entries(text: str) -> bool:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    leader_lines = sum(1 for line in lines if re.search(r"\.{4,}\s*\d+\s*$", line))
    main_entries = sum(1 for line in lines if re.match(r"^[1-9]\s+.+\.{4,}\s*\d+\s*$", line))
    return leader_lines >= 5 or main_entries >= 3


def _chapter_texts(text: str) -> dict[str, str]:
    chapters: dict[str, list[str]] = {}
    current: str | None = None
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if _is_toc_entry_line(line):
            continue
        match = re.match(r"^([1-9])\s*([\u4e00-\u9fffA-Za-z].{0,40})$", line)
        if match:
            current = f"{match.group(1)} {match.group(2).strip()}"
            chapters.setdefault(current, [])
            continue
        if current:
            chapters[current].append(line)
    return {chapter: "\n".join(lines) for chapter, lines in chapters.items()}


def _find_chapter(chapters: dict[str, str], number: str) -> str | None:
    for chapter, body in chapters.items():
        if chapter.startswith(f"{number} "):
            return body
    return None


def _find_test_chapter(chapters: dict[str, str]) -> str | None:
    for chapter, body in chapters.items():
        if re.search(r"(测试|调试|验证|实验)", chapter):
            return body
    return _find_chapter(chapters, "4")


def _is_toc_entry_line(line: str) -> bool:
    return bool(
        re.search(r"\.{4,}\s*\d+\s*$", line)
        or re.search(r"\t+\s*\d+\s*$", line)
    )


def _body_citation_numbers(text: str) -> set[str]:
    lines = text.splitlines()
    ref_heading_indices = [
        idx for idx, line in enumerate(lines)
        if re.sub(r"\s+", "", line) == "参考文献"
    ]
    end = ref_heading_indices[-1] if ref_heading_indices else len(lines)
    body = "\n".join(lines[:end])
    numbers = set(re.findall(r"\[(\d+)\]", body))
    numbers.update(re.findall(r"\[(\d+)(?=[,，、])", body))
    return numbers
