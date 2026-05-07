from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class ContentImprovementPlan:
    target: Path
    actions: list[str] = field(default_factory=list)
    markdown: str = ""


def build_content_plan(audit_result: dict[str, Any], target: Path) -> ContentImprovementPlan:
    content = audit_result["content_review"]
    actions: list[str] = []
    issue_codes = {issue.code for issue in content.issues}

    if "thin-test-chapter" in issue_codes or "weak-test-method" in issue_codes:
        actions.append("补强第 4 章系统测试：增加测试环境、测试用例表、测试步骤、预期结果、实测结果、结果分析。")
    if "web-reference-heavy" in issue_codes:
        actions.append("优化参考文献：减少普通网页来源，补充近五年期刊、标准、芯片手册、协议规范或学位论文。")
    if "low-citation-coverage" in issue_codes:
        actions.append("复核正文引用：确保正文中引用编号与参考文献列表一一对应。")
    if "missing-engineering-topic" in issue_codes:
        actions.append("补足工程设计链路：明确需求分析、总体设计、模块实现、测试验证和总结之间的对应关系。")
    if "language-typo" in issue_codes or "repeated-word" in issue_codes or "punctuation" in issue_codes:
        actions.append("做形式校对：修正错别字、重复词、异常标点和术语不一致问题。")
    if "thin-chapter" in issue_codes or "chapter-imbalance" in issue_codes:
        actions.append("平衡各章篇幅：对偏薄章节补充设计依据、关键参数、实现过程、问题分析和小结。")
    if "figure-heavy-chapter" in issue_codes or "figure-without-explanation" in issue_codes:
        actions.append("补充图表说明：每张关键图前说明目的，图后解释模块含义、参数、数据结果和对设计结论的支撑。")

    if not actions:
        actions.append("内容规则未发现阻断项，进入人工终审或视觉模型终审。")

    markdown = _markdown(audit_result, target, actions)
    return ContentImprovementPlan(target=target, actions=actions, markdown=markdown)


def write_content_plan(audit_result: dict[str, Any], target: Path, out_path: Path) -> ContentImprovementPlan:
    plan = build_content_plan(audit_result, target)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(plan.markdown, encoding="utf-8")
    return plan


def _markdown(audit_result: dict[str, Any], target: Path, actions: list[str]) -> str:
    content = audit_result["content_review"]
    lines = [
        "# Content Improvement Plan",
        "",
        f"- Target: `{target}`",
        f"- Content score: {content.score}/100",
        f"- Chinese chars: {content.chinese_chars}",
        f"- References: {content.reference_count}",
        f"- Foreign references: {content.foreign_reference_count}",
        f"- Web references: {content.web_reference_count}",
        "",
        "## Chapter Coverage",
        "",
    ]
    for chapter, count in content.chapter_char_counts.items():
        lines.append(f"- {chapter}: {count} Chinese chars")
    lines.extend(["", "## Required Actions", ""])
    for action in actions:
        lines.append(f"- {action}")
    lines.extend([
        "",
        "## Suggested Test Chapter Structure",
        "",
        "第 4 章应至少覆盖以下小节或等价内容：",
        "",
        "- 测试环境：硬件清单、软件版本、网络/云平台环境、传感器型号和阈值设置。",
        "- 测试用例：用表格列出编号、测试目标、输入条件、操作步骤、预期结果、实测结果和结论。",
        "- 功能测试：采集、显示、RFID 识别、云端上传、报警联动等逐项验证。",
        "- 异常测试：温湿度越界、气体浓度异常、网络中断、RFID 未授权等场景。",
        "- 结果分析：说明系统是否满足需求、误差来源、稳定性和后续改进。",
    ])
    return "\n".join(lines) + "\n"
