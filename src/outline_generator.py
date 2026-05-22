"""Rule-based PPT outline generation.

This module intentionally avoids external AI dependencies. It can later be
replaced or augmented with an LLM-backed outline generator while keeping the
same return shape for the renderer.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List


@dataclass
class Slide:
    """A simple intermediate representation for one slide."""

    slide_type: str
    title: str
    subtitle: str
    bullets: List[str]
    notes: str = ""


STYLE_ALIASES: Dict[str, str] = {
    "business": "商务风",
    "tech": "科技风",
    "education": "教育风",
    "marketing": "营销风",
    "minimal": "极简风",
}


def normalize_style(style: str) -> str:
    """Map free-form style text to one of the supported style names."""

    value = (style or "").strip()
    lowered = value.lower()

    if any(word in value for word in ["商务", "商业", "汇报", "高级"]):
        return "商务风"
    if any(word in value for word in ["科技", "AI", "数字", "未来"]) or "tech" in lowered:
        return "科技风"
    if any(word in value for word in ["教育", "课程", "课件", "教学"]):
        return "教育风"
    if any(word in value for word in ["营销", "增长", "品牌", "传播"]):
        return "营销风"
    if any(word in value for word in ["极简", "简约", "干净"]) or "minimal" in lowered:
        return "极简风"

    return STYLE_ALIASES.get(lowered, value or "商务风")


def generate_outline(
    topic: str,
    pages: int = 8,
    style: str = "商务风",
    audience: str = "通用受众",
    purpose: str = "演示汇报",
    requirements: str = "",
) -> List[Slide]:
    """Generate a complete slide outline.

    The outline always includes a cover, agenda, content pages, summary, and
    closing slide. The total number of slides equals ``pages``.
    """

    safe_pages = max(4, min(int(pages or 8), 30))
    safe_topic = topic.strip() or "未命名主题"
    safe_style = normalize_style(style)
    safe_audience = audience.strip() or "通用受众"
    safe_purpose = purpose.strip() or "演示汇报"
    safe_requirements = requirements.strip()

    content_count = max(1, safe_pages - 4)
    content_titles = _content_titles(safe_purpose, safe_requirements, content_count)

    slides: List[Slide] = [
        Slide(
            slide_type="cover",
            title=safe_topic,
            subtitle=f"{safe_purpose} | 面向：{safe_audience}",
            bullets=[f"风格：{safe_style}", "自动生成的可编辑 PPTX"],
            notes="封面页应清晰呈现主题、用途和受众。",
        ),
        Slide(
            slide_type="agenda",
            title="目录",
            subtitle="本次演示的核心结构",
            bullets=_agenda_items(content_titles),
            notes="目录页用于帮助听众建立全局预期。",
        ),
    ]

    for index, section_title in enumerate(content_titles, start=1):
        slides.append(
            Slide(
                slide_type="content",
                title=section_title,
                subtitle=_section_subtitle(section_title, safe_topic, safe_audience),
                bullets=_section_bullets(section_title, safe_topic, safe_audience, safe_requirements, index),
                notes=f"内容页 {index}：围绕 {section_title} 展开。",
            )
        )

    slides.extend(
        [
            Slide(
                slide_type="summary",
                title="总结与建议",
                subtitle="把演示内容收束为可执行结论",
                bullets=[
                    f"围绕「{safe_topic}」形成统一认知",
                    "明确优先级、执行节奏和判断标准",
                    "将下一步行动拆成可跟踪任务",
                ],
                notes="总结页强调结论、共识和行动。",
            ),
            Slide(
                slide_type="closing",
                title="谢谢",
                subtitle="期待进一步交流",
                bullets=[
                    "确认关键问题",
                    "同步后续安排",
                    "沉淀行动清单",
                ],
                notes="结束页保持简洁，方便进入问答或讨论。",
            ),
        ]
    )

    return slides[:safe_pages]


def _content_titles(purpose: str, requirements: str, count: int) -> List[str]:
    base = [
        "背景与目标",
        "核心洞察",
        "方案结构",
        "执行路径",
        "资源与节奏",
        "风险与应对",
        "衡量指标",
        "下一步计划",
    ]

    if any(word in purpose for word in ["课程", "课件", "教育", "培训"]):
        base = ["学习目标", "知识导入", "核心概念", "课堂活动", "练习任务", "学习反馈", "课后巩固", "下一课预告"]
    elif any(word in purpose for word in ["营销", "品牌", "传播", "增长"]):
        base = ["市场背景", "用户洞察", "传播定位", "内容策略", "渠道打法", "转化路径", "指标复盘", "行动计划"]
    elif any(word in purpose for word in ["路演", "融资", "投资", "BP"]):
        base = ["市场机会", "用户痛点", "产品方案", "核心优势", "商业模式", "增长路径", "团队与计划", "融资用途"]

    requirement_topics = _extract_requirement_topics(requirements)
    merged = []
    for item in requirement_topics + base:
        if item and item not in merged:
            merged.append(item)

    while len(merged) < count:
        merged.append(f"重点模块 {len(merged) + 1}")

    return merged[:count]


def _extract_requirement_topics(requirements: str) -> List[str]:
    if not requirements:
        return []

    separators = ["、", "，", ",", ";", "；", "\n"]
    normalized = requirements
    for separator in separators:
        normalized = normalized.replace(separator, "|")

    topics = []
    for part in normalized.split("|"):
        cleaned = part.strip()
        cleaned = cleaned.replace("展示", "").replace("包含", "").replace("说明", "").strip()
        if 2 <= len(cleaned) <= 12:
            topics.append(cleaned)
    return topics[:6]


def _agenda_items(content_titles: List[str]) -> List[str]:
    if len(content_titles) <= 5:
        return content_titles
    return [content_titles[0], content_titles[1], "方案与执行", "指标与风险", "总结与行动"]


def _section_subtitle(title: str, topic: str, audience: str) -> str:
    return f"围绕「{topic}」向{audience}说明：{title}"


def _section_bullets(title: str, topic: str, audience: str, requirements: str, index: int) -> List[str]:
    requirement_hint = requirements or "结合目标、场景和落地动作进行说明"
    templates = [
        f"明确本页与「{topic}」的关系，避免信息发散",
        f"从{audience}最关心的问题出发组织表达",
        f"提炼 2-3 个关键判断，支撑「{title}」",
        f"落到可执行动作：{requirement_hint}",
    ]

    if index % 3 == 0:
        templates[2] = "用对比、案例或数据增强可信度"
    if index % 4 == 0:
        templates[3] = "明确负责人、时间点和交付物"

    return templates
