
from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Iterable, Iterator
from uuid import uuid4

from app.config import settings
from app.services.model_client import ModelClient
from app.services.pptx_exporter import export_slides_to_pptx
from app.storage.db import (
    get_project,
    list_pages,
    list_project_tasks,
    make_progress,
    replace_pages,
    update_page,
    update_project,
    update_task,
)


OUTLINE_SYSTEM_PROMPT = "You are a helpful assistant for PPT outline generation."
CONTENT_SYSTEM_PROMPT = "You are a helpful assistant for PPT content generation."


def utc_now_iso() -> str:
    return datetime.utcnow().isoformat()


def _normalize_style(style: Any) -> str:
    return "technical" if str(style or "").lower() == "technical" else "management"


def _normalize_title(text: str) -> str:
    value = str(text or "").strip()
    value = re.sub(r"^\s*#{1,6}\s*", "", value)
    value = re.sub(r"^\s*\d+(?:\.\d+){0,3}\s+", "", value)
    value = re.sub(r"^\s*[-*]+\s*", "", value)
    value = re.sub(r"\s+", " ", value)
    return value.strip(" -:\t\n")


def _sanitize_outline_title_text(text: str, fallback: str = "") -> str:
    title = _normalize_title(text)
    if not title:
        return _normalize_title(fallback)
    # Keep third-party flow but avoid polluting cover title with generic "PPT大纲" prefix.
    title = re.sub(r"^(PPT\s*大纲|大纲)\s*[:：\-—]*\s*", "", title, flags=re.IGNORECASE).strip()
    return title or _normalize_title(fallback)


def _sanitize_outline_markdown_for_third_party(markdown_text: str, topic_hint: str = "") -> str:
    raw = str(markdown_text or "")
    lines = raw.splitlines()
    if not lines:
        return raw

    fallback = _normalize_title(topic_hint) or "项目汇报"
    first_heading_idx: int | None = None
    for idx, line in enumerate(lines):
        if line.strip().startswith("# "):
            first_heading_idx = idx
            break

    if first_heading_idx is None:
        return raw

    heading = lines[first_heading_idx].strip()[2:]
    cleaned = _sanitize_outline_title_text(heading, fallback=fallback)
    lines[first_heading_idx] = f"# {cleaned or fallback}"
    return "\n".join(lines)


def _is_cover_title(title: str) -> bool:
    low = _normalize_title(title).lower()
    return bool(low) and any(k in low for k in ("封面", "标题", "cover", "title"))


def _is_toc_title(title: str) -> bool:
    low = _normalize_title(title).lower()
    return bool(low) and any(k in low for k in ("目录", "议程", "agenda", "contents", "toc"))


def _clean_bullets(items: Iterable[str], limit: int = 6) -> list[str]:
    out: list[str] = []
    seen: set[str] = set()
    for raw in items:
        text = _normalize_title(str(raw or ""))
        if not text:
            continue
        key = text.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(text)
        if len(out) >= limit:
            break
    return out


def clean_outline_items(items: Iterable[str]) -> list[str]:
    return _clean_bullets(items, limit=240)


def _default_body_points(title: str) -> list[str]:
    base = _normalize_title(title) or "本页主题"
    return [
        f"{base}的背景与关键问题",
        f"{base}的核心信息与事实依据",
        f"{base}的行动建议与下一步计划",
    ]


def _format_slide_payload(
    title: str,
    bullets: list[str],
    notes: str,
    slide_type: str,
    raw_text: str = "",
) -> dict[str, Any]:
    final_title = _normalize_title(title) or "未命名页面"
    final_bullets = _clean_bullets(bullets, limit=6)
    if slide_type not in {"title", "toc"} and len(final_bullets) < 3:
        final_bullets = _default_body_points(final_title)

    summary = final_bullets[0] if final_bullets else final_title
    return {
        "title": final_title,
        "bullets": final_bullets,
        "detail_points": final_bullets,
        "summary_text": summary,
        "text_blocks": final_bullets[:4],
        "content_format": "summary_plus_three" if len(final_bullets) >= 3 else "summary_plus_two",
        "layout_profile": "summary_plus_three" if len(final_bullets) >= 3 else "summary_plus_two",
        "notes": str(notes or "").strip(),
        "slide_type": slide_type,
        "evidence": final_bullets[:3],
        "chart_data": None,
        "text": raw_text.strip() if raw_text.strip() else "\n".join(final_bullets),
        "generated_image_path": None,
    }


def _merge_outline_material(outline_text: str, material_text: str) -> str:
    sections: list[str] = []
    outline = str(outline_text or "").strip()
    material = str(material_text or "").strip()
    if outline:
        sections.append(f"用户提供的大纲草稿：\n{outline}")
    if material:
        sections.append(f"资料文件内容：\n{material}")
    return "\n\n".join(sections)


def _material_excerpt(material_text: str, title: str, max_chars: int = 2200) -> str:
    material = str(material_text or "").strip()
    if not material:
        return ""

    title_tokens = re.findall(r"[A-Za-z]{2,}|[\u4e00-\u9fff]{2,}", str(title or ""))
    title_tokens = [token.lower() for token in title_tokens][:10]

    parts = re.split(r"[。；;！？!?\n]+", material)
    scored: list[tuple[int, str]] = []
    for part in parts:
        sentence = re.sub(r"\s+", " ", part).strip()
        if len(sentence) < 8:
            continue
        score = 0
        low = sentence.lower()
        for token in title_tokens:
            if token in low:
                score += 1
        scored.append((score, sentence))

    scored.sort(key=lambda x: (x[0], len(x[1])), reverse=True)
    if not scored:
        return material[:max_chars]

    picked: list[str] = []
    used = 0
    for score, sentence in scored:
        if score <= 0 and picked:
            break
        if used + len(sentence) + 1 > max_chars:
            continue
        picked.append(sentence)
        used += len(sentence) + 1
        if len(picked) >= 12:
            break

    return "\n".join(picked) if picked else material[:max_chars]


def _extract_numbered_prefix(line: str) -> tuple[int | None, int | None, int | None, str]:
    text = str(line or "").strip()
    m3 = re.match(r"^(\d+)\.(\d+)\.(\d+)\s+(.+)$", text)
    if m3:
        return int(m3.group(1)), int(m3.group(2)), int(m3.group(3)), _normalize_title(m3.group(4))

    m2 = re.match(r"^(\d+)\.(\d+)\s+(.+)$", text)
    if m2:
        return int(m2.group(1)), int(m2.group(2)), None, _normalize_title(m2.group(3))

    m1 = re.match(r"^(\d+)\.?\s+(.+)$", text)
    if m1:
        return int(m1.group(1)), None, None, _normalize_title(m1.group(2))

    return None, None, None, _normalize_title(text)


def _ensure_chapter(chapters: list[dict[str, Any]], chapter_index: int, chapter_title: str = "") -> dict[str, Any]:
    for chapter in chapters:
        if int(chapter["index"]) == int(chapter_index):
            if chapter_title and (not chapter.get("title") or str(chapter.get("title")).startswith("第")):
                chapter["title"] = chapter_title
            return chapter

    chapter = {
        "index": int(chapter_index),
        "title": chapter_title or f"第{chapter_index}章",
        "sections": [],
    }
    chapters.append(chapter)
    return chapter


def _ensure_section(chapter: dict[str, Any], section_index: int, section_title: str = "") -> dict[str, Any]:
    sections = list(chapter.get("sections") or [])
    for section in sections:
        if int(section["index"]) == int(section_index):
            if section_title and (not section.get("title") or str(section.get("title")).startswith("第")):
                section["title"] = section_title
            chapter["sections"] = sections
            return section

    section = {
        "index": int(section_index),
        "title": section_title or f"第{chapter.get('index', 1)}.{section_index}节",
        "points": [],
    }
    sections.append(section)
    chapter["sections"] = sections
    return section


def _parse_outline_structure(markdown_text: str, topic_hint: str) -> dict[str, Any]:
    title = _normalize_title(topic_hint) or "项目汇报"
    chapters: list[dict[str, Any]] = []
    current_chapter: dict[str, Any] | None = None
    current_section: dict[str, Any] | None = None

    for raw in str(markdown_text or "").splitlines():
        line = str(raw).strip()
        if not line:
            continue
        if line.startswith("```"):
            continue

        if line.startswith("# "):
            parsed_title = _normalize_title(line[2:])
            if parsed_title:
                title = parsed_title
            continue

        if line.startswith("## "):
            candidate = _normalize_title(line[3:])
            cidx, _, _, chapter_title = _extract_numbered_prefix(candidate)
            chapter_index = cidx if cidx is not None else len(chapters) + 1
            current_chapter = _ensure_chapter(chapters, chapter_index, chapter_title)
            current_section = None
            continue

        if line.startswith("### "):
            candidate = _normalize_title(line[4:])
            cidx, sidx, _, section_title = _extract_numbered_prefix(candidate)
            if cidx is None:
                chapter_index = int(current_chapter["index"]) if current_chapter else len(chapters) + 1
            else:
                chapter_index = cidx
            current_chapter = _ensure_chapter(chapters, chapter_index)

            section_index = sidx if sidx is not None else len(list(current_chapter.get("sections") or [])) + 1
            current_section = _ensure_section(current_chapter, section_index, section_title)
            continue

        cidx, sidx, pidx, text = _extract_numbered_prefix(line)
        if cidx is not None and sidx is not None:
            current_chapter = _ensure_chapter(chapters, cidx)
            current_section = _ensure_section(current_chapter, sidx)
            if pidx is not None:
                if text:
                    current_section["points"] = _clean_bullets(
                        [*list(current_section.get("points") or []), text],
                        limit=12,
                    )
            else:
                if text:
                    current_section["title"] = text
            continue

        if line.startswith("- ") or line.startswith("* "):
            text = _normalize_title(line[2:])
            if text and current_section is not None:
                current_section["points"] = _clean_bullets([*list(current_section.get("points") or []), text], limit=12)
            continue

        plain = _normalize_title(line)
        if not plain:
            continue

        if current_chapter is None:
            current_chapter = _ensure_chapter(chapters, 1, "核心内容")
        if current_section is None:
            section_index = len(list(current_chapter.get("sections") or [])) + 1
            current_section = _ensure_section(current_chapter, section_index, plain)
        else:
            current_section["points"] = _clean_bullets([*list(current_section.get("points") or []), plain], limit=12)

    if not chapters:
        fallback_chapter = _ensure_chapter(chapters, 1, "核心内容")
        _ensure_section(fallback_chapter, 1, _normalize_title(topic_hint) or "主题概览")

    for chapter in chapters:
        sections = list(chapter.get("sections") or [])
        if not sections:
            sections = [{"index": 1, "title": chapter.get("title") or "核心内容", "points": []}]
        cleaned_sections: list[dict[str, Any]] = []
        for idx, section in enumerate(sections, start=1):
            sec_title = _normalize_title(str(section.get("title") or "")) or f"第{chapter.get('index', idx)}.{idx}节"
            points = _clean_bullets([str(x) for x in list(section.get("points") or [])], limit=6)
            if not points:
                points = _default_body_points(sec_title)
            cleaned_sections.append({"index": idx, "title": sec_title, "points": points})
        chapter["sections"] = cleaned_sections

    normalized_chapters: list[dict[str, Any]] = []
    for cidx, chapter in enumerate(chapters, start=1):
        chapter_title = _normalize_title(str(chapter.get("title") or "")) or f"第{cidx}章"
        normalized_sections: list[dict[str, Any]] = []
        for sidx, section in enumerate(list(chapter.get("sections") or []), start=1):
            normalized_sections.append(
                {
                    "index": sidx,
                    "title": _normalize_title(str(section.get("title") or "")) or f"第{cidx}.{sidx}节",
                    "points": _clean_bullets([str(x) for x in list(section.get("points") or [])], limit=6),
                }
            )
        normalized_chapters.append({"index": cidx, "title": chapter_title, "sections": normalized_sections})

    return {"title": title, "chapters": normalized_chapters}


def _outline_structure_to_markdown(structure: dict[str, Any]) -> str:
    title = _normalize_title(str(structure.get("title") or "")) or "项目汇报"
    chapters = list(structure.get("chapters") or [])

    lines: list[str] = [f"# {title}"]
    for cidx, chapter in enumerate(chapters, start=1):
        chapter_title = _normalize_title(str(chapter.get("title") or "")) or f"第{cidx}章"
        lines.append(f"## {cidx}. {chapter_title}")

        for sidx, section in enumerate(list(chapter.get("sections") or []), start=1):
            section_title = _normalize_title(str(section.get("title") or "")) or f"第{cidx}.{sidx}节"
            lines.append(f"### {cidx}.{sidx} {section_title}")

            points = _clean_bullets([str(x) for x in list(section.get("points") or [])], limit=6)
            if not points:
                points = _default_body_points(section_title)
            for pidx, point in enumerate(points, start=1):
                lines.append(f"{cidx}.{sidx}.{pidx} {point}")

    return "\n".join(lines)


def _is_thank_title(title: str) -> bool:
    low = _normalize_title(title).lower()
    return bool(low) and any(k in low for k in ("感谢", "致谢", "结尾", "尾页", "结束", "thank", "thanks", "closing"))


def _canonicalize_outline_markdown(raw_markdown: str, topic_hint: str) -> str:
    fallback_title = _normalize_title(topic_hint) or "项目汇报"

    title = ""
    title_locked = False
    lines: list[str] = []

    chapter_no = 0
    section_no = 0
    point_no = 0
    skip_current_chapter = False

    def ensure_chapter(default_title: str = "核心内容") -> None:
        nonlocal chapter_no, section_no, point_no, skip_current_chapter
        if chapter_no <= 0:
            chapter_no = 1
            section_no = 0
            point_no = 0
            skip_current_chapter = False
            lines.append(f"## {chapter_no}. {default_title}")

    def ensure_section(default_title: str = "核心要点") -> None:
        nonlocal section_no, point_no
        if section_no <= 0:
            section_no = 1
            point_no = 0
            lines.append(f"### {chapter_no}.{section_no} {default_title}")

    for raw in str(raw_markdown or "").splitlines():
        line = str(raw).strip()
        if not line or line.startswith("```"):
            continue

        if line.startswith("# "):
            candidate = _normalize_title(line[2:])
            if not candidate:
                continue
            if _is_thank_title(candidate):
                continue
            if not title_locked:
                title = candidate
                title_locked = True
            continue

        if line.startswith("## "):
            candidate = _normalize_title(re.sub(r"^\d+\.?\s*", "", line[3:]))
            if (
                not candidate
                or _is_cover_title(candidate)
                or _is_toc_title(candidate)
                or _is_thank_title(candidate)
            ):
                skip_current_chapter = True
                section_no = 0
                point_no = 0
                continue
            chapter_no += 1
            section_no = 0
            point_no = 0
            skip_current_chapter = False
            lines.append(f"## {chapter_no}. {candidate}")
            continue

        if line.startswith("### "):
            if skip_current_chapter:
                continue
            ensure_chapter()
            candidate = _normalize_title(re.sub(r"^\d+\.\d+\s*", "", line[4:]))
            if (
                not candidate
                or _is_cover_title(candidate)
                or _is_toc_title(candidate)
                or _is_thank_title(candidate)
            ):
                continue
            section_no += 1
            point_no = 0
            lines.append(f"### {chapter_no}.{section_no} {candidate}")
            continue

        if skip_current_chapter:
            continue

        plain_line = line[5:].strip() if line.startswith("#### ") else line
        m_section = re.match(r"^\d+\.\d+\s+(.+)$", plain_line)
        if m_section and not re.match(r"^\d+\.\d+\.\d+\s+.+$", plain_line):
            ensure_chapter()
            candidate = _normalize_title(m_section.group(1))
            if candidate and not (_is_cover_title(candidate) or _is_toc_title(candidate) or _is_thank_title(candidate)):
                section_no += 1
                point_no = 0
                lines.append(f"### {chapter_no}.{section_no} {candidate}")
            continue

        m_point = re.match(r"^\d+\.\d+\.\d+\s+(.+)$", plain_line)
        if m_point:
            candidate = _normalize_title(m_point.group(1))
        elif plain_line.startswith("- ") or plain_line.startswith("* "):
            candidate = _normalize_title(plain_line[2:])
        else:
            candidate = ""

        if not candidate:
            continue

        ensure_chapter()
        ensure_section()
        point_no += 1
        lines.append(f"{chapter_no}.{section_no}.{point_no} {candidate}")

    if not title:
        title = fallback_title

    if not lines:
        lines = [
            "## 1. 核心内容",
            "### 1.1 核心要点",
            "1.1.1 背景与目标",
            "1.1.2 关键分析",
            "1.1.3 行动建议",
        ]

    return "\n".join([f"# {title}", *lines])

def _outline_titles_for_response(structure: dict[str, Any]) -> list[str]:
    titles: list[str] = []
    for chapter in list(structure.get("chapters") or []):
        for section in list(chapter.get("sections") or []):
            sec_title = _normalize_title(str(section.get("title") or ""))
            if sec_title:
                titles.append(sec_title)
    return titles


def _parse_point_detail_map(expanded_markdown: str) -> tuple[dict[str, str], dict[str, list[str]]]:
    point_titles: dict[str, str] = {}
    detail_map: dict[str, list[str]] = {}
    current_key: str | None = None

    for raw in str(expanded_markdown or "").splitlines():
        line = str(raw).strip()
        if not line or line.startswith("```"):
            continue
        if line.startswith("# "):
            continue
        if line.startswith("## "):
            continue
        if line.startswith("### "):
            continue

        normalized = line[5:].strip() if line.startswith("#### ") else line
        m = re.match(r"^(\d+)\.(\d+)\.(\d+)\s+(.+)$", normalized)
        if m:
            key = f"{int(m.group(1))}.{int(m.group(2))}.{int(m.group(3))}"
            point_title = _normalize_title(m.group(4))
            if point_title:
                point_titles[key] = point_title
            current_key = key
            detail_map.setdefault(key, [])
            continue

        if current_key is None:
            continue

        detail = _normalize_title(normalized)
        if detail:
            detail_map.setdefault(current_key, []).append(detail)

    return point_titles, detail_map


def _material_detail_fallback(material_text: str, point_title: str) -> str:
    excerpt = _material_excerpt(material_text, point_title, max_chars=220)
    if not excerpt:
        return ""
    parts = [
        _normalize_title(part)
        for part in re.split(r"[。；;！？!?\n]+", excerpt)
        if _normalize_title(part)
    ]
    if parts:
        return parts[0]
    return _normalize_title(excerpt)


def _canonicalize_content_markdown(
    outline_markdown: str,
    expanded_markdown: str,
    topic: str,
    material_text: str,
) -> str:
    outline_structure = _parse_outline_structure(outline_markdown, topic)
    title = _normalize_title(str(outline_structure.get("title") or topic or "项目汇报")) or "项目汇报"
    point_titles, detail_map = _parse_point_detail_map(expanded_markdown)

    lines: list[str] = [f"# {title}"]
    for cidx, chapter in enumerate(list(outline_structure.get("chapters") or []), start=1):
        chapter_title = _normalize_title(str(chapter.get("title") or "")) or f"第{cidx}章"
        lines.append(f"## {cidx}. {chapter_title}")

        for sidx, section in enumerate(list(chapter.get("sections") or []), start=1):
            section_title = _normalize_title(str(section.get("title") or "")) or f"第{cidx}.{sidx}节"
            lines.append(f"### {cidx}.{sidx} {section_title}")

            base_points = _clean_bullets([str(x) for x in list(section.get("points") or [])], limit=6)
            if not base_points:
                base_points = _default_body_points(section_title)

            for pidx, base_point in enumerate(base_points, start=1):
                key = f"{cidx}.{sidx}.{pidx}"
                point_title = _normalize_title(point_titles.get(key) or base_point) or f"要点{pidx}"
                lines.append(f"{key} {point_title}")

                details = list(detail_map.get(key) or [])
                detail = details[0] if details else ""
                if not detail:
                    detail = _material_detail_fallback(material_text, point_title)
                if not detail:
                    detail = f"围绕“{point_title}”补充背景、依据与行动建议。"
                lines.append(detail)

    return "\n".join(lines)


def _outline_pages_from_structure(structure: dict[str, Any]) -> list[dict[str, Any]]:
    title = _normalize_title(str(structure.get("title") or "")) or "项目汇报"
    chapters = list(structure.get("chapters") or [])

    toc_items = [_normalize_title(str(chapter.get("title") or "")) for chapter in chapters]
    toc_items = [item for item in toc_items if item]

    pages: list[dict[str, Any]] = [
        {"title": "封面", "points": [title, "汇报人", "日期"]},
        {"title": "目录", "points": toc_items[:8] if toc_items else ["内容概览"]},
    ]

    for chapter in chapters:
        chapter_title = _normalize_title(str(chapter.get("title") or "")) or "章节"
        for section in list(chapter.get("sections") or []):
            section_title = _normalize_title(str(section.get("title") or "")) or chapter_title
            section_points = _clean_bullets([str(x) for x in list(section.get("points") or [])], limit=6)
            if not section_points:
                section_points = _default_body_points(section_title)
            pages.append(
                {
                    "title": section_title,
                    "points": section_points,
                    "chapter": chapter_title,
                }
            )

    return pages


def _build_outline_prompt(topic: str, material_text: str, style: str) -> str:
    subject = _normalize_title(topic) or "项目汇报"
    _ = style
    material = str(material_text or "").strip()
    if len(material) > 8000:
        material = material[:8000]

    material_block = ""
    if material:
        material_block = f"\n\n参考资料（仅用于约束事实与补充背景，禁止编造资料中不存在的数字、时间和结论）：\n{material}"

    return f"""
你需要严格按 third-party 生成 PPT 大纲，输出必须是 Markdown，并且仅输出大纲本体。

硬性要求（必须全部满足）：
1. 第一行必须是 `# {subject}`，不要出现“PPT大纲”“大纲”这类前缀词。
2. 必须包含 6 个一级章节，格式为 `## 1. ...` 到 `## 6. ...`。
3. 每个一级章节下包含 3-5 个二级小节，格式为 `### 1.1 ...`。
4. 每个二级小节下包含 3 条要点，格式为 `1.1.1 ...`、`1.1.2 ...`、`1.1.3 ...`。
5. 只允许三层编号（到 `1.1.1`），禁止出现 `1.1.1.1`。
6. 不允许输出封面页、目录页、致谢页等词汇。
7. 只输出 Markdown 文本，不要解释，不要代码块标记。

输出结构示例（必须同构）：
# {subject}
## 1. 章节标题
### 1.1 小节标题
1.1.1 要点一
1.1.2 要点二
1.1.3 要点三

{material_block}
""".strip()


def _build_content_prompt(outline_markdown: str, material_text: str) -> str:
    material = str(material_text or "").strip()
    if len(material) > 8000:
        material = material[:8000]

    material_block = ""
    if material:
        material_block = f"""
参考资料（扩充内容时请优先依据以下资料，禁止编造资料中没有的数字、百分比、时间和结论；若资料不足请审慎表述）：
{material}
"""

    return f"""
你要严格按 third-party 的内容扩充流程处理大纲。

规则：
1. 对所有 `#`、`##`、`###` 行必须逐字原样保留，不能改写、不能增删、不能重排。
2. 对每一行 `1.1.1 ...` 这样的要点，必须在其下一行补充一段 20-50 字说明。
3. 不要新增任何额外标题，不要输出封面、目录、致谢等额外内容。
4. 输出必须为 Markdown 纯文本，不要加解释和代码块。
5. 如果参考资料不足，可做稳健表述，但不能编造资料中没有的具体数字/时间/结论。

输出示例：
### 1.1 小节标题
1.1.1 要点一
对要点一的扩充说明（20-50字）。
1.1.2 要点二
对要点二的扩充说明（20-50字）。
1.1.3 要点三
对要点三的扩充说明（20-50字）。
{material_block}
以下是需要处理的文本:
{outline_markdown}
""".strip()


def _estimate_total_ppt_pages(outline_markdown: str) -> int:
    total = 1
    for raw in str(outline_markdown or "").splitlines():
        line = raw.strip()
        if line.startswith("# "):
            total += 1
        elif line.startswith("## "):
            total += 1
        elif line.startswith("### "):
            total += 1
    total += 1
    return max(total, 1)


def _estimate_current_progress(partial_markdown: str, total_pages: int) -> int:
    chapters: set[str] = set()
    sections: set[str] = set()

    for raw in str(partial_markdown or "").splitlines():
        line = raw.strip()
        if line.startswith("## "):
            chapters.add(_normalize_title(line[3:]).lower())
        elif line.startswith("### "):
            sections.add(_normalize_title(line[4:]).lower())

    estimated = 2 + len(chapters) + len(sections)
    upper = max(1, int(total_pages) - 1)
    if estimated < 1:
        estimated = 1
    if estimated > upper:
        estimated = upper
    return estimated


def _parse_outline_pages_from_rows(rows: list[Any]) -> list[dict[str, Any]]:
    pages: list[dict[str, Any]] = []
    for row in rows:
        try:
            outline = json.loads(str(row["outline_content"] or "{}"))
        except Exception:
            outline = {}
        title = _normalize_title(str(outline.get("title") or ""))
        if not title:
            continue
        points = _clean_bullets([str(x) for x in list(outline.get("points") or [])], limit=8)
        pages.append({"title": title, "points": points})
    return pages


def _outline_markdown_from_pages(topic: str, pages: list[dict[str, Any]]) -> str:
    title = _normalize_title(topic) or "项目汇报"
    content_pages = [
        page
        for idx, page in enumerate(pages)
        if idx > 1 and not _is_cover_title(str(page.get("title") or "")) and not _is_toc_title(str(page.get("title") or ""))
    ]

    chapter = {
        "index": 1,
        "title": "项目汇报",
        "sections": [],
    }
    for idx, page in enumerate(content_pages, start=1):
        section_title = _normalize_title(str(page.get("title") or "")) or f"第1.{idx}节"
        points = _clean_bullets([str(x) for x in list(page.get("points") or [])], limit=6)
        if not points:
            points = _default_body_points(section_title)
        chapter["sections"].append({"index": idx, "title": section_title, "points": points})

    if not chapter["sections"]:
        chapter["sections"].append({"index": 1, "title": "核心内容", "points": _default_body_points("核心内容")})

    structure = {"title": title, "chapters": [chapter]}
    return _outline_structure_to_markdown(structure)


def _outline_bundle_from_raw(raw_markdown: str, topic_hint: str) -> dict[str, Any]:
    canonical_raw = _canonicalize_outline_markdown(raw_markdown, topic_hint)
    structure = _parse_outline_structure(canonical_raw, topic_hint)
    outline_markdown = _outline_structure_to_markdown(structure)
    pages = _outline_pages_from_structure(structure)
    outline_titles = _outline_titles_for_response(structure)
    return {
        "outline_markdown": outline_markdown,
        "pages": pages,
        "outline_titles": outline_titles,
        "title": _normalize_title(str(structure.get("title") or topic_hint or "项目汇报")) or "项目汇报",
    }


def _outline_bundle_from_raw_preserve(raw_markdown: str, topic_hint: str) -> dict[str, Any]:
    text = _sanitize_outline_markdown_for_third_party(raw_markdown, topic_hint).strip()
    title = _sanitize_outline_title_text(topic_hint, fallback="Project Report") or "Project Report"
    pages: list[dict[str, Any]] = []
    current_page: dict[str, Any] | None = None
    current_chapter = ""

    for raw in text.splitlines():
        line = str(raw or "").strip()
        if not line or line.startswith("```"):
            continue

        if line.startswith("# "):
            parsed_title = _sanitize_outline_title_text(line[2:], fallback=title)
            if parsed_title:
                title = parsed_title
            continue

        if line.startswith("## "):
            chapter = _normalize_title(re.sub(r"^\d+\.?\s*", "", line[3:]))
            if chapter:
                current_chapter = chapter
            current_page = None
            continue

        if line.startswith("### "):
            section = _normalize_title(re.sub(r"^\d+\.\d+\s*", "", line[4:]))
            if not section:
                continue
            current_page = {"title": section, "points": [], "chapter": current_chapter}
            pages.append(current_page)
            continue

        if current_page is None:
            continue

        point = ""
        m_point = re.match(r"^\d+\.\d+\.\d+\s+(.+)$", line)
        if m_point:
            point = _normalize_title(m_point.group(1))
        elif line.startswith("- ") or line.startswith("* "):
            point = _normalize_title(line[2:])
        elif line.startswith("#### "):
            point = _normalize_title(line[5:])

        if point:
            current_page["points"] = _clean_bullets([*list(current_page.get("points") or []), point], limit=12)

    if not pages:
        return _outline_bundle_from_raw(text, topic_hint)

    for page in pages:
        points = _clean_bullets([str(x) for x in list(page.get("points") or [])], limit=8)
        if not points:
            points = _default_body_points(str(page.get("title") or "Core Content"))
        page["points"] = points
        if not str(page.get("chapter") or "").strip():
            page.pop("chapter", None)

    return {
        "outline_markdown": text,
        "pages": pages,
        "outline_titles": [str(p.get("title") or "").strip() for p in pages if str(p.get("title") or "").strip()],
        "title": title,
    }
def _parse_expanded_content_sections(expanded_markdown: str) -> dict[str, Any]:
    ppt_title = ""
    chapters: list[dict[str, Any]] = []
    current_chapter: dict[str, Any] | None = None
    current_section: dict[str, Any] | None = None
    current_item: dict[str, Any] | None = None

    def ensure_chapter(title: str) -> dict[str, Any]:
        nonlocal chapters
        chapter = {"title": title, "sections": []}
        chapters.append(chapter)
        return chapter

    def ensure_section(chapter: dict[str, Any], title: str, raw_heading: str) -> dict[str, Any]:
        section = {
            "chapter_title": chapter.get("title") or "章节",
            "title": title,
            "items": [],
            "raw_lines": [raw_heading],
        }
        chapter["sections"].append(section)
        return section

    for raw in str(expanded_markdown or "").splitlines():
        line = str(raw).strip()
        if not line:
            continue
        if line.startswith("```"):
            continue

        if line.startswith("# "):
            ppt_title = _normalize_title(line[2:])
            continue

        if line.startswith("## "):
            chapter_title = _normalize_title(re.sub(r"^\d+\.?\s+", "", line[3:]))
            current_chapter = ensure_chapter(chapter_title or f"第{len(chapters) + 1}章")
            current_section = None
            current_item = None
            continue

        if line.startswith("### "):
            if current_chapter is None:
                current_chapter = ensure_chapter("核心章节")
            section_title = _normalize_title(re.sub(r"^\d+\.\d+\s+", "", line[4:]))
            current_section = ensure_section(current_chapter, section_title or f"第{len(current_chapter['sections']) + 1}节", line)
            current_item = None
            continue

        if current_chapter is None:
            current_chapter = ensure_chapter("核心章节")
        if current_section is None:
            current_section = ensure_section(current_chapter, "核心内容", "### 1.1 核心内容")

        current_section["raw_lines"].append(line)

        no_prefix = line[5:].strip() if line.startswith("#### ") else line
        m = re.match(r"^(\d+\.\d+\.\d+)\s+(.+)$", no_prefix)
        if m:
            point_title = _normalize_title(m.group(2))
            if point_title:
                current_item = {"title": point_title, "detail": ""}
                current_section["items"].append(current_item)
            continue

        detail = _normalize_title(no_prefix)
        if not detail:
            continue

        if current_item is None:
            current_item = {"title": detail, "detail": ""}
            current_section["items"].append(current_item)
        else:
            prev = str(current_item.get("detail") or "")
            current_item["detail"] = f"{prev} {detail}".strip()

    sections_flat: list[dict[str, Any]] = []
    for chapter in chapters:
        for section in list(chapter.get("sections") or []):
            items: list[dict[str, str]] = []
            for item in list(section.get("items") or []):
                point_title = _normalize_title(str(item.get("title") or ""))
                if not point_title:
                    continue
                items.append(
                    {
                        "title": point_title,
                        "detail": _normalize_title(str(item.get("detail") or "")),
                    }
                )
            section["items"] = items
            sections_flat.append(section)

    return {
        "title": ppt_title,
        "chapters": chapters,
        "sections": sections_flat,
    }


def _payloads_from_expanded_markdown(
    topic: str,
    material_text: str,
    outline_pages: list[dict[str, Any]],
    outline_markdown: str,
    expanded_markdown: str,
) -> list[dict[str, Any]]:
    parsed = _parse_expanded_content_sections(expanded_markdown)
    sections = list(parsed.get("sections") or [])

    outline_structure = _parse_outline_structure(outline_markdown, topic)
    chapter_titles = [
        _normalize_title(str(chapter.get("title") or ""))
        for chapter in list(outline_structure.get("chapters") or [])
    ]
    chapter_titles = [title for title in chapter_titles if title]

    used: set[int] = set()
    cursor = 0

    payloads: list[dict[str, Any]] = []
    for idx, page in enumerate(outline_pages):
        page_index = idx + 1
        page_title = _normalize_title(str(page.get("title") or f"第{page_index}页"))
        page_points = _clean_bullets([str(x) for x in list(page.get("points") or [])], limit=6)

        if page_index == 1 or _is_cover_title(page_title):
            payloads.append(
                _format_slide_payload(
                    title=topic or page_title,
                    bullets=[topic or page_title, "汇报人", "日期"],
                    notes="封面页",
                    slide_type="title",
                )
            )
            continue

        if page_index == 2 or _is_toc_title(page_title):
            toc = chapter_titles[:8] if chapter_titles else page_points[:8]
            if not toc:
                toc = ["内容概览"]
            payloads.append(
                _format_slide_payload(
                    title="目录",
                    bullets=toc,
                    notes="目录页",
                    slide_type="toc",
                )
            )
            continue

        selected_idx: int | None = None
        key = page_title.lower()
        for sec_idx, section in enumerate(sections):
            if sec_idx in used:
                continue
            sec_key = _normalize_title(str(section.get("title") or "")).lower()
            if sec_key == key:
                selected_idx = sec_idx
                break

        if selected_idx is None:
            while cursor < len(sections) and cursor in used:
                cursor += 1
            if cursor < len(sections):
                selected_idx = cursor
                cursor += 1

        selected: dict[str, Any] | None = None
        if selected_idx is not None:
            selected = sections[selected_idx]
            used.add(selected_idx)

        items = list((selected or {}).get("items") or [])
        bullets = _clean_bullets([str(item.get("title") or "") for item in items], limit=6)
        if len(bullets) < 3:
            bullets = _clean_bullets(page_points, limit=6)
        if len(bullets) < 3:
            bullets = _default_body_points(page_title)

        details = [str(item.get("detail") or "").strip() for item in items if str(item.get("detail") or "").strip()]
        notes = "\n".join(details[:6]).strip()
        if not notes:
            notes = _material_excerpt(material_text, page_title, max_chars=260) or "资料不足，建议补充原始资料后再生成。"

        raw_lines = [str(x) for x in list((selected or {}).get("raw_lines") or [])]
        raw_text = "\n".join(raw_lines)
        payloads.append(
            _format_slide_payload(
                title=page_title,
                bullets=bullets,
                notes=notes,
                slide_type="summary",
                raw_text=raw_text,
            )
        )

    return payloads


def _build_content_markdown_from_slides(outline_markdown: str, slides: list[dict[str, Any]], topic: str) -> str:
    structure = _parse_outline_structure(outline_markdown, topic)
    title = _normalize_title(str(structure.get("title") or topic or "项目汇报")) or "项目汇报"

    slide_map: dict[str, dict[str, Any]] = {}
    for slide in slides:
        key = _normalize_title(str(slide.get("title") or "")).lower()
        if key and key not in slide_map:
            slide_map[key] = slide

    lines: list[str] = [f"# {title}"]
    for cidx, chapter in enumerate(list(structure.get("chapters") or []), start=1):
        chapter_title = _normalize_title(str(chapter.get("title") or "")) or f"第{cidx}章"
        lines.append(f"## {cidx}. {chapter_title}")

        for sidx, section in enumerate(list(chapter.get("sections") or []), start=1):
            section_title = _normalize_title(str(section.get("title") or "")) or f"第{cidx}.{sidx}节"
            lines.append(f"### {cidx}.{sidx} {section_title}")

            slide = slide_map.get(section_title.lower(), {})
            bullets = _clean_bullets([str(x) for x in list(slide.get("bullets") or section.get("points") or [])], limit=6)
            if not bullets:
                bullets = _default_body_points(section_title)

            note_sentences = [
                sentence.strip()
                for sentence in re.split(r"[。！？!?\n]+", str(slide.get("notes") or ""))
                if sentence.strip()
            ]

            for pidx, bullet in enumerate(bullets[:4], start=1):
                lines.append(f"{cidx}.{sidx}.{pidx} {bullet}")
                detail = note_sentences[pidx - 1] if pidx - 1 < len(note_sentences) else f"围绕“{bullet}”补充背景、依据与行动建议。"
                lines.append(detail)

    return "\n".join(lines)


def _pick_latest_description_markdown(project_id: str) -> tuple[str, str]:
    tasks = list_project_tasks(project_id, limit=40)
    for task in tasks:
        task_type = str(task["task_type"])
        status = str(task["status"])
        if task_type != "GENERATE_DESCRIPTIONS" or status != "COMPLETED":
            continue

        raw_result = str(task["result_json"] or "")
        if not raw_result:
            continue
        try:
            payload = json.loads(raw_result)
        except Exception:
            continue
        outline_markdown = str(payload.get("outline_markdown") or "").strip()
        raw_content_markdown = str(payload.get("raw_content_markdown") or "").strip()
        content_markdown = raw_content_markdown or str(payload.get("content_markdown") or "").strip()
        if outline_markdown and content_markdown:
            return outline_markdown, content_markdown

    return "", ""


def _project_slides_from_rows(rows: list[Any]) -> tuple[list[dict[str, Any]], list[str]]:
    slides: list[dict[str, Any]] = []
    outline_titles: list[str] = []

    for idx, row in enumerate(rows, start=1):
        try:
            outline = json.loads(str(row["outline_content"] or "{}"))
        except Exception:
            outline = {}
        try:
            description = json.loads(str(row["description_content"] or "{}"))
        except Exception:
            description = {}

        title = _normalize_title(str(description.get("title") or outline.get("title") or f"第{idx}页")) or f"第{idx}页"
        bullets = _clean_bullets([str(x) for x in list(description.get("bullets") or outline.get("points") or [])], limit=6)
        notes = str(description.get("notes") or "").strip()

        slide_type = str(description.get("slide_type") or "").strip().lower()
        if not slide_type:
            if idx == 1 or _is_cover_title(title):
                slide_type = "title"
            elif idx == 2 or _is_toc_title(title):
                slide_type = "toc"
            else:
                slide_type = "summary"

        payload = {
            "title": title,
            "bullets": bullets,
            "detail_points": [str(x) for x in list(description.get("detail_points") or bullets)],
            "text_blocks": [str(x) for x in list(description.get("text_blocks") or bullets[:4])],
            "summary_text": str(description.get("summary_text") or (bullets[0] if bullets else title)),
            "content_format": str(description.get("content_format") or "summary_plus_three"),
            "layout_profile": str(description.get("layout_profile") or "summary_plus_three"),
            "notes": notes,
            "slide_type": slide_type,
            "evidence": [str(x) for x in list(description.get("evidence") or bullets[:3])],
            "chart_data": description.get("chart_data"),
            "text": str(description.get("text") or "").strip(),
            "generated_image_path": description.get("generated_image_path"),
        }
        slides.append(payload)
        outline_titles.append(_normalize_title(str(outline.get("title") or title)) or title)

    return slides, outline_titles

class NewBackendFlowEngine:
    def __init__(self, use_mock: bool = False) -> None:
        self.client = ModelClient()
        self.use_mock = bool(use_mock) or (not self.client.enabled())

    def _mock_outline_markdown(self, topic: str) -> str:
        section_count = 6
        chapter_count = 3

        sections_per_chapter = [section_count // chapter_count] * chapter_count
        for idx in range(section_count % chapter_count):
            sections_per_chapter[idx] += 1

        lines = [f"# {_normalize_title(topic) or '项目汇报'}"]
        section_cursor = 0
        for cidx in range(1, chapter_count + 1):
            lines.append(f"## {cidx}. 章节{cidx}")
            for _ in range(sections_per_chapter[cidx - 1]):
                section_cursor += 1
                lines.append(f"### {cidx}.{section_cursor} 页面要点{section_cursor}")
                for pidx in range(1, 4):
                    lines.append(f"{cidx}.{section_cursor}.{pidx} 页面要点{section_cursor}-{pidx}")
        return "\n".join(lines)

    def _mock_expand_markdown(self, outline_markdown: str) -> str:
        lines: list[str] = []
        for raw in str(outline_markdown or "").splitlines():
            line = raw.strip()
            if not line:
                continue
            lines.append(line)
            match = re.match(r"^(\d+\.\d+\.\d+)\s+(.+)$", line)
            if match:
                point = _normalize_title(match.group(2)) or "要点"
                lines.append(f"围绕“{point}”补充背景、依据和可执行建议，形成可直接上屏的讲解内容。")
        return "\n".join(lines)

    def generate_outline_markdown(
        self,
        topic: str,
        material_text: str,
        style: str,
    ) -> str:
        if self.use_mock:
            return self._mock_outline_markdown(topic)

        prompt = _build_outline_prompt(topic, material_text, style)
        chunks: list[str] = []
        for chunk in self.client.chat_text_stream(
            system_prompt=OUTLINE_SYSTEM_PROMPT,
            user_prompt=prompt,
            temperature=0.0,
        ):
            if chunk:
                chunks.append(chunk)
        merged = "".join(chunks).strip()
        if merged:
            return merged
        return self.client.chat_text(
            system_prompt=OUTLINE_SYSTEM_PROMPT,
            user_prompt=prompt,
            temperature=0.0,
        )

    def stream_outline_markdown(
        self,
        topic: str,
        material_text: str,
        style: str,
    ) -> Iterator[str]:
        if self.use_mock:
            yield self._mock_outline_markdown(topic)
            return

        prompt = _build_outline_prompt(topic, material_text, style)
        yield from self.client.chat_text_stream(
            system_prompt=OUTLINE_SYSTEM_PROMPT,
            user_prompt=prompt,
            temperature=0.0,
        )

    def expand_content_markdown(self, outline_markdown: str, material_text: str) -> str:
        if self.use_mock:
            return self._mock_expand_markdown(outline_markdown)

        prompt = _build_content_prompt(outline_markdown, material_text)
        chunks: list[str] = []
        for chunk in self.client.chat_text_stream(
            system_prompt=CONTENT_SYSTEM_PROMPT,
            user_prompt=prompt,
            temperature=0.0,
        ):
            if chunk:
                chunks.append(chunk)
        merged = "".join(chunks).strip()
        if merged:
            return merged
        return self.client.chat_text(
            system_prompt=CONTENT_SYSTEM_PROMPT,
            user_prompt=prompt,
            temperature=0.0,
        )

    def stream_expand_content_markdown(self, outline_markdown: str, material_text: str) -> Iterator[str]:
        if self.use_mock:
            yield self._mock_expand_markdown(outline_markdown)
            return

        prompt = _build_content_prompt(outline_markdown, material_text)
        yield from self.client.chat_text_stream(
            system_prompt=CONTENT_SYSTEM_PROMPT,
            user_prompt=prompt,
            temperature=0.0,
        )


_engine = NewBackendFlowEngine(use_mock=settings.use_mock_llm)


class OutlinePreviewAdapter:
    def generate_outline_bundle(self, title: str, style: str, material: str) -> dict[str, Any]:
        raw_markdown = _engine.generate_outline_markdown(title, material, _normalize_style(style))
        outline_markdown = _sanitize_outline_markdown_for_third_party(raw_markdown, title)
        return _outline_bundle_from_raw_preserve(outline_markdown, title)

    def generate_outline(self, title: str, style: str, material: str) -> list[str]:
        bundle = self.generate_outline_bundle(title, style, material)
        return [str(x) for x in list(bundle.get("outline_titles") or [])]


llm = OutlinePreviewAdapter()


def stream_outline_preview_events(
    title: str,
    style: str,
    material_text: str,
) -> Iterator[dict[str, Any]]:
    full_text = ""
    for chunk in _engine.stream_outline_markdown(title, material_text, _normalize_style(style)):
        if not chunk:
            continue
        full_text += chunk
        yield {"type": "chunk", "text": chunk}

    cleaned = _sanitize_outline_markdown_for_third_party(full_text, title)
    bundle = _outline_bundle_from_raw_preserve(cleaned, title)
    yield {
        "type": "done",
        "outline": list(bundle.get("outline_titles") or []),
        "outline_markdown": str(bundle.get("outline_markdown") or ""),
    }


def _outline_bundle_for_project(
    project_row: dict[str, Any] | Any,
    requested_outline: list[str] | None = None,
    requested_outline_markdown: str | None = None,
) -> dict[str, Any]:
    project_title = str(project_row["title"] or "项目汇报")

    markdown_candidate = str(requested_outline_markdown or "").strip()
    if markdown_candidate:
        cleaned = _sanitize_outline_markdown_for_third_party(markdown_candidate, project_title)
        return _outline_bundle_from_raw_preserve(cleaned, project_title)
    if requested_outline:
        pages: list[dict[str, Any]] = []
        for item in requested_outline:
            section = _normalize_title(str(item or ""))
            if not section:
                continue
            pages.append({"title": section, "points": _default_body_points(section)})
        if pages:
            markdown_lines = [f"# {project_title}", "## 1. Core Content"]
            for idx, page in enumerate(pages, start=1):
                section_title = str(page.get("title") or f"Slide {idx}")
                markdown_lines.append(f"### 1.{idx} {section_title}")
                for pidx, point in enumerate(list(page.get("points") or []), start=1):
                    markdown_lines.append(f"1.{idx}.{pidx} {point}")
            return {
                "outline_markdown": "\n".join(markdown_lines),
                "pages": pages,
                "outline_titles": [str(x.get("title") or "") for x in pages],
                "title": project_title,
            }
    existing_outline_text = str(project_row["outline_text"] or "").strip()
    if existing_outline_text and ("### " in existing_outline_text or "## " in existing_outline_text):
        cleaned = _sanitize_outline_markdown_for_third_party(existing_outline_text, project_title)
        return _outline_bundle_from_raw_preserve(cleaned, project_title)

    material_text = _merge_outline_material(
        str(project_row["outline_text"] or ""),
        str(project_row["material_text"] or ""),
    )
    style = _normalize_style(project_row.get("style") if hasattr(project_row, "get") else project_row["style"])
    raw_markdown = _engine.generate_outline_markdown(project_title, material_text, style)
    cleaned = _sanitize_outline_markdown_for_third_party(raw_markdown, project_title)
    return _outline_bundle_from_raw_preserve(cleaned, project_title)


def get_outline_for_project(
    project_row: dict[str, Any] | Any,
    requested_outline: list[str] | None = None,
    requested_outline_markdown: str | None = None,
) -> tuple[list[dict[str, Any]], str]:
    bundle = _outline_bundle_for_project(project_row, requested_outline, requested_outline_markdown)
    return list(bundle.get("pages") or []), str(bundle.get("outline_markdown") or "")


def rebuild_project_pages(project_id: str, outline_pages: list[dict[str, Any]], outline_markdown: str = "") -> None:
    now = utc_now_iso()

    rows: list[dict[str, Any]] = []
    for idx, page in enumerate(outline_pages):
        title = _normalize_title(str(page.get("title") or f"第{idx + 1}页")) or f"第{idx + 1}页"
        points = _clean_bullets([str(x) for x in list(page.get("points") or [])], limit=8)

        payload: dict[str, Any] = {
            "title": title,
            "points": points,
        }
        chapter = _normalize_title(str(page.get("chapter") or ""))
        if chapter:
            payload["chapter"] = chapter

        rows.append(
            {
                "page_id": str(uuid4()),
                "project_id": project_id,
                "order_index": idx,
                "outline_content": json.dumps(payload, ensure_ascii=False),
                "description_content": None,
                "status": "DRAFT",
                "created_at": now,
                "updated_at": now,
            }
        )

    replace_pages(project_id, rows)

    project = get_project(project_id)
    topic = str(project["title"] or "项目汇报") if project else "项目汇报"
    markdown = str(outline_markdown or "").strip() or _outline_markdown_from_pages(topic, outline_pages)
    markdown = _sanitize_outline_markdown_for_third_party(markdown, topic)

    update_project(
        project_id,
        {
            "outline_text": markdown,
            "status": "OUTLINE_GENERATED",
            "updated_at": now,
        },
    )

def _write_payloads_to_pages(
    rows: list[Any],
    payloads: list[dict[str, Any]],
    task_id: str | None = None,
) -> tuple[int, int]:
    completed = 0
    failed = 0

    for idx, row in enumerate(rows):
        page_id = str(row["page_id"])
        try:
            payload = payloads[idx] if idx < len(payloads) else _format_slide_payload(
                title=f"第{idx + 1}页",
                bullets=_default_body_points(f"第{idx + 1}页"),
                notes="资料不足，建议补充原始资料后再生成。",
                slide_type="summary",
            )
            update_page(
                page_id,
                {
                    "description_content": json.dumps(payload, ensure_ascii=False),
                    "status": "DESCRIPTION_GENERATED",
                    "updated_at": utc_now_iso(),
                },
            )
            completed += 1
        except Exception as exc:
            failed += 1
            update_page(
                page_id,
                {
                    "status": "FAILED",
                    "updated_at": utc_now_iso(),
                },
            )
            if task_id:
                update_task(task_id, {"error_message": str(exc)})

    return completed, failed


def _generate_descriptions_core(
    project: Any,
    rows: list[Any],
    on_chunk: Callable[[str], None] | None = None,
    on_progress: Callable[[int, int], None] | None = None,
) -> dict[str, Any]:
    topic = str(project["title"] or "项目汇报")
    material_text = str(project["material_text"] or "")
    outline_pages = _parse_outline_pages_from_rows(rows)
    outline_markdown = str(project["outline_text"] or "").strip()
    if not outline_markdown:
        outline_markdown = _outline_markdown_from_pages(topic, outline_pages)
    outline_markdown = _sanitize_outline_markdown_for_third_party(outline_markdown, topic)

    total_pages = _estimate_total_ppt_pages(outline_markdown)
    if on_progress:
        on_progress(1, total_pages)

    chunks: list[str] = []
    current = 1

    try:
        for chunk in _engine.stream_expand_content_markdown(outline_markdown, material_text):
            if not chunk:
                continue
            chunks.append(chunk)
            if on_chunk:
                on_chunk(chunk)
            merged = "".join(chunks)
            estimated = _estimate_current_progress(merged, total_pages)
            if estimated > current:
                current = estimated
                if on_progress:
                    on_progress(current, total_pages)
    except Exception:
        chunks = []

    raw_content_markdown = "".join(chunks).strip()
    if not raw_content_markdown:
        raw_content_markdown = _engine.expand_content_markdown(outline_markdown, material_text).strip()
        if on_chunk and raw_content_markdown:
            on_chunk(raw_content_markdown)

    if not raw_content_markdown:
        raise RuntimeError("content expansion returned empty markdown")

    content_markdown = str(raw_content_markdown).strip()

    payloads = _payloads_from_expanded_markdown(
        topic=topic,
        material_text=material_text,
        outline_pages=outline_pages,
        outline_markdown=outline_markdown,
        expanded_markdown=content_markdown,
    )

    if on_progress:
        on_progress(total_pages, total_pages)

    return {
        "topic": topic,
        "outline_markdown": outline_markdown,
        "raw_content_markdown": raw_content_markdown,
        "content_markdown": content_markdown,
        "total_ppt_pages": total_pages,
        "payloads": payloads,
        "outline_pages": outline_pages,
    }


def generate_descriptions_task(task_id: str, project_id: str) -> None:
    project = get_project(project_id)
    if not project:
        update_task(
            task_id,
            {
                "status": "FAILED",
                "error_message": "project not found",
                "completed_at": utc_now_iso(),
            },
        )
        return

    pages = list_pages(project_id)
    if not pages:
        update_task(
            task_id,
            {
                "status": "FAILED",
                "error_message": "no pages to generate",
                "completed_at": utc_now_iso(),
            },
        )
        return

    total_estimated = _estimate_total_ppt_pages(str(project["outline_text"] or ""))
    update_task(
        task_id,
        {
            "status": "PROCESSING",
            "progress_json": make_progress(total_estimated, 0, 0, "expanding_outline_with_third_party_prompt"),
        },
    )

    latest_current = 0

    def on_progress(current: int, total: int) -> None:
        nonlocal latest_current
        if current <= latest_current:
            return
        latest_current = current
        update_task(
            task_id,
            {
                "progress_json": make_progress(total, current, 0, "streaming_expanded_content"),
            },
        )

    try:
        result = _generate_descriptions_core(project, pages, on_chunk=None, on_progress=on_progress)
    except Exception as exc:
        update_task(
            task_id,
            {
                "status": "FAILED",
                "error_message": str(exc),
                "completed_at": utc_now_iso(),
            },
        )
        return

    completed, failed = _write_payloads_to_pages(pages, list(result.get("payloads") or []), task_id=task_id)

    final_status = "COMPLETED" if failed == 0 else "FAILED"
    if failed == 0:
        update_project(
            project_id,
            {
                "status": "DESCRIPTIONS_GENERATED",
                "updated_at": utc_now_iso(),
            },
        )

    update_task(
        task_id,
        {
            "status": final_status,
            "progress_json": make_progress(
                int(result.get("total_ppt_pages") or max(completed + failed, 1)),
                int(result.get("total_ppt_pages") or completed),
                failed,
                "descriptions_done",
            ),
            "error_message": None if failed == 0 else f"{failed} pages failed",
            "result_json": json.dumps(
                {
                    "outline_markdown": str(result.get("outline_markdown") or ""),
                    "raw_content_markdown": str(result.get("raw_content_markdown") or ""),
                    "content_markdown": str(result.get("content_markdown") or ""),
                    "total_ppt_pages": int(result.get("total_ppt_pages") or 0),
                    "generated_pages": completed,
                },
                ensure_ascii=False,
            ),
            "completed_at": utc_now_iso(),
        },
    )


def stream_generate_descriptions_events(project_id: str) -> Iterator[dict[str, Any]]:
    project = get_project(project_id)
    if not project:
        raise RuntimeError("project not found")

    pages = list_pages(project_id)
    if not pages:
        raise RuntimeError("no pages to generate")

    partial_text = ""
    topic = str(project["title"] or "项目汇报")
    material_text = str(project["material_text"] or "")
    outline_pages = _parse_outline_pages_from_rows(pages)
    outline_markdown = str(project["outline_text"] or "").strip() or _outline_markdown_from_pages(topic, outline_pages)
    outline_markdown = _sanitize_outline_markdown_for_third_party(outline_markdown, topic)
    total_pages = _estimate_total_ppt_pages(outline_markdown)

    yield {
        "type": "meta",
        "total": total_pages,
        "current": 1,
        "outline_markdown": outline_markdown,
    }

    current = 1
    try:
        for chunk in _engine.stream_expand_content_markdown(outline_markdown, material_text):
            if not chunk:
                continue
            partial_text += chunk
            yield {"type": "chunk", "text": chunk}
            estimated = _estimate_current_progress(partial_text, total_pages)
            if estimated > current:
                current = estimated
                yield {"type": "progress", "current": current, "total": total_pages}
    except Exception:
        partial_text = ""

    raw_content_markdown = partial_text.strip()
    if not raw_content_markdown:
        raw_content_markdown = _engine.expand_content_markdown(outline_markdown, material_text)
        if raw_content_markdown:
            yield {"type": "chunk", "text": raw_content_markdown}

    content_markdown = str(raw_content_markdown or "").strip()

    payloads = _payloads_from_expanded_markdown(
        topic=topic,
        material_text=material_text,
        outline_pages=outline_pages,
        outline_markdown=outline_markdown,
        expanded_markdown=content_markdown,
    )

    completed, failed = _write_payloads_to_pages(pages, payloads)

    if failed == 0:
        update_project(
            project_id,
            {
                "status": "DESCRIPTIONS_GENERATED",
                "updated_at": utc_now_iso(),
            },
        )

    yield {
        "type": "done",
        "total": total_pages,
        "current": total_pages,
        "generated_pages": completed,
        "failed_pages": failed,
        "raw_content_markdown": raw_content_markdown,
        "content_markdown": content_markdown,
    }


def generate_ppt_task(task_id: str, project_id: str) -> None:
    project = get_project(project_id)
    if not project:
        update_task(
            task_id,
            {
                "status": "FAILED",
                "error_message": "project not found",
                "completed_at": utc_now_iso(),
            },
        )
        return

    pages = list_pages(project_id)
    if not pages:
        update_task(
            task_id,
            {
                "status": "FAILED",
                "error_message": "no pages to export",
                "completed_at": utc_now_iso(),
            },
        )
        return

    total = len(pages)
    update_task(
        task_id,
        {
            "status": "PROCESSING",
            "progress_json": make_progress(total, total, 0, "exporting_pptx"),
        },
    )

    try:
        slides, outline_titles = _project_slides_from_rows(pages)
        if not slides:
            raise RuntimeError("no slide content available for export")

        outline_markdown, content_markdown = _pick_latest_description_markdown(project_id)
        if not outline_markdown:
            outline_markdown = str(project["outline_text"] or "").strip()
        if not outline_markdown:
            raise RuntimeError("outline markdown missing; please generate outline first")
        if not content_markdown:
            raise RuntimeError("description markdown missing; please generate descriptions first")

        ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        filename = f"{project_id}_{ts}_{uuid4().hex[:8]}.pptx"
        out_path = Path(settings.export_dir) / filename

        exported = export_slides_to_pptx(
            slides=slides,
            out_path=out_path,
            template_id=str(project["template_id"] or "a2p_2"),
            topic=str(project["title"] or "项目汇报"),
            outline=outline_titles,
            subtitle="",
            toc_items=[],
            style=_normalize_style(project["style"]),
            theme_seed=f"{project_id}|{project['title']}|third-party",
            outline_markdown=outline_markdown,
            content_markdown=content_markdown,
        )

        pptx_url = f"/exports/{exported}"
        update_project(
            project_id,
            {
                "status": "COMPLETED",
                "pptx_url": pptx_url,
                "updated_at": utc_now_iso(),
            },
        )
        update_task(
            task_id,
            {
                "status": "COMPLETED",
                "progress_json": make_progress(total, total, 0, "done"),
                "result_json": json.dumps(
                    {
                        "pptx_url": pptx_url,
                        "outline_markdown": outline_markdown,
                        "content_markdown": content_markdown,
                    },
                    ensure_ascii=False,
                ),
                "completed_at": utc_now_iso(),
            },
        )
    except Exception as exc:
        update_task(
            task_id,
            {
                "status": "FAILED",
                "error_message": str(exc),
                "completed_at": utc_now_iso(),
            },
        )




