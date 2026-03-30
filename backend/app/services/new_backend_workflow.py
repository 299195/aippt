from __future__ import annotations

import json
import re
from datetime import datetime
from typing import Any, Iterable

from app.config import settings
from app.services.model_client import ModelClient
from app.services.project_workflow import (
    generate_ppt_task as legacy_generate_ppt_task,
    rebuild_project_pages as legacy_rebuild_project_pages,
)
from app.storage.db import (
    get_project,
    list_pages,
    make_progress,
    update_page,
    update_project,
    update_task,
)


def utc_now_iso() -> str:
    return datetime.utcnow().isoformat()


def _normalize_style(style: Any) -> str:
    return "technical" if str(style or "").lower() == "technical" else "management"


def _normalize_title(text: str) -> str:
    raw = str(text or "").strip()
    raw = re.sub(r"^\s*#{1,6}\s*", "", raw)
    raw = re.sub(r"^\s*\d+(?:\.\d+){0,3}\s*", "", raw)
    raw = re.sub(r"^\s*[-*]+\s*", "", raw)
    raw = re.sub(r"\s+", " ", raw).strip(" -:\t\n")
    return raw


def _is_cover_title(title: str) -> bool:
    low = _normalize_title(title).lower()
    if not low:
        return False
    return any(k in low for k in ("封面", "标题页", "cover", "title"))


def _is_toc_title(title: str) -> bool:
    low = _normalize_title(title).lower()
    if not low:
        return False
    return any(k in low for k in ("目录", "议程", "agenda", "contents", "toc"))


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
    return _clean_bullets(items, limit=80)


def _style_instruction(style: Any) -> str:
    normalized = _normalize_style(style)
    if normalized == "technical":
        return "技术汇报口径：强调系统机制、技术依据、实现路径和可验证指标。"
    return "管理汇报口径：强调业务目标、关键结论、风险与决策建议。"


def _merge_outline_material(outline_text: str, material_text: str) -> str:
    sections: list[str] = []
    outline = str(outline_text or "").strip()
    material = str(material_text or "").strip()
    if outline:
        sections.append(f"用户提供的大纲草稿：\n{outline}")
    if material:
        sections.append(f"资料文件内容：\n{material}")
    return "\n\n".join(sections)


def _parse_outline_pages_from_rows(rows: list[Any]) -> list[dict[str, Any]]:
    pages: list[dict[str, Any]] = []
    for row in rows:
        outline_raw = row["outline_content"]
        try:
            outline = json.loads(str(outline_raw or "{}"))
        except Exception:
            outline = {}
        title = _normalize_title(str(outline.get("title") or ""))
        if not title:
            continue
        points = _clean_bullets(list(outline.get("points") or []), limit=8)
        pages.append({"title": title, "points": points})
    return pages


def _safe_json_object(text: str) -> dict[str, Any]:
    payload = str(text or "").strip()
    if not payload:
        raise ValueError("empty response")

    if payload.startswith("```"):
        lines = payload.splitlines()
        if lines and lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].strip() == "```":
            lines = lines[:-1]
        payload = "\n".join(lines).strip()

    candidates: list[str] = [payload]

    start = payload.find("{")
    end = payload.rfind("}")
    if start >= 0 and end > start:
        candidates.append(payload[start : end + 1])

    last_error: Exception | None = None
    for candidate in candidates:
        try:
            parsed = json.loads(candidate)
            if isinstance(parsed, dict):
                return parsed
        except Exception as exc:  # noqa: PERF203
            last_error = exc

    raise ValueError(f"invalid json object: {last_error}")


def _material_excerpt(material_text: str, title: str, max_chars: int = 2200) -> str:
    material = str(material_text or "").strip()
    if not material:
        return ""

    title_tokens = re.findall(r"[A-Za-z]{2,}|[\u4e00-\u9fff]{2,}", str(title or ""))
    title_tokens = [t.lower() for t in title_tokens][:8]

    parts = re.split(r"[。\n！？!?;；]", material)
    scored: list[tuple[int, str]] = []
    for part in parts:
        s = re.sub(r"\s+", " ", part).strip()
        if len(s) < 8:
            continue
        score = 0
        low = s.lower()
        for token in title_tokens:
            if token in low:
                score += 1
        scored.append((score, s))

    scored.sort(key=lambda x: (x[0], len(x[1])), reverse=True)
    if not scored:
        return material[:max_chars]

    picked: list[str] = []
    total = 0
    for score, sentence in scored:
        if score <= 0 and picked:
            break
        if total + len(sentence) + 1 > max_chars:
            continue
        picked.append(sentence)
        total += len(sentence) + 1
        if len(picked) >= 12:
            break

    if not picked:
        return material[:max_chars]
    return "\n".join(picked)


def _default_body_points(title: str) -> list[str]:
    base = _normalize_title(title) or "本页主题"
    return [
        f"{base}的背景与问题定义",
        f"{base}的关键事实与核心依据",
        f"{base}的执行建议与下一步计划",
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
    content_format = "summary_plus_three" if len(final_bullets) >= 3 else "summary_plus_two"

    return {
        "title": final_title,
        "bullets": final_bullets,
        "detail_points": final_bullets,
        "summary_text": summary,
        "text_blocks": final_bullets[:4],
        "content_format": content_format,
        "layout_profile": content_format,
        "notes": str(notes or "").strip(),
        "slide_type": slide_type,
        "evidence": final_bullets[:3],
        "chart_data": None,
        "text": raw_text.strip() if raw_text.strip() else "\n".join(final_bullets),
        "generated_image_path": None,
    }


class NewBackendFlowEngine:
    def __init__(self, use_mock: bool = False) -> None:
        self.client = ModelClient()
        # If model config is missing, degrade gracefully to deterministic fallback.
        self.use_mock = bool(use_mock) or (not self.client.enabled())

    def _chat(self, prompt: str, temperature: float = 0.0) -> str:
        return self.client.chat_text(
            system_prompt="你是资深PPT策划与文案专家，严格遵循输出格式。",
            user_prompt=prompt,
            temperature=temperature,
        )

    def _mock_outline_pages(self) -> list[dict[str, Any]]:
        return [
            {"title": "封面", "points": ["主题", "汇报人", "日期"]},
            {"title": "目录", "points": ["背景", "现状", "方案", "计划", "风险", "结论"]},
            {"title": "背景与目标", "points": _default_body_points("背景与目标")},
            {"title": "现状分析", "points": _default_body_points("现状分析")},
            {"title": "关键问题", "points": _default_body_points("关键问题")},
            {"title": "解决方案", "points": _default_body_points("解决方案")},
            {"title": "实施路径", "points": _default_body_points("实施路径")},
            {"title": "总结与建议", "points": _default_body_points("总结与建议")},
        ]

    def _outline_prompt(self, topic: str, material_text: str, target_pages: int, style: str) -> str:
        material = str(material_text or "").strip()
        style_hint = _style_instruction(style)
        return f"""
请为“{topic}”生成一个详细的PPT大纲，参考 Ai-To-PPTX 的结构风格。
风格要求：{style_hint}

必须严格使用以下 Markdown 结构：
# PPT总标题
## 章节标题
### 页面标题
1.1.1 页面要点1
1.1.2 页面要点2
1.1.3 页面要点3

硬性约束：
1. “### 页面标题”的总数量必须是 {target_pages} 页。
2. 第1页标题必须是“封面”，第2页标题必须是“目录”。
3. 其余页面按主题展开，尽量形成行业总结型汇报。
4. 每个页面下给出3个以“数字.数字.数字”开头的要点行。
5. 不输出与大纲无关的解释文本，不要输出代码块。
6. 必须结合资料内容；资料缺失的信息不要虚构。

资料内容如下：
{material if material else "（未提供资料文件）"}
""".strip()

    def _parse_outline_markdown(self, markdown_text: str) -> list[dict[str, Any]]:
        pages: list[dict[str, Any]] = []
        current_page: dict[str, Any] | None = None

        for raw in str(markdown_text or "").splitlines():
            line = str(raw).strip()
            if not line:
                continue

            if line.startswith("### "):
                if current_page:
                    pages.append(current_page)
                current_page = {"title": _normalize_title(line[4:]), "points": []}
                continue

            if current_page is None and line.startswith("## "):
                # Fallback: if model did not output ###, treat ## as page.
                pages.append({"title": _normalize_title(line[3:]), "points": []})
                continue

            if current_page is not None:
                if line.startswith("- "):
                    current_page["points"].append(_normalize_title(line[2:]))
                    continue
                if re.match(r"^\d+(?:\.\d+){1,3}\s+", line):
                    current_page["points"].append(
                        _normalize_title(re.sub(r"^\d+(?:\.\d+){1,3}\s+", "", line))
                    )
                    continue

        if current_page:
            pages.append(current_page)

        cleaned: list[dict[str, Any]] = []
        seen: set[str] = set()
        for page in pages:
            title = _normalize_title(page.get("title") or "")
            if not title:
                continue
            key = title.lower()
            if key in seen:
                continue
            seen.add(key)
            cleaned.append(
                {
                    "title": title,
                    "points": _clean_bullets(list(page.get("points") or []), limit=6),
                }
            )
        return cleaned

    def _ensure_target_outline(self, pages: list[dict[str, Any]], target_pages: int) -> list[dict[str, Any]]:
        target = max(8, min(12, int(target_pages)))

        cover = next((p for p in pages if _is_cover_title(str(p.get("title")))), None)
        toc = next((p for p in pages if _is_toc_title(str(p.get("title")))), None)

        body: list[dict[str, Any]] = []
        seen: set[str] = set()
        for page in pages:
            title = _normalize_title(page.get("title") or "")
            if not title:
                continue
            if _is_cover_title(title) or _is_toc_title(title):
                continue
            key = title.lower()
            if key in seen:
                continue
            seen.add(key)
            body.append(
                {
                    "title": title,
                    "points": _clean_bullets(list(page.get("points") or []), limit=6),
                }
            )

        required_body = max(0, target - 2)
        while len(body) < required_body:
            idx = len(body) + 1
            fallback_title = f"补充页{idx}"
            body.append({"title": fallback_title, "points": _default_body_points(fallback_title)})

        body = body[:required_body]

        cover_page = {
            "title": _normalize_title((cover or {}).get("title") or "封面"),
            "points": _clean_bullets((cover or {}).get("points") or ["主题", "汇报人", "日期"], limit=4),
        }
        if not cover_page["points"]:
            cover_page["points"] = ["主题", "汇报人", "日期"]

        toc_points = [item["title"] for item in body[:6]]
        toc_page = {
            "title": _normalize_title((toc or {}).get("title") or "目录"),
            "points": toc_points if toc_points else ["内容概览"],
        }

        return [cover_page, toc_page, *body]

    def generate_outline_pages(
        self,
        topic: str,
        material_text: str,
        target_pages: int,
        style: str = "management",
    ) -> list[dict[str, Any]]:
        base_pages = self._mock_outline_pages()
        if self.use_mock:
            return self._ensure_target_outline(base_pages, target_pages)

        try:
            raw = self._chat(
                self._outline_prompt(topic, material_text, target_pages, style),
                temperature=0.0,
            )
            parsed = self._parse_outline_markdown(raw)
            return self._ensure_target_outline(parsed, target_pages)
        except Exception:
            return self._ensure_target_outline(base_pages, target_pages)

    def _page_prompt(
        self,
        topic: str,
        material_text: str,
        outline_titles: list[str],
        page_title: str,
        page_points: list[str],
        style: str,
    ) -> str:
        style_hint = _style_instruction(style)
        outline_text = "\n".join([f"{i + 1}. {t}" for i, t in enumerate(outline_titles)])
        points_text = "\n".join([f"- {p}" for p in page_points]) if page_points else "- （无显式要点）"
        material_excerpt = _material_excerpt(material_text, page_title, max_chars=2600)
        return f"""
你是 Ai-To-PPTX 的单页文案生成器。
请根据“主题 + 当前页标题 + 输入资料”生成该页内容。
风格要求：{style_hint}

主题：{topic}
当前页标题：{page_title}
全局大纲标题：
{outline_text}

当前页要点：
{points_text}

资料（仅可使用以下信息）：
{material_excerpt if material_excerpt else "（资料为空）"}

要求：
1. 必须围绕“当前页标题”写作，不能偏题。
2. 内容必须依据资料，禁止编造资料中没有的数字、百分比、时间和结论。
3. 如果资料不足，请用审慎表述，不要虚构事实。
4. bullets 生成 4-6 条，每条 16-40 字，表达完整、可直接上屏。
5. 仅输出 JSON 对象，不要输出其他文字。

JSON格式：
{{
  "title": "页面标题",
  "bullets": ["要点1", "要点2", "要点3", "要点4"],
  "notes": "演讲备注，1-3句"
}}
""".strip()

    def _fallback_page_payload(
        self,
        raw_text: str,
        page_title: str,
        page_points: list[str],
        material_text: str,
    ) -> dict[str, Any]:
        lines = [str(x).strip() for x in str(raw_text or "").splitlines()]
        bullets: list[str] = []
        for line in lines:
            cleaned = _normalize_title(line)
            if not cleaned:
                continue
            if cleaned in {"{", "}", "[", "]"}:
                continue
            if ":" in cleaned and cleaned.split(":", 1)[0].strip().lower() in {"title", "bullets", "notes"}:
                continue
            if cleaned.startswith('"') and cleaned.endswith('"'):
                cleaned = cleaned.strip('"')
            if len(cleaned) < 8:
                continue
            bullets.append(cleaned)
            if len(bullets) >= 6:
                break

        normalized_bullets = _clean_bullets(bullets, limit=6)
        if len(normalized_bullets) < 3:
            normalized_bullets = _clean_bullets(page_points, limit=6)
        if len(normalized_bullets) < 3:
            normalized_bullets = _default_body_points(page_title)

        notes = _material_excerpt(material_text, page_title, max_chars=220) or "资料不足，建议补充原始文件后再生成。"
        return _format_slide_payload(
            title=page_title,
            bullets=normalized_bullets,
            notes=notes,
            slide_type="summary",
            raw_text=raw_text,
        )

    def generate_page_payload(
        self,
        topic: str,
        material_text: str,
        outline_titles: list[str],
        page_outline: dict[str, Any],
        page_index: int,
        total_pages: int,
        style: str = "management",
    ) -> dict[str, Any]:
        _ = total_pages
        style_normalized = _normalize_style(style)
        page_title = _normalize_title(page_outline.get("title") or f"第{page_index}页")
        page_points = _clean_bullets(list(page_outline.get("points") or []), limit=6)

        if page_index == 1 or _is_cover_title(page_title):
            return _format_slide_payload(
                title=topic or page_title,
                bullets=["主题", "汇报人", "日期"],
                notes="封面页",
                slide_type="title",
            )

        if page_index == 2 or _is_toc_title(page_title):
            toc_items = [t for t in outline_titles if not _is_cover_title(t) and not _is_toc_title(t)]
            return _format_slide_payload(
                title="目录",
                bullets=toc_items[:8],
                notes="目录页",
                slide_type="toc",
            )

        if self.use_mock:
            bullets = page_points if page_points else _default_body_points(page_title)
            note = _material_excerpt(material_text, page_title, max_chars=180) or "资料不足，建议补充原始文件后再生成。"
            return _format_slide_payload(
                title=page_title,
                bullets=bullets,
                notes=note,
                slide_type="summary",
            )

        try:
            raw = self._chat(
                self._page_prompt(
                    topic,
                    material_text,
                    outline_titles,
                    page_title,
                    page_points,
                    style_normalized,
                ),
                temperature=0.1,
            )
        except Exception:
            return self._fallback_page_payload(
                raw_text="",
                page_title=page_title,
                page_points=page_points,
                material_text=material_text,
            )

        try:
            parsed = _safe_json_object(raw)
        except Exception:
            return self._fallback_page_payload(
                raw_text=raw,
                page_title=page_title,
                page_points=page_points,
                material_text=material_text,
            )

        title = _normalize_title(str(parsed.get("title") or page_title)) or page_title
        bullets = _clean_bullets(list(parsed.get("bullets") or []), limit=6)
        if len(bullets) < 3:
            bullets = _clean_bullets(page_points, limit=6)
        if len(bullets) < 3:
            bullets = _default_body_points(title)

        notes = str(parsed.get("notes") or "").strip()
        if not notes:
            notes = _material_excerpt(material_text, title, max_chars=220) or "资料不足，建议补充原始文件后再生成。"

        return _format_slide_payload(
            title=title,
            bullets=bullets,
            notes=notes,
            slide_type="summary",
            raw_text=raw,
        )


_engine = NewBackendFlowEngine(use_mock=settings.use_mock_llm)


class OutlinePreviewAdapter:
    def generate_outline(self, title: str, style: str, material: str, target_pages: int) -> list[str]:
        pages = _engine.generate_outline_pages(title, material, target_pages, style=_normalize_style(style))
        return [str(page.get("title") or f"第{i + 1}页") for i, page in enumerate(pages)]


llm = OutlinePreviewAdapter()


def _outline_list_to_pages(items: list[str], target_pages: int) -> list[dict[str, Any]]:
    cleaned = clean_outline_items(items)
    if not cleaned:
        return _engine.generate_outline_pages("项目汇报", "", target_pages)

    pages = [{"title": _normalize_title(item), "points": []} for item in cleaned if _normalize_title(item)]
    if not pages:
        return _engine.generate_outline_pages("项目汇报", "", target_pages)
    return _engine._ensure_target_outline(pages, target_pages)


def get_outline_for_project(
    project_row: dict[str, Any] | Any,
    requested_outline: list[str] | None = None,
) -> list[dict[str, Any]]:
    target_pages = int(project_row["target_pages"])
    if requested_outline:
        return _outline_list_to_pages(requested_outline, target_pages)

    title = str(project_row["title"] or "项目汇报")
    material_text = _merge_outline_material(
        str(project_row["outline_text"] or ""),
        str(project_row["material_text"] or ""),
    )
    style = _normalize_style(project_row.get("style") if hasattr(project_row, "get") else project_row["style"])
    return _engine.generate_outline_pages(title, material_text, target_pages, style=style)


def rebuild_project_pages(project_id: str, outline_pages: list[dict[str, Any]]) -> None:
    legacy_rebuild_project_pages(project_id, outline_pages)


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

    total = len(pages)
    update_task(
        task_id,
        {
            "status": "PROCESSING",
            "progress_json": make_progress(total, 0, 0, "generating_descriptions"),
        },
    )

    outline_pages = _parse_outline_pages_from_rows(pages)
    outline_titles = [str(item.get("title") or f"第{i + 1}页") for i, item in enumerate(outline_pages)]
    topic = str(project["title"] or "项目汇报")
    material_text = str(project["material_text"] or "")
    style = _normalize_style(project["style"])

    completed = 0
    failed = 0

    for idx, row in enumerate(pages):
        page_id = str(row["page_id"])
        try:
            page_outline = outline_pages[idx] if idx < len(outline_pages) else {"title": f"第{idx + 1}页", "points": []}
            payload = _engine.generate_page_payload(
                topic=topic,
                material_text=material_text,
                outline_titles=outline_titles,
                page_outline=page_outline,
                page_index=idx + 1,
                total_pages=total,
                style=style,
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
            update_task(
                task_id,
                {
                    "error_message": str(exc),
                },
            )

        update_task(
            task_id,
            {
                "progress_json": make_progress(total, completed, failed, "generating_descriptions"),
            },
        )

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
            "progress_json": make_progress(total, completed, failed, "descriptions_done"),
            "error_message": None if failed == 0 else f"{failed} pages failed",
            "completed_at": utc_now_iso(),
        },
    )


def generate_ppt_task(task_id: str, project_id: str) -> None:
    # Reuse the current exporter bridge to keep frontend APIs unchanged.
    legacy_generate_ppt_task(task_id, project_id)
