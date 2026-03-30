from __future__ import annotations



import json

import re

from difflib import SequenceMatcher

from datetime import datetime

from pathlib import Path

from typing import Any, Iterable

from uuid import uuid4



from app.config import settings

from app.services.banana_ai_service import (

    BananaAIService,

    BananaProjectContext,

    build_idea_prompt,

    enforce_target_pages,

    make_project_context_from_row,

)

from app.services.image_generator import image_generator

from app.services.pptx_exporter import export_slides_to_pptx

from app.storage.db import (

    get_project,

    list_pages,

    make_progress,

    replace_pages,

    update_page,

    update_project,

    update_task,

)





banana_ai = BananaAIService(use_mock=settings.use_mock_llm)





class OutlinePreviewAdapter:

    """Compatibility adapter used by /outline/preview."""



    def generate_outline(self, title: str, style: str, material: str, target_pages: int) -> list[str]:

        context = BananaProjectContext(

            idea_prompt=build_idea_prompt(title, style, material),

            creation_type="idea",

        )

        pages = banana_ai.generate_outline(context, language="zh")

        pages = enforce_target_pages(pages, target_pages)

        return [str(page.get("title") or f"第{i + 1}页") for i, page in enumerate(pages)]





llm = OutlinePreviewAdapter()






def _normalize_style(style: Any) -> str:
    return "technical" if str(style or "").lower() == "technical" else "management"

def utc_now_iso() -> str:

    return datetime.utcnow().isoformat()






def clean_outline_items(items: Iterable[str]) -> list[str]:
    out: list[str] = []
    for raw in items:
        txt = _cleanup_text(str(raw or ""), max_len=120)
        if not txt:
            continue
        txt = re.sub(r"^\d+\s*[\.\u3001\)\]\uff09]\s*", "", txt)
        txt = txt.strip()
        if txt:
            out.append(txt)
    return out


def _outline_list_to_pages(items: list[str]) -> list[dict[str, Any]]:
    pages: list[dict[str, Any]] = []
    for idx, item in enumerate(items, start=1):
        title = _cleanup_text(str(item).strip(), max_len=120)
        parts = re.split(r"[\uFF1A:]", title, maxsplit=1)
        if len(parts) == 2:
            title = _cleanup_text(parts[1].strip() or parts[0].strip(), max_len=120)
        if not title:
            title = f"Page {idx}"
        pages.append({"title": title, "points": []})
    return pages


def _context_for_project(project_row: dict[str, Any] | Any) -> BananaProjectContext:
    return make_project_context_from_row(project_row)


def get_outline_for_project(
    project_row: dict[str, Any] | Any,
    requested_outline: list[str] | None = None,
) -> list[dict[str, Any]]:
    target_pages = int(project_row["target_pages"])

    if requested_outline:
        pages = _outline_list_to_pages(clean_outline_items(requested_outline))
        return enforce_target_pages(pages, target_pages)

    outline_text = str(project_row["outline_text"] or "")
    creation_type = str(project_row["creation_type"] or "idea")

    context = _context_for_project(project_row)

    if creation_type == "outline" and outline_text.strip():
        context.creation_type = "outline"
        context.outline_text = outline_text
        pages = banana_ai.parse_outline_text(context, language="zh")
        return enforce_target_pages(pages, target_pages)

    pages = banana_ai.generate_outline(context, language="zh")
    return enforce_target_pages(pages, target_pages)


def rebuild_project_pages(project_id: str, outline_pages: list[dict[str, Any]]) -> None:
    now = utc_now_iso()
    rows: list[dict[str, Any]] = []
    outline_lines: list[str] = []

    for idx, item in enumerate(outline_pages):
        title = str(item.get("title") or f"Page {idx + 1}")
        points = [str(x) for x in list(item.get("points") or []) if str(x).strip()]

        payload: dict[str, Any] = {
            "title": title,
            "points": points,
        }
        if item.get("part"):
            payload["part"] = str(item.get("part"))

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
        outline_lines.append(f"{idx + 1}. {title}")

    replace_pages(project_id, rows)
    update_project(
        project_id,
        {
            "outline_text": "\n".join(outline_lines),
            "status": "OUTLINE_GENERATED",
            "updated_at": now,
        },
    )


def _safe_load_json(raw: str | None, fallback: Any) -> Any:
    if not raw:
        return fallback
    try:
        return json.loads(raw)
    except Exception:
        return fallback


_TITLE_LABEL_KEYS = (
    "page title",
    "title",
    "\u9875\u9762\u6807\u9898",
    "\u6807\u9898",
)
_TEXT_LABEL_KEYS = (
    "page text",
    "content",
    "body",
    "\u9875\u9762\u6587\u5b57",
    "\u9875\u9762\u5185\u5bb9",
    "\u6b63\u6587",
    "\u5185\u5bb9",
)
_NOTES_LABEL_KEYS = (
    "notes",
    "note",
    "materials",
    "material",
    "reference",
    "\u56fe\u7247\u7d20\u6750",
    "\u5176\u4ed6\u9875\u9762\u7d20\u6750",
    "\u89c6\u89c9\u5143\u7d20",
    "\u89c6\u89c9\u7126\u70b9",
    "\u6392\u7248\u5e03\u5c40",
    "\u6f14\u8bb2\u8005\u5907\u6ce8",
    "\u7d20\u6750",
)
_ALL_LABEL_KEYS = _TITLE_LABEL_KEYS + _TEXT_LABEL_KEYS + _NOTES_LABEL_KEYS


def _normalize_newlines(text: str) -> str:
    return str(text or "").replace("\r\n", "\n").replace("\r", "\n").replace("\x00", "").strip()


def _strip_markdown_prefix(text: str) -> str:
    out = str(text or "")
    out = re.sub(r"^\s*#{1,6}\s*", "", out)
    out = re.sub(r"^\s*[-*]+\s*", "", out)
    out = re.sub(r"^\s*\d+\s*[\.\u3001\)\]\uff09]\s*", "", out)
    return out.strip()


def _strip_xml_like(text: str) -> str:
    out = str(text or "")
    out = re.sub(r"</?[A-Za-z_][A-Za-z0-9._:-]*(?:\s[^>\n]*)?>", " ", out)
    out = re.sub(r"&lt;/?[A-Za-z_][^&]{0,120}&gt;", " ", out, flags=re.I)
    out = re.sub(r"\bxmlns(?::\w+)?=\"[^\"]*\"", " ", out, flags=re.I)
    return out


def _cleanup_text(text: str, max_len: int | None = None) -> str:
    out = _normalize_newlines(text)
    out = _strip_markdown_prefix(out)
    out = _strip_xml_like(out)
    out = re.sub(r"\s+", " ", out).strip(" -:\t\n")
    if max_len and len(out) > max_len:
        out = out[:max_len].rstrip()
    return out



def _normalize_compare_key(text: str) -> str:
    cleaned = _cleanup_text(text, max_len=220).lower()
    cleaned = re.sub(r"(本页|本章|本节|本部分|页面|章节|小标题|标题)", "", cleaned)
    cleaned = re.sub(r"[\s\-_.,:;!?()\[\]{}<>/\\|]+", "", cleaned)
    cleaned = re.sub(r"[^\w\u4e00-\u9fff%]+", "", cleaned)
    return cleaned


def _is_title_redundant_line(line: str, title: str) -> bool:
    line_key = _normalize_compare_key(line)
    title_key = _normalize_compare_key(title)

    if not line_key or not title_key:
        return False
    if line_key == title_key:
        return True
    if line_key.startswith(title_key) and (len(line_key) - len(title_key) <= 10):
        return True
    if title_key.startswith(line_key) and (len(title_key) - len(line_key) <= 6):
        return True

    ratio = SequenceMatcher(None, line_key, title_key).ratio()
    return ratio >= 0.9 and abs(len(line_key) - len(title_key)) <= 8


def _split_sentences(text: str) -> list[str]:
    parts = re.split(r"(?:\n+|[。！？!?；;])", _normalize_newlines(text))
    out: list[str] = []
    for part in parts:
        item = _cleanup_text(part, max_len=180)
        if len(item) >= 6:
            out.append(item)
    return out
def _parse_labeled_line(line: str) -> tuple[str | None, str]:
    m = re.match(r"^\s*([^:\uFF1A]{1,40})[\uFF1A:]\s*(.*)$", str(line or ""))
    if not m:
        return None, ""
    return m.group(1).strip().lower(), m.group(2).strip()


def _is_label(label: str | None, keys: tuple[str, ...]) -> bool:
    if not label:
        return False
    val = str(label).strip().lower()
    return any(k in val for k in keys)


def _extract_labeled_section(raw_text: str, start_keys: tuple[str, ...], stop_keys: tuple[str, ...] | None = None) -> str:
    text_norm = _normalize_newlines(raw_text)
    lines = text_norm.split("\n")
    stop_set = stop_keys or _ALL_LABEL_KEYS

    capture = False
    out_lines: list[str] = []
    for raw_line in lines:
        line = raw_line.strip()
        label, value = _parse_labeled_line(line)

        if label and _is_label(label, start_keys):
            capture = True
            if value:
                out_lines.append(value)
            continue

        if capture and label and _is_label(label, stop_set):
            break

        if capture:
            out_lines.append(line)

    return "\n".join(out_lines).strip()




def _has_numeric_signal(title: str, bullets: list[str]) -> bool:
    text = f"{title} {' '.join(bullets)}"
    return bool(re.search(r"(?<![A-Za-z])[-+]?\d+(?:\.\d+)?", text))


def _infer_slide_type(title: str, bullets: list[str], page_index: int, total: int) -> str:
    full = f"{title} {' '.join(bullets)}".lower()
    title_lower = str(title or "").lower()

    def contains_any(value: str, words: tuple[str, ...]) -> bool:
        return any(w in value for w in words)

    if page_index == 1 or contains_any(title_lower, ("\u5c01\u9762", "\u6807\u9898", "cover", "title")):
        return "title"
    if page_index == 2 or contains_any(title_lower, ("\u76ee\u5f55", "\u8bae\u7a0b", "agenda", "contents", "toc")):
        return "toc"

    # Avoid over-classifying slides as risk just because they contain "问题/issue".
    strict_risk_keywords = (
        "\u98ce\u9669",
        "\u9690\u60a3",
        "\u6f0f\u6d1e",
        "\u5a01\u80c1",
        "\u963b\u585e",
        "\u963b\u788d",
        "\u7f3a\u9677",
        "\u5b89\u5168",
        "\u5408\u89c4",
        "risk",
        "threat",
        "vulnerability",
        "blocker",
    )
    problem_keywords = ("\u95ee\u9898", "\u6311\u6218", "issue", "challenge")
    mitigation_keywords = ("\u5e94\u5bf9", "\u7f13\u89e3", "\u6cbb\u7406", "\u89c4\u907f", "mitigation", "countermeasure")
    non_risk_problem_context = ("\u95ee\u9898\u5b9a\u4e49", "\u95ee\u9898\u63d0\u51fa", "\u4f18\u5316\u95ee\u9898", "\u7814\u7a76\u95ee\u9898")

    if contains_any(full, strict_risk_keywords):
        return "risk"
    if contains_any(full, problem_keywords) and contains_any(full, mitigation_keywords):
        return "risk"
    if contains_any(full, problem_keywords) and not contains_any(full, non_risk_problem_context):
        if contains_any(full, ("\u5931\u6548", "\u635f\u5931", "\u5f02\u5e38", "\u6545\u969c", "failure", "loss")):
            return "risk"

    if contains_any(full, non_risk_problem_context):
        return "summary"

    if contains_any(full, ("\u653b\u51fb", "\u5a01\u80c1\u5efa\u6a21", "threat model", "attack")):
        return "risk"
    if contains_any(full, ("\u8ba1\u5212", "\u8def\u7ebf", "\u91cc\u7a0b\u7891", "\u9636\u6bb5", "\u8fdb\u5ea6", "\u6392\u671f", "timeline", "plan", "roadmap", "milestone")):
        return "timeline"

    # Data page only when data-related wording and concrete numeric signal both exist.
    if _has_numeric_signal(title, bullets) and contains_any(full, ("\u6570\u636e", "\u6307\u6807", "\u540c\u6bd4", "\u73af\u6bd4", "\u589e\u957f", "%", "roi", "gmv", "metric", "kpi", "trend")):
        return "data"

    return "summary"


def _extract_title(raw_text: str, fallback: str) -> str:
    raw = _normalize_newlines(raw_text)

    section_title = _extract_labeled_section(raw, _TITLE_LABEL_KEYS, _TEXT_LABEL_KEYS + _NOTES_LABEL_KEYS)
    title = _cleanup_text(section_title, max_len=90)
    if title:
        return title

    for line in raw.split("\n"):
        label, value = _parse_labeled_line(line)
        if label and _is_label(label, _TITLE_LABEL_KEYS):
            t = _cleanup_text(value, max_len=90)
            if t:
                return t
            continue
        if label and _is_label(label, _ALL_LABEL_KEYS):
            continue
        candidate = _cleanup_text(line, max_len=90)
        if candidate:
            return candidate

    cleaned_fallback = _cleanup_text(fallback, max_len=90)
    return cleaned_fallback or str(fallback or "Untitled")


def _extract_page_text_section(raw_text: str) -> str:
    raw = _normalize_newlines(raw_text)

    section = _extract_labeled_section(raw, _TEXT_LABEL_KEYS, _NOTES_LABEL_KEYS)
    if section.strip():
        return section.strip()

    lines: list[str] = []
    for line in raw.split("\n"):
        label, value = _parse_labeled_line(line)
        if label and _is_label(label, _TITLE_LABEL_KEYS):
            continue
        if label and _is_label(label, _NOTES_LABEL_KEYS):
            continue
        if label and _is_label(label, _TEXT_LABEL_KEYS):
            if value.strip():
                lines.append(value.strip())
            continue
        lines.append(line.strip())

    return "\n".join(lines).strip()


_BULLET_PREFIX_RE = re.compile(r"^\s*(?:[-*]|[0-9]{1,2}\s*[\.\)\]])\s*")


def _is_explicit_bullet_line(line: str) -> bool:
    return bool(_BULLET_PREFIX_RE.match(str(line or "")))


def _strip_bullet_prefix(line: str) -> str:
    return _BULLET_PREFIX_RE.sub("", str(line or "")).strip()


def _extract_text_blocks(text_section: str, title: str = "") -> list[str]:
    text_body = re.sub(r"```[\s\S]*?```", " ", _normalize_newlines(text_section))

    paragraphs: list[str] = []
    current_lines: list[str] = []
    seen_keys: set[str] = set()

    def flush_current() -> None:
        nonlocal current_lines
        if not current_lines:
            return
        merged = _cleanup_text(" ".join(current_lines), max_len=280)
        current_lines = []
        if not merged:
            return
        if _is_title_redundant_line(merged, title):
            return
        key = _normalize_compare_key(merged)
        if not key or key in seen_keys:
            return
        seen_keys.add(key)
        paragraphs.append(merged)

    for raw_line in text_body.split("\n"):
        stripped = str(raw_line or "").strip()
        if not stripped:
            flush_current()
            continue

        label, value = _parse_labeled_line(stripped)
        source = value if (label and _is_label(label, _ALL_LABEL_KEYS)) else stripped
        if not source:
            continue
        if _is_explicit_bullet_line(source):
            continue

        cleaned = _cleanup_text(source, max_len=220)
        if not cleaned:
            continue
        if _is_title_redundant_line(cleaned, title):
            continue
        if re.search(r"</?[A-Za-z_][A-Za-z0-9._:-]*", cleaned):
            continue

        current_lines.append(cleaned)

    flush_current()

    if not paragraphs:
        for sent in _split_sentences(text_body):
            if _is_title_redundant_line(sent, title):
                continue
            key = _normalize_compare_key(sent)
            if not key or key in seen_keys:
                continue
            seen_keys.add(key)
            paragraphs.append(sent)
            if len(paragraphs) >= 3:
                break

    return paragraphs[:4]


def _extract_bullets(text_section: str, fallback_points: list[str], title: str = "") -> list[str]:
    bullet_lines: list[str] = []
    seen_keys: set[str] = set()
    text_body = re.sub(r"```[\s\S]*?```", " ", _normalize_newlines(text_section))

    def push_candidate(raw_line: str) -> None:
        cleaned = _cleanup_text(raw_line, max_len=180)
        if not cleaned:
            return
        if re.search(r"</?[A-Za-z_][A-Za-z0-9._:-]*", cleaned):
            return
        if _is_title_redundant_line(cleaned, title):
            return
        key = _normalize_compare_key(cleaned)
        if not key or key in seen_keys:
            return
        seen_keys.add(key)
        bullet_lines.append(cleaned)

    for raw in text_body.splitlines():
        source = str(raw or "").strip()
        if not source:
            continue

        label, value = _parse_labeled_line(source)
        if label and _is_label(label, _ALL_LABEL_KEYS):
            source = value
        if not source:
            continue

        if _is_explicit_bullet_line(source):
            push_candidate(_strip_bullet_prefix(source))
            if len(bullet_lines) >= 6:
                break
            continue

        cleaned = _cleanup_text(source, max_len=180)
        if 8 <= len(cleaned) <= 36 and not re.search(r"[。！？!?；;]$", cleaned):
            push_candidate(cleaned)
            if len(bullet_lines) >= 6:
                break

    if len(bullet_lines) < 2:
        for point in fallback_points:
            push_candidate(str(point))
            if len(bullet_lines) >= 4:
                break

    if len(bullet_lines) < 2:
        for sent in _split_sentences(text_body):
            if len(sent) < 10:
                continue
            push_candidate(sent)
            if len(bullet_lines) >= 3:
                break

    while bullet_lines and _is_title_redundant_line(bullet_lines[0], title):
        bullet_lines.pop(0)

    return bullet_lines[:6]


def _derive_content_format(summary_text: str, detail_points: list[str], text_blocks: list[str], slide_type: str) -> str:
    if slide_type in {"risk", "timeline", "data"}:
        return "typed_slide"

    point_count = len(detail_points)
    block_count = len(text_blocks)
    has_summary = bool(summary_text)

    if point_count == 0:
        return "narrative_split" if block_count >= 2 else "narrative_banner"
    if has_summary and point_count == 2:
        return "summary_plus_two"
    if has_summary and point_count == 3:
        return "summary_plus_three"
    if has_summary and point_count >= 4 and block_count >= 2:
        return "center_hub_four"
    if has_summary and point_count >= 4:
        return "summary_plus_four"
    if point_count == 3:
        return "points_three_columns"
    if point_count == 4:
        return "points_four_grid"
    if point_count >= 5:
        return "points_five_split"
    if has_summary and block_count >= 2:
        return "top_bottom_story"
    if has_summary:
        return "left_summary_right_points"
    if block_count >= 3:
        return "four_quadrant_mixed"
    return "mixed_content"


def _extract_notes(raw_text: str, extra_fields: dict[str, str] | None) -> str:
    notes_parts: list[str] = []

    notes_section = _extract_labeled_section(raw_text, _NOTES_LABEL_KEYS, ())
    cleaned_notes = _cleanup_text(notes_section, max_len=800)
    if cleaned_notes:
        notes_parts.append(cleaned_notes)

    if extra_fields:
        for name, value in extra_fields.items():
            value_s = _cleanup_text(str(value or ""), max_len=200)
            if value_s:
                notes_parts.append(f"{name}: {value_s}")

    merged = "\n".join(notes_parts).strip()
    return merged or "Generated from banana workflow"




def _extract_chart_data_from_text(title: str, bullets: list[str], notes: str) -> dict[str, Any] | None:
    unit_candidates: list[str] = []
    pairs: list[tuple[str, float]] = []

    candidates = [str(x) for x in bullets]
    candidates.extend([line.strip() for line in str(notes or "").splitlines() if line.strip()])

    for raw in candidates:
        line = _cleanup_text(raw, max_len=200)
        if not line:
            continue

        m = re.search(r"([-+]?\d+(?:\.\d+)?)\s*(%|x|\u500d|ms|s|h|\u4e07\u5143|\u4ebf\u5143|\u4e07|\u5143|\u4e2a|\u4eba)?", line)
        if not m:
            continue

        try:
            value = float(m.group(1))
        except Exception:
            continue

        unit = str(m.group(2) or "").strip()
        if unit:
            unit_candidates.append(unit)

        label = line[: m.start()] + line[m.end() :]
        label = re.sub(r"[\uff1a:\uff0c,\u3002\uff1b;\uff08\uff09()\[\]\-]+", " ", label)
        label = re.sub(r"\s+", " ", label).strip()
        if not label:
            label = f"\u6307\u6807{len(pairs) + 1}"

        pairs.append((label[:24], value))

    if len(pairs) < 2:
        return None

    labels: list[str] = []
    values: list[float] = []
    seen: set[str] = set()
    for label, value in pairs:
        key = label.lower()
        if key in seen:
            continue
        seen.add(key)
        labels.append(label)
        values.append(value)
        if len(labels) >= 6:
            break

    if len(labels) < 2:
        return None

    unit = ""
    if unit_candidates:
        counts: dict[str, int] = {}
        for u in unit_candidates:
            counts[u] = counts.get(u, 0) + 1
        unit = max(counts.items(), key=lambda x: x[1])[0]

    return {
        "labels": labels,
        "values": values,
        "unit": unit,
    }


def _description_to_slide_payload(
    description_text: str,
    page_outline: dict[str, Any],
    page_index: int,
    total_pages: int,
    extra_fields: dict[str, str] | None = None,
) -> dict[str, Any]:
    fallback_title = str(page_outline.get("title") or f"Page {page_index}")
    fallback_points = [str(x) for x in list(page_outline.get("points") or [])]

    normalized_text = _normalize_newlines(description_text)
    title = _extract_title(normalized_text, fallback_title)
    page_text = _extract_page_text_section(normalized_text)

    bullets = _extract_bullets(page_text, fallback_points, title=title)
    text_blocks = _extract_text_blocks(page_text, title=title)

    summary_text = ""
    if text_blocks:
        summary_text = _cleanup_text(text_blocks[0], max_len=240)
    elif bullets:
        summary_text = _cleanup_text(bullets[0], max_len=240)
    elif fallback_points:
        summary_text = _cleanup_text(str(fallback_points[0]), max_len=240)

    detail_points: list[str] = []
    seen_keys: set[str] = set()

    def push_point(raw: str) -> None:
        cleaned = _cleanup_text(raw, max_len=180)
        if not cleaned:
            return
        if summary_text and _normalize_compare_key(cleaned) == _normalize_compare_key(summary_text):
            return
        key = _normalize_compare_key(cleaned)
        if not key or key in seen_keys:
            return
        seen_keys.add(key)
        detail_points.append(cleaned)

    for item in bullets:
        push_point(item)
        if len(detail_points) >= 6:
            break

    if len(detail_points) < 2:
        for item in fallback_points:
            push_point(str(item))
            if len(detail_points) >= 4:
                break

    if len(detail_points) < 2:
        for sent in _split_sentences(page_text):
            if len(sent) < 10:
                continue
            push_point(sent)
            if len(detail_points) >= 3:
                break

    notes = _extract_notes(normalized_text, extra_fields)
    slide_type = _infer_slide_type(title, detail_points, page_index, total_pages)

    chart_data = _extract_chart_data_from_text(title, detail_points, notes) if slide_type == "data" else None
    if slide_type == "data" and chart_data is None:
        slide_type = "summary"

    content_format = _derive_content_format(summary_text, detail_points, text_blocks, slide_type)

    evidence_source = detail_points if detail_points else text_blocks
    evidence = [str(x) for x in evidence_source[:3]]

    payload: dict[str, Any] = {
        "title": title,
        "bullets": detail_points,
        "detail_points": detail_points,
        "summary_text": summary_text,
        "text_blocks": text_blocks,
        "content_format": content_format,
        "layout_profile": content_format,
        "notes": notes,
        "slide_type": slide_type,
        "evidence": evidence,
        "chart_data": chart_data,
        "text": normalized_text,
        "generated_image_path": None,
    }
    if extra_fields:
        payload["extra_fields"] = extra_fields
    return payload


def _outline_pages_from_db(pages: list[Any]) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    for idx, page in enumerate(pages, start=1):
        outline = _safe_load_json(page["outline_content"], {})
        out.append(
            {
                "title": str(outline.get("title") or f"Page {idx}"),
                "points": [str(x) for x in list(outline.get("points") or [])],
                "part": outline.get("part"),
            }
        )
    return out


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



    try:

        context = _context_for_project(project)

        outline = _outline_pages_from_db(pages)



        completed = 0

        failed = 0



        for idx, page in enumerate(pages):

            page_id = str(page["page_id"])

            page_outline = outline[idx]

            try:

                result = banana_ai.generate_page_description(

                    project_context=context,

                    outline=outline,

                    page_outline=page_outline,

                    page_index=idx + 1,

                    language="zh",

                    detail_level="detailed",

                )

                desc_text = str(result.get("text") or "")

                extra_fields = result.get("extra_fields") if isinstance(result.get("extra_fields"), dict) else None



                payload = _description_to_slide_payload(

                    description_text=desc_text,

                    page_outline=page_outline,

                    page_index=idx + 1,

                    total_pages=total,

                    extra_fields=extra_fields,

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

            except Exception:

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

    except Exception as exc:

        update_task(

            task_id,

            {

                "status": "FAILED",

                "error_message": str(exc),

                "completed_at": utc_now_iso(),

            },

        )





def _collect_project_slides(project: Any, pages: list[Any]) -> list[dict[str, Any]]:

    total = len(pages)

    outline = _outline_pages_from_db(pages)



    slides: list[dict[str, Any]] = []

    for idx, page in enumerate(pages):

        desc = _safe_load_json(page["description_content"], None)

        page_outline = outline[idx]



        if not isinstance(desc, dict):

            result = banana_ai.generate_page_description(

                project_context=_context_for_project(project),

                outline=outline,

                page_outline=page_outline,

                page_index=idx + 1,

                language="zh",

                detail_level="detailed",

            )

            desc = _description_to_slide_payload(

                description_text=str(result.get("text") or ""),

                page_outline=page_outline,

                page_index=idx + 1,

                total_pages=total,

                extra_fields=result.get("extra_fields") if isinstance(result.get("extra_fields"), dict) else None,

            )

            update_page(

                str(page["page_id"]),

                {

                    "description_content": json.dumps(desc, ensure_ascii=False),

                    "status": "DESCRIPTION_GENERATED",

                    "updated_at": utc_now_iso(),

                },

            )



        if "text" in desc and "bullets" in desc and "title" in desc:

            normalized = desc

        else:

            normalized = _description_to_slide_payload(

                description_text=str(desc.get("text") or ""),

                page_outline=page_outline,

                page_index=idx + 1,

                total_pages=total,

                extra_fields=desc.get("extra_fields") if isinstance(desc.get("extra_fields"), dict) else None,

            )

            update_page(

                str(page["page_id"]),

                {

                    "description_content": json.dumps(normalized, ensure_ascii=False),

                    "status": "DESCRIPTION_GENERATED",

                    "updated_at": utc_now_iso(),

                },

            )



        slide = {
            "page": idx + 1,
            "title": str(normalized.get("title") or page_outline["title"]),
            "bullets": [str(x) for x in list(normalized.get("bullets") or [])],
            "detail_points": [str(x) for x in list(normalized.get("detail_points") or normalized.get("bullets") or [])],
            "summary_text": str(normalized.get("summary_text") or ""),
            "text_blocks": [str(x) for x in list(normalized.get("text_blocks") or [])],
            "content_format": str(normalized.get("content_format") or ""),
            "layout_profile": str(normalized.get("layout_profile") or ""),
            "notes": str(normalized.get("notes") or ""),
            "slide_type": str(normalized.get("slide_type") or "summary"),
            "evidence": normalized.get("evidence"),
            "chart_data": normalized.get("chart_data"),
            "generated_image_path": normalized.get("generated_image_path"),
        }

        slides.append(slide)



    return slides



def _existing_image_path(raw: Any) -> str | None:

    if not raw:

        return None

    path = Path(str(raw))

    if path.exists() and path.is_file():

        return str(path)

    return None





def _ensure_slide_images(project: Any, project_id: str, pages: list[Any], slides: list[dict[str, Any]], task_id: str) -> None:

    if not image_generator.enabled():

        return



    total = max(1, len(slides))

    project_title = str(project["title"] or "")

    style = _normalize_style(project["style"])



    for idx, slide in enumerate(slides):

        page_index = idx + 1

        if str(slide.get("slide_type") or "").lower() == "title":

            continue



        existing = _existing_image_path(slide.get("generated_image_path"))

        if existing:

            slide["generated_image_path"] = existing

            continue



        generated_path = image_generator.generate_for_slide(

            project_id=project_id,

            page_index=page_index,

            topic=project_title,

            title=str(slide.get("title") or f"第{page_index}页"),

            bullets=[str(x) for x in list(slide.get("bullets") or [])],

            notes=str(slide.get("notes") or ""),

            style=style,

        )

        if not generated_path:

            continue



        slide["generated_image_path"] = generated_path



        raw_desc = _safe_load_json(pages[idx]["description_content"], {})

        if isinstance(raw_desc, dict):

            raw_desc["generated_image_path"] = generated_path

            update_page(

                str(pages[idx]["page_id"]),

                {

                    "description_content": json.dumps(raw_desc, ensure_ascii=False),

                    "updated_at": utc_now_iso(),

                },

            )



        update_task(

            task_id,

            {

                "progress_json": make_progress(total, min(total, page_index), 0, "generating_images"),

            },

        )







def _is_cover_or_toc_title(title: str) -> bool:
    t = _cleanup_text(str(title or ""), max_len=120).lower()
    if not t:
        return False
    return any(
        k in t
        for k in (
            "\u5c01\u9762",
            "\u6807\u9898\u9875",
            "cover",
            "title",
            "\u76ee\u5f55",
            "\u8bae\u7a0b",
            "agenda",
            "contents",
            "toc",
        )
    )


def _derive_toc_items(outline_titles: list[str], slides: list[dict[str, Any]]) -> list[str]:
    # Prefer final slide titles so TOC exactly matches rendered body pages.
    items: list[str] = []
    for slide in slides:
        slide_type = str(slide.get("slide_type") or "").lower()
        title = _cleanup_text(str(slide.get("title") or ""), max_len=120)
        if not title:
            continue
        if slide_type in {"title", "toc"}:
            continue
        if _is_cover_or_toc_title(title):
            continue
        items.append(title)

    if items:
        return items

    # Fallback: derive from outline titles after removing cover/toc pages.
    return [
        _cleanup_text(str(x), max_len=120)
        for x in outline_titles
        if _cleanup_text(str(x), max_len=120) and not _is_cover_or_toc_title(str(x))
    ]


def _derive_export_topic(raw_title: str, material_text: str, outline_titles: list[str]) -> str:
    raw = _cleanup_text(str(raw_title or ""), max_len=200)
    mat = _cleanup_text(str(material_text or ""), max_len=300)

    core = re.sub(r"(?i)\b(pptx?|slides?|deck)\b", " ", raw)
    core = re.sub(
        r"(\u8bf7|\u5e2e\u6211|\u7ed9\u6211|\u7ed9\u51fa|\u751f\u6210|\u5236\u4f5c|\u505a\u4e00\u4efd|\u505a\u4e2a|\u4e00\u4e2a|\u4e00\u4efd|\u5173\u4e8e|\u4e3b\u9898|\u6c47\u62a5|\u6f14\u793a\u6587\u7a3f|\u6f14\u793a)",
        " ",
        core,
    )
    core = re.sub(r"\s+", " ", core).strip(" -:\uff1a")

    # Preserve/normalize acronyms such as rag -> RAG.
    core = re.sub(r"[A-Za-z]{2,8}", lambda m: m.group(0).upper(), core)

    toc_like = [
        _cleanup_text(x, max_len=80)
        for x in outline_titles
        if _cleanup_text(x, max_len=80) and not _is_cover_or_toc_title(x)
    ]
    focus = "\u3001".join(toc_like[:2])

    suffix = "\u7b54\u8fa9\u6c47\u62a5" if ("\u7b54\u8fa9" in raw or "\u7b54\u8fa9" in mat) else "\u4e13\u9898\u6c47\u62a5"

    if core and focus:
        if core.lower() in focus.lower():
            title = f"{focus}{suffix}"
        else:
            title = f"{core}\uff1a{focus}{suffix}"
    elif core:
        title = core
        if not any(k in title for k in ("\u6c47\u62a5", "\u7b54\u8fa9", "\u62a5\u544a")):
            title = f"{title}{suffix}"
    elif focus:
        title = f"{focus}{suffix}"
    else:
        title = "\u9879\u76ee\u4e13\u9898\u6c47\u62a5"

    title = re.sub(r"\s+", " ", title).strip()
    if len(title) > 42:
        title = title[:42].rstrip("\uff0c\u3002\uff1b\uff1a\u3001 ")

    return title or "\u9879\u76ee\u4e13\u9898\u6c47\u62a5"


def _derive_export_subtitle(style: str, page_count: int) -> str:
    # Keep cover clean by default; user-provided subtitle should be passed explicitly.
    return ""



def _derive_theme_seed(topic: str, style: str, project_id: str, toc_items: list[str]) -> str:
    sample = "|".join(toc_items[:3])
    return f"{topic}|{style}|{project_id}|{sample}"


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

            "progress_json": make_progress(total, 0, 0, "building_slides"),

        },

    )



    try:

        slides = _collect_project_slides(project, pages)

        _ensure_slide_images(project, project_id, pages, slides, task_id)

        update_task(

            task_id,

            {

                "progress_json": make_progress(total, total, 0, "exporting_pptx"),

            },

        )



        ts = datetime.now().strftime("%Y%m%d_%H%M%S")

        filename = f"{project_id}_{ts}.pptx"

        out_path = settings.export_dir / filename



        outline = [str(_safe_load_json(page["outline_content"], {}).get("title") or "") for page in pages]
        toc_items = _derive_toc_items(outline, slides)
        export_topic = _derive_export_topic(str(project["title"] or ""), str(project["material_text"] or ""), outline)
        style = _normalize_style(project["style"])
        export_subtitle = _derive_export_subtitle(style, len(outline))
        theme_seed = _derive_theme_seed(export_topic, style, project_id, toc_items)

        exported = export_slides_to_pptx(
            slides,
            out_path,
            str(project["template_id"] or "no_template"),
            export_topic,
            outline,
            subtitle=export_subtitle,
            toc_items=toc_items,
            style=style,
            theme_seed=theme_seed,
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

                "result_json": json.dumps({"pptx_url": pptx_url}, ensure_ascii=False),

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





def _slide_payload_to_description_text(payload: dict[str, Any]) -> str:
    title = str(payload.get("title") or "Untitled")
    bullets = [str(x) for x in list(payload.get("bullets") or [])]
    notes = str(payload.get("notes") or "")

    lines = [f"Page Title: {title}", "", "Page Text:"]
    for item in bullets[:5]:
        cleaned = _cleanup_text(item, max_len=180)
        if cleaned:
            lines.append(f"- {cleaned}")

    if notes:
        lines.extend(["", "Notes:", notes])

    return "\n".join(lines)

def _rewrite_requirement(action: str) -> str:

    mapping = {

        "concise": "请将所有页面内容精简为更短的表达，保留关键信息和结论。",

        "management": "请把所有页面改写成管理层汇报口径，强调结果、风险和决策建议。",

        "technical": "请把所有页面改写成技术汇报口径，强调现状、细节和实施计划。",

    }

    return mapping.get(action, "请优化页面描述表达。")





def rewrite_project(project_id: str, action: str) -> str:

    project = get_project(project_id)

    if not project:

        raise ValueError("project not found")



    pages = list_pages(project_id)

    if not pages:

        raise ValueError("no pages")



    outline = _outline_pages_from_db(pages)



    current_descriptions: list[dict[str, Any]] = []

    for idx, page in enumerate(pages):

        desc = _safe_load_json(page["description_content"], {})

        if isinstance(desc, dict) and desc.get("text"):

            raw_text = str(desc.get("text"))

        elif isinstance(desc, dict):

            raw_text = _slide_payload_to_description_text(desc)

        else:

            raw_text = ""



        current_descriptions.append(

            {

                "index": idx,

                "title": outline[idx]["title"],

                "description_content": {"text": raw_text},

            }

        )



    context = _context_for_project(project)

    user_requirement = _rewrite_requirement(action)



    refined = banana_ai.refine_descriptions(

        current_descriptions=current_descriptions,

        user_requirement=user_requirement,

        project_context=context,

        outline=outline,

        previous_requirements=None,

        language="zh",

    )



    if len(refined) < len(pages):

        for idx in range(len(refined), len(pages)):

            refined.append(current_descriptions[idx]["description_content"]["text"])

    elif len(refined) > len(pages):

        refined = refined[: len(pages)]

    rewritten_slides: list[dict[str, Any]] = []

    total = len(pages)

    project_title = str(project["title"] or "")

    image_style = action if action in {"management", "technical"} else _normalize_style(project["style"])



    for idx, refined_text in enumerate(refined):

        page_id = str(pages[idx]["page_id"])

        payload = _description_to_slide_payload(

            description_text=str(refined_text),

            page_outline=outline[idx],

            page_index=idx + 1,

            total_pages=total,

            extra_fields=None,

        )

        payload["page"] = idx + 1



        if image_generator.enabled() and str(payload.get("slide_type") or "").lower() != "title":

            generated_path = image_generator.generate_for_slide(

                project_id=project_id,

                page_index=idx + 1,

                topic=project_title,

                title=str(payload.get("title") or f"第{idx + 1}页"),

                bullets=[str(x) for x in list(payload.get("bullets") or [])],

                notes=str(payload.get("notes") or ""),

                style=image_style,

            )

            if generated_path:

                payload["generated_image_path"] = generated_path



        rewritten_slides.append(payload)



        update_page(

            page_id,

            {

                "description_content": json.dumps(payload, ensure_ascii=False),

                "status": "DESCRIPTION_GENERATED",

                "updated_at": utc_now_iso(),

            },

        )

    style = _normalize_style(project["style"])

    if action in {"management", "technical"}:

        style = action



    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    filename = f"{project_id}_{ts}.pptx"

    out_path = settings.export_dir / filename

    outline_titles = [item["title"] for item in outline]
    toc_items = _derive_toc_items(outline_titles, rewritten_slides)
    export_topic = _derive_export_topic(str(project["title"] or ""), str(project["material_text"] or ""), outline_titles)
    export_subtitle = _derive_export_subtitle(style, len(outline_titles))
    theme_seed = _derive_theme_seed(export_topic, style, project_id, toc_items)

    exported = export_slides_to_pptx(
        rewritten_slides,
        out_path,
        str(project["template_id"] or "no_template"),
        export_topic,
        outline_titles,
        subtitle=export_subtitle,
        toc_items=toc_items,
        style=style,
        theme_seed=theme_seed,
    )



    pptx_url = f"/exports/{exported}"

    update_project(

        project_id,

        {

            "style": style,

            "status": "COMPLETED",

            "pptx_url": pptx_url,

            "updated_at": utc_now_iso(),

        },

    )

    return pptx_url





















