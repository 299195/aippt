from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Dict, List, Tuple

from app.services.template_catalog import resolve_template_assets

_PROJECT_ROOT = Path(__file__).resolve().parents[3]
_PPTX_GENERATOR_DIR = Path(__file__).resolve().parents[2] / "pptx_generator"
_PPTX_GENERATOR_SCRIPT = _PPTX_GENERATOR_DIR / "generate_deck.js"

_AIPPTX_DIR = _PROJECT_ROOT / "third_party" / "ai-to-pptx-backend"
_AIPPTX_TEMPLATE_DIR = _AIPPTX_DIR / "json"
_AIPPTX_BRIDGE_SCRIPT = _AIPPTX_DIR / "local_generate_pptx.php"

_AIPPTX_TEMPLATE_NAMES: dict[str, str] = {
    "0": "\u8bfe\u7a0b\u5b66\u4e60\u6c47\u62a5",
    "1": "\u8bfb\u4e66\u5206\u4eab\u6f14\u793a",
    "2": "\u84dd\u8272\u901a\u7528\u5546\u52a1",
    "3": "\u84dd\u8272\u5de5\u4f5c\u6c47\u62a5\u603b\u7ed3",
    "a2p_0": "\u8bfe\u7a0b\u5b66\u4e60\u6c47\u62a5",
    "a2p_1": "\u8bfb\u4e66\u5206\u4eab\u6f14\u793a",
    "a2p_2": "\u84dd\u8272\u901a\u7528\u5546\u52a1",
    "a2p_3": "\u84dd\u8272\u5de5\u4f5c\u6c47\u62a5\u603b\u7ed3",
    "course_learning_report": "\u8bfe\u7a0b\u5b66\u4e60\u6c47\u62a5",
    "book_share_demo": "\u8bfb\u4e66\u5206\u4eab\u6f14\u793a",
    "blue_business": "\u84dd\u8272\u901a\u7528\u5546\u52a1",
    "blue_work_summary": "\u84dd\u8272\u5de5\u4f5c\u6c47\u62a5\u603b\u7ed3",
    "no_template": "\u84dd\u8272\u901a\u7528\u5546\u52a1",
    "executive_clean": "\u84dd\u8272\u901a\u7528\u5546\u52a1",
}
_AIPPTX_DEFAULT_TEMPLATE = "\u84dd\u8272\u901a\u7528\u5546\u52a1"


def _is_toc_like_title(title: str) -> bool:
    low = str(title or "").lower()
    return any(k in low for k in ("鐩綍", "璁▼", "agenda", "contents", "table of contents", "toc"))


def _content_slides(slides: List[Dict]) -> List[Dict]:
    out: List[Dict] = []
    for item in slides:
        title = str(item.get("title") or "")
        slide_type = str(item.get("slide_type") or "").lower()
        if slide_type in {"title", "toc"}:
            continue
        low = title.lower()
        if any(k in low for k in ("cover", "灏侀潰", "agenda")) or _is_toc_like_title(title):
            continue
        out.append(item)
    return out


def _normalize_md_text(value: str, max_len: int = 96) -> str:
    text = str(value or "").replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("```", " ")
    text = text.replace("#", " ")
    text = re.sub(r"^\s*[-*]+\s*", "", text)
    text = re.sub(r"^\s*\d+(?:\.\d+){0,3}\s*", "", text)
    text = re.sub(r"\s+", " ", text).strip(" -:\t\n")
    if max_len > 0 and len(text) > max_len:
        text = text[:max_len].rstrip()
    return text


def _default_topic(topic: str, body_slides: List[Dict]) -> str:
    normalized_topic = _normalize_md_text(topic, max_len=64)
    if normalized_topic:
        return normalized_topic
    if body_slides:
        first = _normalize_md_text(str(body_slides[0].get("title") or ""), max_len=64)
        if first:
            return first
    return "椤圭洰姹囨姤"


def _default_toc_items(body_slides: List[Dict], outline: List[str] | None) -> List[str]:
    from_body = [_normalize_md_text(str(item.get("title") or ""), max_len=48) for item in body_slides]
    from_body = [x for x in from_body if x]
    if from_body:
        return from_body
    if outline:
        normalized = [_normalize_md_text(str(x), max_len=48) for x in outline if str(x).strip()]
        return [x for x in normalized if x and not _is_toc_like_title(x)]
    return []


def _dedupe_keep_order(values: List[str], limit: int) -> List[str]:
    out: List[str] = []
    seen: set[str] = set()
    for raw in values:
        text = _normalize_md_text(raw, max_len=140)
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


def _split_sentences(text: str, max_len: int = 120) -> List[str]:
    chunks = re.split(r"[銆傦紒锛??锛?]\s*|\n+", str(text or ""))
    parts: List[str] = []
    for raw in chunks:
        val = _normalize_md_text(raw, max_len=max_len)
        if val and len(val) >= 6:
            parts.append(val)
    return parts


def _is_redundant_pair(title: str, detail: str) -> bool:
    t = _normalize_md_text(title, max_len=120).lower()
    d = _normalize_md_text(detail, max_len=220).lower()
    if not t or not d:
        return False
    if t == d:
        return True
    return d.startswith(t) and (len(d) - len(t) <= 12)


def _section_pairs(slide: Dict) -> List[Tuple[str, str]]:
    title = _normalize_md_text(str(slide.get("title") or ""), max_len=48) or "Key Point"
    bullets = [str(x) for x in list(slide.get("bullets") or [])]
    detail_points = [str(x) for x in list(slide.get("detail_points") or [])]
    text_blocks = [str(x) for x in list(slide.get("text_blocks") or [])]
    notes = str(slide.get("notes") or "")

    headings = _dedupe_keep_order([*bullets, *detail_points, *text_blocks], limit=6)
    if not headings:
        headings = [
            f"{title} Core Topic",
            f"{title} Key Insight",
            f"{title} Action Suggestion",
        ]

    detail_candidates = _dedupe_keep_order(
        [
            *detail_points,
            *text_blocks,
            *_split_sentences(notes, max_len=120),
        ],
        limit=12,
    )

    pairs: List[Tuple[str, str]] = []
    for head in headings:
        if len(pairs) >= 4:
            break
        detail = ""
        for candidate in detail_candidates:
            if not _is_redundant_pair(head, candidate):
                detail = candidate
                break
        if detail:
            detail_candidates = [x for x in detail_candidates if x != detail]
        else:
            detail = f"Explain {head} with context, impact, and actionable suggestions."
        pairs.append((head, _normalize_md_text(detail, max_len=140)))

    if len(pairs) < 2:
        pairs.append(
            (
                f"{title} Background and Current State",
                "Describe the current status, key challenges, and why this topic matters now.",
            )
        )
    if len(pairs) < 3:
        pairs.append(
            (
                f"{title} Plan and Execution",
                "Provide key actions, implementation path, milestones, and measurable outcomes.",
            )
        )

    return pairs[:4]


def _preferred_chapter_titles(
    body_slides: List[Dict],
    outline: List[str] | None,
    toc_items: List[str] | None,
) -> List[str]:
    from_toc = _dedupe_keep_order([str(x) for x in (toc_items or [])], limit=64)
    if from_toc:
        return from_toc

    from_outline = [
        _normalize_md_text(str(x), max_len=48)
        for x in (outline or [])
        if str(x).strip()
    ]
    from_outline = [x for x in from_outline if x and not _is_toc_like_title(x)]
    from_outline = _dedupe_keep_order(from_outline, limit=64)
    if from_outline:
        return from_outline

    from_body = [_normalize_md_text(str(item.get("title") or ""), max_len=48) for item in body_slides]
    from_body = [x for x in from_body if x]
    return _dedupe_keep_order(from_body, limit=64)


def _build_chapter_groups(
    body_slides: List[Dict],
    outline: List[str] | None,
    toc_items: List[str] | None,
) -> List[Tuple[str, List[Dict]]]:
    if not body_slides:
        return []

    chapter_titles = _preferred_chapter_titles(body_slides, outline, toc_items)
    if not chapter_titles:
        chapter_titles = [_normalize_md_text(str(body_slides[0].get("title") or ""), max_len=48) or "鏍稿績绔犺妭"]

    chapter_count = max(1, min(len(chapter_titles), len(body_slides)))
    chapter_titles = chapter_titles[:chapter_count]

    total = len(body_slides)
    base = total // chapter_count
    remainder = total % chapter_count

    groups: List[Tuple[str, List[Dict]]] = []
    cursor = 0
    for idx in range(chapter_count):
        take = base + (1 if idx < remainder else 0)
        chunk = body_slides[cursor : cursor + take]
        cursor += take
        if not chunk:
            continue
        groups.append((chapter_titles[idx], chunk))
    return groups


def _build_outline_markdown(
    topic: str,
    body_slides: List[Dict],
    outline: List[str] | None,
    toc_items: List[str] | None,
) -> str:
    chapters = _build_chapter_groups(body_slides, outline, toc_items)
    lines: List[str] = [f"# {topic}"]
    for chapter_idx, (chapter_title, chapter_slides) in enumerate(chapters, start=1):
        lines.append(f"## {chapter_idx}. {chapter_title}")
        for section_idx, slide in enumerate(chapter_slides, start=1):
            section_title = _normalize_md_text(str(slide.get("title") or ""), max_len=48) or f"灏忚妭{chapter_idx}-{section_idx}"
            lines.append(f"### {chapter_idx}.{section_idx} {section_title}")
            for point_idx, (point_title, _) in enumerate(_section_pairs(slide), start=1):
                lines.append(f"{chapter_idx}.{section_idx}.{point_idx} {point_title}")
    return "\n".join(lines)


def _build_content_markdown(
    topic: str,
    body_slides: List[Dict],
    outline: List[str] | None,
    toc_items: List[str] | None,
) -> str:
    chapters = _build_chapter_groups(body_slides, outline, toc_items)
    lines: List[str] = [f"# {topic}"]
    for chapter_idx, (chapter_title, chapter_slides) in enumerate(chapters, start=1):
        lines.append(f"## {chapter_idx}. {chapter_title}")
        for section_idx, slide in enumerate(chapter_slides, start=1):
            section_title = _normalize_md_text(str(slide.get("title") or ""), max_len=48) or f"灏忚妭{chapter_idx}-{section_idx}"
            lines.append(f"### {chapter_idx}.{section_idx} {section_title}")
            for point_idx, (point_title, point_detail) in enumerate(_section_pairs(slide), start=1):
                lines.append(f"{chapter_idx}.{section_idx}.{point_idx} {point_title}")
                lines.append(point_detail)
    return "\n".join(lines)


def _resolve_ai_to_pptx_template_json(template_id: str) -> Path:
    key = str(template_id or "").strip().lower()
    template_name = _AIPPTX_TEMPLATE_NAMES.get(key, "")
    if not template_name:
        raw = str(template_id or "").strip()
        if (_AIPPTX_TEMPLATE_DIR / f"{raw}.json").exists():
            template_name = raw
        else:
            template_name = _AIPPTX_DEFAULT_TEMPLATE

    template_path = _AIPPTX_TEMPLATE_DIR / f"{template_name}.json"
    if not template_path.exists():
        raise RuntimeError(f"Ai-To-PPTX template not found: {template_path}")
    return template_path


def _discover_bundled_php() -> str | None:
    tools_dir = _PROJECT_ROOT / "backend" / "tools" / "php"
    if not tools_dir.exists():
        return None

    direct = tools_dir / "php.exe"
    if direct.is_file():
        return str(direct)

    candidates = sorted(tools_dir.glob("php-*"), reverse=True)
    for folder in candidates:
        candidate = folder / "php.exe"
        if candidate.is_file():
            return str(candidate)
    return None


def _php_has_zip_extension(php_bin: str) -> bool:
    try:
        completed = subprocess.run(
            [php_bin, "-m"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
            check=False,
        )
    except Exception:
        return False

    if completed.returncode != 0:
        return False

    modules = {line.strip().lower() for line in (completed.stdout or "").splitlines() if line.strip()}
    return "zip" in modules


def _find_php_bin() -> str:
    candidates: list[str] = []

    configured = str(os.getenv("AIPPT_PHP_BIN", "") or "").strip()
    if configured:
        configured_path = Path(configured)
        if configured_path.is_file():
            candidates.append(str(configured_path))
        resolved_configured = shutil.which(configured)
        if resolved_configured:
            candidates.append(resolved_configured)

    resolved_default = shutil.which("php")
    if resolved_default:
        candidates.append(resolved_default)

    bundled = _discover_bundled_php()
    if bundled:
        candidates.append(bundled)

    unique_candidates = list(dict.fromkeys(candidates))
    for php_bin in unique_candidates:
        if _php_has_zip_extension(php_bin):
            return php_bin

    if unique_candidates:
        checked = "; ".join(unique_candidates)
        raise RuntimeError(
            "No usable PHP runtime with zip extension found. "
            f"Checked: {checked}. "
            "Install PHP>=7.4 with zip enabled or set AIPPT_PHP_BIN to a PHP binary that has zip."
        )

    raise RuntimeError(
        "PHP runtime not found. Install PHP>=7.4 with zip extension, "
        "or set AIPPT_PHP_BIN to php executable path."
    )


def _ensure_ai_to_pptx_ready() -> None:
    if not _AIPPTX_DIR.exists():
        raise RuntimeError(f"Ai-To-PPTX backend directory missing: {_AIPPTX_DIR}")
    if not _AIPPTX_BRIDGE_SCRIPT.exists():
        raise RuntimeError(f"Ai-To-PPTX bridge script missing: {_AIPPTX_BRIDGE_SCRIPT}")
    if not _AIPPTX_TEMPLATE_DIR.exists():
        raise RuntimeError(f"Ai-To-PPTX template directory missing: {_AIPPTX_TEMPLATE_DIR}")


def _export_with_ai_to_pptx(
    slides: List[Dict],
    out_path: Path,
    template_id: str,
    topic: str,
    outline: List[str] | None,
    subtitle: str,
    toc_items: List[str] | None,
) -> str:
    _ensure_ai_to_pptx_ready()
    php_bin = _find_php_bin()

    body_slides = _content_slides(slides)
    if not body_slides:
        body_slides = [x for x in slides if _normalize_md_text(str(x.get("title") or ""), max_len=80)]
    if not body_slides:
        raise RuntimeError("no slide content available for Ai-To-PPTX export")

    topic_text = _default_topic(topic, body_slides)
    outline_markdown = _build_outline_markdown(topic_text, body_slides, outline, toc_items)
    content_markdown = _build_content_markdown(topic_text, body_slides, outline, toc_items)
    template_path = _resolve_ai_to_pptx_template_json(template_id)
    author_text = _normalize_md_text(subtitle or "", max_len=32)
    last_page_text = _normalize_md_text(os.getenv("AIPPTX_LAST_PAGE_TEXT", "鎰熻阿鑱嗗惉"), max_len=32) or "鎰熻阿鑱嗗惉"

    target_path = out_path if out_path.is_absolute() else (Path.cwd() / out_path)
    target_path.parent.mkdir(parents=True, exist_ok=True)

    with NamedTemporaryFile("w", suffix=".md", delete=False, encoding="utf-8") as outline_tmp:
        outline_tmp.write(outline_markdown)
        outline_path = Path(outline_tmp.name)
    with NamedTemporaryFile("w", suffix=".md", delete=False, encoding="utf-8") as content_tmp:
        content_tmp.write(content_markdown)
        content_path = Path(content_tmp.name)

    try:
        cmd = [
            php_bin,
            str(_AIPPTX_BRIDGE_SCRIPT),
            str(template_path),
            str(outline_path),
            str(content_path),
            str(target_path),
            author_text,
            last_page_text,
        ]
        completed = subprocess.run(
            cmd,
            cwd=str(_AIPPTX_DIR),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
            check=False,
        )
    finally:
        outline_path.unlink(missing_ok=True)
        content_path.unlink(missing_ok=True)

    if completed.returncode != 0:
        stderr = (completed.stderr or "").strip()
        stdout = (completed.stdout or "").strip()
        detail = stderr or stdout or f"exit code {completed.returncode}"
        if "did not produce output" in detail.lower():
            detail = f"{detail} (php: {php_bin})"
        raise RuntimeError(f"Ai-To-PPTX export failed: {detail}")

    if not target_path.exists():
        raise RuntimeError("Ai-To-PPTX exporter did not produce output file")

    return target_path.name


def _ensure_pptx_generator_ready() -> None:
    if not _PPTX_GENERATOR_SCRIPT.exists():
        raise RuntimeError(f"pptx-generator script missing: {_PPTX_GENERATOR_SCRIPT}")

    deps_marker = _PPTX_GENERATOR_DIR / "node_modules" / "pptxgenjs"
    if not deps_marker.exists():
        raise RuntimeError(
            "pptx-generator dependencies are not installed. "
            "Run: cd backend/pptx_generator && npm install"
        )


def _export_with_node_generator(
    slides: List[Dict],
    out_path: Path,
    template_id: str,
    topic: str,
    outline: List[str] | None,
    subtitle: str,
    toc_items: List[str] | None,
    style: str,
    theme_seed: str,
) -> str:
    _ensure_pptx_generator_ready()

    assets = resolve_template_assets(template_id)
    template_pptx_path = assets.get("pptx_path")

    body_slides = _content_slides(slides)
    effective_toc = [str(x).strip() for x in (toc_items or []) if str(x).strip()]
    if not effective_toc:
        effective_toc = _default_toc_items(body_slides, outline)

    payload = {
        "topic": _default_topic(topic, body_slides),
        "subtitle": subtitle or "",
        "templateId": template_id,
        "style": style,
        "themeSeed": theme_seed or "",
        "tocItems": effective_toc,
        "outline": outline[:] if outline else [str(item.get("title") or "") for item in body_slides],
        "slides": slides,
        "templatePptxPath": str(template_pptx_path) if template_pptx_path and template_pptx_path.exists() else None,
    }

    target_path = out_path if out_path.is_absolute() else (Path.cwd() / out_path)
    target_path.parent.mkdir(parents=True, exist_ok=True)

    with NamedTemporaryFile("w", suffix=".json", delete=False, encoding="utf-8") as tmp:
        json.dump(payload, tmp, ensure_ascii=False)
        payload_path = Path(tmp.name)

    node_bin = os.getenv("PPTX_NODE_BIN", "node")
    cmd = [
        node_bin,
        str(_PPTX_GENERATOR_SCRIPT),
        "--input",
        str(payload_path),
        "--output",
        str(target_path),
    ]

    try:
        completed = subprocess.run(
            cmd,
            cwd=str(_PPTX_GENERATOR_DIR),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
            check=False,
        )
    finally:
        payload_path.unlink(missing_ok=True)

    if completed.returncode != 0:
        stderr = (completed.stderr or "").strip()
        stdout = (completed.stdout or "").strip()
        detail = stderr or stdout or f"exit code {completed.returncode}"
        raise RuntimeError(f"pptx-generator failed: {detail}")

    if not target_path.exists():
        raise RuntimeError("pptx-generator did not produce output file")

    return target_path.name


def export_slides_to_pptx(
    slides: List[Dict],
    out_path: Path,
    template_id: str = "no_template",
    topic: str = "",
    outline: List[str] | None = None,
    *,
    subtitle: str = "",
    toc_items: List[str] | None = None,
    style: str = "management",
    theme_seed: str = "",
) -> str:
    # ai_to_pptx: strict Ai-To-PPTX exporter (same engine as SmartSchoolAI backend)
    # auto: try Ai-To-PPTX first, fallback to legacy node exporter
    # legacy: force existing node exporter
    engine = str(os.getenv("AIPPT_EXPORT_ENGINE", "ai_to_pptx") or "ai_to_pptx").strip().lower()
    if engine not in {"ai_to_pptx", "auto", "legacy"}:
        engine = "ai_to_pptx"

    ai_error: Exception | None = None
    if engine in {"ai_to_pptx", "auto"}:
        try:
            return _export_with_ai_to_pptx(
                slides=slides,
                out_path=out_path,
                template_id=template_id,
                topic=topic,
                outline=outline,
                subtitle=subtitle,
                toc_items=toc_items,
            )
        except Exception as exc:  # noqa: PERF203
            ai_error = exc
            if engine == "ai_to_pptx":
                raise RuntimeError(
                    "Ai-To-PPTX export failed. Install PHP>=7.4 with zip extension "
                    "and set AIPPT_PHP_BIN if needed. "
                    f"Original error: {exc}"
                ) from exc

    try:
        return _export_with_node_generator(
            slides=slides,
            out_path=out_path,
            template_id=template_id,
            topic=topic,
            outline=outline,
            subtitle=subtitle,
            toc_items=toc_items,
            style=style,
            theme_seed=theme_seed,
        )
    except Exception as node_exc:
        if ai_error is None:
            raise
        raise RuntimeError(
            f"Both exporters failed. Ai-To-PPTX error: {ai_error}; "
            f"legacy exporter error: {node_exc}"
        ) from node_exc

