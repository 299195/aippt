"""Banana-style AI service adapted for the FastAPI backend."""

from __future__ import annotations

import base64
import json
import logging
import random
import re
import ssl
from dataclasses import dataclass, field
from datetime import datetime
from io import BytesIO
from textwrap import dedent
from typing import Any, Dict, Iterator, List, Optional
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from app.services.banana_prompts import (
    get_all_descriptions_stream_prompt,
    get_description_split_prompt,
    get_description_to_outline_prompt,
    get_description_to_outline_prompt_markdown,
    get_descriptions_refinement_prompt,
    get_image_edit_prompt,
    get_image_generation_prompt,
    get_layout_caption_prompt,
    get_outline_generation_prompt,
    get_outline_generation_prompt_markdown,
    get_outline_parsing_prompt,
    get_outline_parsing_prompt_markdown,
    get_page_description_prompt,
    get_ppt_page_content_extraction_prompt,
    get_style_extraction_prompt,
)
from PIL import Image

from app.config import settings
from app.services.model_client import ModelClient

logger = logging.getLogger(__name__)


def build_idea_prompt(title: str, style: str, material_text: str) -> str:
    title_s = (title or "").strip()
    style_s = (style or "management").strip()
    material_s = (material_text or "").strip()

    parts = [f"主题：{title_s}", f"汇报风格：{style_s}"]
    if material_s:
        parts.append("补充材料：")
        parts.append(material_s)
    return "\n".join(parts).strip()


@dataclass
class BananaProjectContext:
    idea_prompt: str | None = None
    outline_text: str | None = None
    description_text: str | None = None
    creation_type: str = "idea"
    outline_requirements: str | None = None
    description_requirements: str | None = None
    reference_files_content: list[dict[str, str]] = field(default_factory=list)


class BananaAIService:
    def __init__(self, use_mock: bool = True) -> None:
        self.use_mock = use_mock
        self.client = ModelClient()
        if not self.use_mock and not self.client.enabled():
            raise RuntimeError("model config incomplete")

    @staticmethod
    def _clean_json_text(text: str) -> str:
        payload = (text or "").strip()
        if payload.startswith("```"):
            lines = payload.splitlines()
            if lines and lines[0].startswith("```"):
                lines = lines[1:]
            if lines and lines[-1].strip() == "```":
                lines = lines[:-1]
            payload = "\n".join(lines).strip()
        return payload

    @classmethod
    def _parse_json_payload(cls, text: str) -> Any:
        payload = cls._clean_json_text(text)
        candidates = [payload]

        start_obj = payload.find("{")
        end_obj = payload.rfind("}")
        if start_obj >= 0 and end_obj > start_obj:
            candidates.append(payload[start_obj : end_obj + 1])

        start_arr = payload.find("[")
        end_arr = payload.rfind("]")
        if start_arr >= 0 and end_arr > start_arr:
            candidates.append(payload[start_arr : end_arr + 1])

        last_error: Exception | None = None
        for candidate in candidates:
            try:
                return json.loads(candidate)
            except Exception as exc:  # noqa: PERF203
                last_error = exc

        raise ValueError(f"json parse failed: {last_error}")

    def _generate_text(self, prompt: str, temperature: float = 0.5) -> str:
        if self.use_mock:
            raise RuntimeError("mock mode should use mock branch")
        return self.client.chat_text(
            system_prompt="You are a helpful assistant for PPT generation.",
            user_prompt=prompt,
            temperature=temperature,
        )

    def _generate_text_stream(self, prompt: str, temperature: float = 0.5) -> Iterator[str]:
        if self.use_mock:
            raise RuntimeError("mock mode should use mock branch")
        yield from self.client.chat_text_stream(
            system_prompt="You are a helpful assistant for PPT generation.",
            user_prompt=prompt,
            temperature=temperature,
        )

    def _generate_json(self, prompt: str, temperature: float = 0.5) -> Any:
        last_error: Exception | None = None
        for _ in range(3):
            try:
                text = self._generate_text(prompt, temperature=temperature)
                return self._parse_json_payload(text)
            except Exception as exc:
                last_error = exc
        raise RuntimeError(f"generate_json failed: {last_error}")

    @staticmethod
    def flatten_outline(outline: List[Dict]) -> List[Dict]:
        pages: list[dict[str, Any]] = []
        for item in outline or []:
            if isinstance(item, dict) and "part" in item and isinstance(item.get("pages"), list):
                for page in item.get("pages", []):
                    if not isinstance(page, dict):
                        continue
                    merged = dict(page)
                    merged["part"] = item.get("part")
                    pages.append(merged)
            elif isinstance(item, dict):
                pages.append(dict(item))
        return pages

    @staticmethod
    def _normalize_page(page: dict[str, Any], fallback_idx: int) -> dict[str, Any]:
        title = str(page.get("title") or f"第{fallback_idx}页").strip()
        points_raw = page.get("points")
        if isinstance(points_raw, list):
            points = [str(x).strip() for x in points_raw if str(x).strip()]
        else:
            points = []
        normalized = {
            "title": title,
            "points": points,
        }
        if page.get("part"):
            normalized["part"] = str(page.get("part"))
        return normalized

    @staticmethod
    def parse_markdown_outline(markdown: str) -> List[Dict]:
        pages: list[dict[str, Any]] = []
        current_part: str | None = None
        current_page: dict[str, Any] | None = None

        for line in (markdown or "").split("\n"):
            stripped = line.strip()
            if not stripped:
                continue
            if stripped.startswith("# ") and not stripped.startswith("## "):
                current_part = stripped[2:].strip()
                continue
            if stripped.startswith("## "):
                if current_page:
                    pages.append(current_page)
                current_page = {"title": stripped[3:].strip(), "points": []}
                if current_part:
                    current_page["part"] = current_part
                continue
            if stripped.startswith("- ") and current_page is not None:
                current_page["points"].append(stripped[2:].strip())

        if current_page:
            pages.append(current_page)

        return pages

    def generate_outline(self, project_context: BananaProjectContext, language: str | None = None) -> List[Dict]:
        if self.use_mock:
            return self._mock_generate_outline(project_context)

        # Primary path: strict JSON outline response.
        try:
            prompt = get_outline_generation_prompt(project_context, language)
            data = self._generate_json(prompt, temperature=0.5)
            if not isinstance(data, list):
                raise RuntimeError("outline output is not list")
            pages = self.flatten_outline(data)
            normalized = [self._normalize_page(page, idx + 1) for idx, page in enumerate(pages)]
            if normalized:
                return normalized
        except Exception:
            logger.warning("outline json path failed, fallback to markdown path", exc_info=True)

        # Fallback path: markdown outline parsing (more tolerant to model format drift).
        md_prompt = get_outline_generation_prompt_markdown(project_context, language)
        md_text = self._generate_text(md_prompt, temperature=0.4)
        md_pages = self.parse_markdown_outline(md_text)
        normalized = [self._normalize_page(page, idx + 1) for idx, page in enumerate(md_pages)]
        if normalized:
            return normalized
        raise RuntimeError("outline generation failed in both json and markdown paths")

    def generate_outline_stream(self, project_context: BananaProjectContext, language: str | None = None):
        if self.use_mock:
            creation_type = project_context.creation_type or "idea"
            if creation_type == "outline":
                pages = self._mock_parse_outline_text(project_context)
            else:
                pages = self._mock_generate_outline(project_context)
            for idx, page in enumerate(pages, start=1):
                yield self._normalize_page(page, idx)
            yield {"__stream_complete__": True}
            return

        creation_type = project_context.creation_type or "idea"
        if creation_type == "outline":
            prompt = get_outline_parsing_prompt_markdown(project_context, language)
        elif creation_type == "descriptions":
            prompt = get_description_to_outline_prompt_markdown(project_context, language)
        else:
            prompt = get_outline_generation_prompt_markdown(project_context, language)

        buffer = ""
        current_part: str | None = None
        current_page: dict[str, Any] | None = None
        stream_complete = False

        for chunk in self._generate_text_stream(prompt, temperature=0.4):
            buffer += chunk
            while "\n" in buffer:
                line, buffer = buffer.split("\n", 1)
                stripped = line.strip()
                if not stripped:
                    continue
                if stripped == "<!-- END -->":
                    stream_complete = True
                    continue
                if stripped.startswith("# ") and not stripped.startswith("## "):
                    current_part = stripped[2:].strip()
                    continue
                if stripped.startswith("## "):
                    if current_page:
                        yield current_page
                    current_page = {"title": stripped[3:].strip(), "points": []}
                    if current_part:
                        current_page["part"] = current_part
                    continue
                if stripped.startswith("- ") and current_page is not None:
                    current_page["points"].append(stripped[2:].strip())

        if buffer.strip():
            for line in buffer.split("\n"):
                stripped = line.strip()
                if not stripped:
                    continue
                if stripped == "<!-- END -->":
                    stream_complete = True
                    continue
                if stripped.startswith("# ") and not stripped.startswith("## "):
                    current_part = stripped[2:].strip()
                    continue
                if stripped.startswith("## "):
                    if current_page:
                        yield current_page
                    current_page = {"title": stripped[3:].strip(), "points": []}
                    if current_part:
                        current_page["part"] = current_part
                    continue
                if stripped.startswith("- ") and current_page is not None:
                    current_page["points"].append(stripped[2:].strip())

        if current_page:
            yield current_page

        yield {"__stream_complete__": stream_complete}
    def parse_outline_text(self, project_context: BananaProjectContext, language: str | None = None) -> List[Dict]:
        if self.use_mock:
            return self._mock_parse_outline_text(project_context)

        prompt = get_outline_parsing_prompt(project_context, language)
        data = self._generate_json(prompt, temperature=0.3)
        if not isinstance(data, list):
            raise RuntimeError("outline parse output is not list")
        pages = self.flatten_outline(data)
        return [self._normalize_page(page, idx + 1) for idx, page in enumerate(pages)]

    def parse_description_to_outline(self, project_context: BananaProjectContext, language: str = "zh") -> List[Dict]:
        if self.use_mock:
            return self._mock_generate_outline(project_context)

        prompt = get_description_to_outline_prompt(project_context, language)
        data = self._generate_json(prompt, temperature=0.4)
        if not isinstance(data, list):
            raise RuntimeError("description->outline output is not list")
        pages = self.flatten_outline(data)
        return [self._normalize_page(page, idx + 1) for idx, page in enumerate(pages)]

    @staticmethod
    def _get_extra_field_names() -> list[str]:
        return ["视觉元素", "视觉焦点", "排版布局", "演讲者备注"]

    @staticmethod
    def _build_extra_field_pattern(field_names: list[str]):
        if not field_names:
            return None
        escaped = "|".join(re.escape(name) for name in field_names)
        return re.compile(rf"^({escaped})[：:]\s*(.*)")

    @staticmethod
    def extract_image_urls_from_markdown(text: str) -> List[str]:
        if not text:
            return []
        pattern = r"!\[.*?\]\((.*?)\)"
        matches = re.findall(pattern, text)
        urls: list[str] = []
        for raw in matches:
            url = str(raw).strip()
            if url and (url.startswith("http://") or url.startswith("https://") or url.startswith("/files/")):
                urls.append(url)
        return urls

    @staticmethod
    def remove_markdown_images(text: str) -> str:
        if not text:
            return text

        def replace_image(match: re.Match) -> str:
            alt_text = match.group(1).strip()
            return alt_text if alt_text else ""

        cleaned_text = re.sub(r"!\[(.*?)\]\([^\)]+\)", replace_image, text)
        cleaned_text = re.sub(r"\n\s*\n\s*\n", "\n\n", cleaned_text)
        return cleaned_text

    @staticmethod
    def download_image_from_url(url: str) -> Optional[Image.Image]:
        try:
            req = Request(url=url, method="GET")
            with urlopen(req, timeout=30) as resp:
                raw = resp.read()
            image = Image.open(BytesIO(raw))
            image.load()
            return image
        except Exception:
            logger.warning("failed to download image from %s", url, exc_info=True)
            return None

    @staticmethod
    def _parse_extra_fields(text: str, field_names: list[str]) -> tuple[str, dict[str, str]]:
        if not field_names:
            return text, {}

        extra_fields: dict[str, str] = {}
        positions: list[tuple[int, int, str]] = []

        for name in field_names:
            match = re.search(rf"\n{re.escape(name)}[：:]\s*", text)
            if match:
                positions.append((match.start(), match.end(), name))

        if not positions:
            return text.strip(), {}

        positions.sort(key=lambda item: item[0])
        for i, (_, end, name) in enumerate(positions):
            next_start = positions[i + 1][0] if i + 1 < len(positions) else len(text)
            value = re.sub(r"<!--.*?-->", "", text[end:next_start], flags=re.S).strip()
            if value:
                extra_fields[name] = value

        cleaned_text = text[: positions[0][0]].strip()
        return cleaned_text, extra_fields

    def generate_page_description(
        self,
        project_context: BananaProjectContext,
        outline: List[Dict],
        page_outline: Dict,
        page_index: int,
        language: str = "zh",
        detail_level: str = "default",
    ) -> Dict:
        if self.use_mock:
            return self._mock_generate_page_description(project_context, outline, page_outline, page_index)

        extra_field_names = self._get_extra_field_names()
        part_info = f"\nThis page belongs to: {page_outline['part']}" if page_outline.get("part") else ""

        prompt = get_page_description_prompt(
            project_context=project_context,
            outline=outline,
            page_outline=page_outline,
            page_index=page_index,
            part_info=part_info,
            language=language,
            detail_level=detail_level,
            extra_fields=extra_field_names,
        )

        response_text = self._generate_text(prompt, temperature=0.65)
        text = dedent(response_text)
        description_text, extra_fields = self._parse_extra_fields(text, extra_field_names)

        result: dict[str, Any] = {"text": description_text}
        if extra_fields:
            result["extra_fields"] = extra_fields
        return result

    def generate_descriptions_stream(
        self,
        project_context: BananaProjectContext,
        outline: List[Dict],
        flat_pages: List[Dict],
        language: str = "zh",
        detail_level: str = "default",
    ):
        extra_field_names = self._get_extra_field_names()

        if self.use_mock:
            for idx, page in enumerate(flat_pages, start=1):
                generated = self._mock_generate_page_description(project_context, outline, page, idx)
                text = str(generated.get("text") or "")
                cleaned, extra_fields = self._parse_extra_fields(text, extra_field_names)
                payload: dict[str, Any] = {
                    "page_index": idx - 1,
                    "description_text": cleaned,
                }
                if extra_fields:
                    payload["extra_fields"] = extra_fields
                yield payload
            yield {"__stream_complete__": True}
            return

        prompt = get_all_descriptions_stream_prompt(
            project_context=project_context,
            outline=outline,
            flat_pages=flat_pages,
            language=language,
            detail_level=detail_level,
            extra_fields=extra_field_names,
        )

        field_pattern = self._build_extra_field_pattern(extra_field_names)
        buffer = ""
        page_index = -1
        current_lines: list[str] = []
        current_field: Optional[str] = None
        extra_fields: dict[str, str] = {}
        stream_complete = False

        def _build_page_result() -> dict[str, Any]:
            result: dict[str, Any] = {
                "page_index": page_index,
                "description_text": "\n".join(current_lines).strip(),
            }
            if extra_fields:
                result["extra_fields"] = dict(extra_fields)
            return result

        def _reset_page_state() -> None:
            nonlocal current_lines, current_field, extra_fields
            current_lines = []
            current_field = None
            extra_fields = {}

        def _process_line(line: str, stripped: str) -> str:
            nonlocal page_index, current_field, stream_complete

            if stripped == "<!-- BEGIN -->":
                if page_index < 0:
                    page_index = 0
                return "continue"

            if stripped == "<!-- END -->":
                stream_complete = True
                return "continue"

            if stripped == "<!-- PAGE_END -->":
                if page_index >= 0 and (current_lines or extra_fields):
                    return "yield_page"
                return "continue"

            if page_index < 0:
                return "continue"

            if field_pattern:
                field_match = field_pattern.match(stripped)
                if field_match:
                    field_name = field_match.group(1)
                    current_field = field_name
                    value = field_match.group(2).strip()
                    if value:
                        extra_fields[field_name] = value
                    return "continue"

            if not stripped:
                return "continue"

            if current_field:
                if current_field in extra_fields:
                    extra_fields[current_field] += "\n" + stripped
                else:
                    extra_fields[current_field] = stripped
            else:
                current_lines.append(line.rstrip())

            return "continue"

        for chunk in self._generate_text_stream(prompt, temperature=0.55):
            buffer += chunk
            while "\n" in buffer:
                line, buffer = buffer.split("\n", 1)
                action = _process_line(line, line.strip())
                if action == "yield_page":
                    yield _build_page_result()
                    _reset_page_state()
                    page_index += 1

        if buffer.strip():
            for line in buffer.split("\n"):
                action = _process_line(line, line.strip())
                if action == "yield_page":
                    yield _build_page_result()
                    _reset_page_state()
                    page_index += 1

        if page_index >= 0 and current_lines:
            yield _build_page_result()

        yield {"__stream_complete__": stream_complete}
    def parse_description_to_page_descriptions(
        self,
        project_context: BananaProjectContext,
        outline: List[Dict],
        language: str = "zh",
    ) -> List[str]:
        if self.use_mock:
            return self._mock_description_split(project_context, outline)

        prompt = get_description_split_prompt(project_context, outline, language)
        data = self._generate_json(prompt, temperature=0.4)
        if not isinstance(data, list):
            raise RuntimeError("description split output is not list")
        return [str(x) for x in data]

    def refine_descriptions(
        self,
        current_descriptions: List[Dict],
        user_requirement: str,
        project_context: BananaProjectContext,
        outline: List[Dict] | None = None,
        previous_requirements: Optional[List[str]] = None,
        language: str = "zh",
    ) -> List[str]:
        if self.use_mock:
            return self._mock_refine_descriptions(current_descriptions, user_requirement)

        prompt = get_descriptions_refinement_prompt(
            current_descriptions=current_descriptions,
            user_requirement=user_requirement,
            project_context=project_context,
            outline=outline,
            previous_requirements=previous_requirements,
            language=language,
        )
        data = self._generate_json(prompt, temperature=0.45)
        if not isinstance(data, list):
            raise RuntimeError("description refine output is not list")
        return [str(x) for x in data]

    def generate_outline_text(self, outline: List[Dict]) -> str:
        text_parts: list[str] = []
        for i, item in enumerate(outline, 1):
            if "part" in item and "pages" in item:
                text_parts.append(f"{i}. {item['part']}")
            else:
                text_parts.append(f"{i}. {item.get('title', 'Untitled')}")
        return "\n".join(text_parts)

    def generate_image_prompt(
        self,
        outline: List[Dict],
        page: Dict,
        page_desc: str,
        page_index: int,
        has_material_images: bool = False,
        extra_requirements: Optional[str] = None,
        language: str = "zh",
        has_template: bool = True,
        aspect_ratio: str = "16:9",
    ) -> str:
        outline_text = self.generate_outline_text(outline)
        current_section = page.get("part") or page.get("title", "Untitled")
        cleaned_page_desc = self.remove_markdown_images(page_desc)

        return get_image_generation_prompt(
            page_desc=cleaned_page_desc,
            outline_text=outline_text,
            current_section=str(current_section),
            has_material_images=has_material_images,
            extra_requirements=extra_requirements,
            language=language,
            has_template=has_template,
            page_index=page_index,
            aspect_ratio=aspect_ratio,
        )

    @staticmethod
    def _image_remote_enabled() -> bool:
        return bool(settings.image_base_url and settings.image_api_key and settings.image_model)

    def _download_bytes(self, url: str) -> bytes:
        req = Request(url=url, method="GET")
        context = ssl.create_default_context()
        with urlopen(req, timeout=settings.image_timeout_sec, context=context) as resp:
            return resp.read()

    def _mock_image_bytes(self, prompt: str, width: int = 1280, height: int = 720) -> bytes:
        seed = abs(hash(prompt)) % 255
        base = (45 + seed % 70, 75 + seed % 70, 105 + seed % 70)
        image = Image.new("RGB", (width, height), base)
        bio = BytesIO()
        image.save(bio, format="PNG")
        return bio.getvalue()

    def _generate_image_bytes(self, prompt: str, size: str = "1536x1024") -> bytes:
        if settings.use_mock_image or not self._image_remote_enabled():
            width, height = 1536, 1024
            if "x" in size:
                left, right = size.split("x", 1)
                try:
                    width, height = int(left), int(right)
                except ValueError:
                    pass
            return self._mock_image_bytes(prompt, max(320, width), max(240, height))

        payload: dict[str, Any] = {
            "model": settings.image_model,
            "prompt": prompt,
            "size": size,
            "response_format": "b64_json",
        }

        url = settings.image_base_url.rstrip("/") + settings.image_gen_path
        req = Request(url=url, method="POST")
        req.add_header("Content-Type", "application/json")
        req.add_header("Authorization", f"Bearer {settings.image_api_key}")

        context = ssl.create_default_context()
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")

        try:
            with urlopen(req, body, timeout=settings.image_timeout_sec, context=context) as resp:
                data = json.loads(resp.read().decode("utf-8"))
        except HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="ignore")
            raise RuntimeError(f"image HTTPError: {exc.code} {detail}") from exc
        except URLError as exc:
            raise RuntimeError(f"image URLError: {exc}") from exc

        items = data.get("data")
        if not isinstance(items, list) or not items:
            raise RuntimeError("image response missing data")

        first = items[0]
        if isinstance(first, dict) and first.get("b64_json"):
            return base64.b64decode(first["b64_json"])
        if isinstance(first, dict) and first.get("url"):
            return self._download_bytes(str(first["url"]))

        raise RuntimeError("image response has no b64_json/url")

    def generate_image(
        self,
        prompt: str,
        ref_image_path: Optional[str] = None,
        aspect_ratio: str = "16:9",
        resolution: str = "2K",
        additional_ref_images: Optional[List[Any]] = None,
    ) -> Optional[Image.Image]:
        try:
            if ref_image_path or additional_ref_images:
                logger.info("reference images are ignored by current image provider path")

            size = settings.image_size or "1536x1024"
            if aspect_ratio == "4:3" and size == "1536x1024":
                size = "1408x1056"

            image_bytes = self._generate_image_bytes(prompt, size=size)
            image = Image.open(BytesIO(image_bytes))
            image.load()
            return image
        except Exception:
            logger.error("error generating image", exc_info=True)
            return None

    def edit_image(
        self,
        prompt: str,
        current_image_path: str,
        original_description: Optional[str] = None,
        aspect_ratio: str = "16:9",
        resolution: str = "2K",
        additional_ref_images: Optional[List[Any]] = None,
    ) -> Optional[Image.Image]:
        edit_instruction = get_image_edit_prompt(prompt, original_description)
        return self.generate_image(
            edit_instruction,
            ref_image_path=current_image_path,
            aspect_ratio=aspect_ratio,
            resolution=resolution,
            additional_ref_images=additional_ref_images,
        )

    def extract_page_content(self, markdown_text: str, language: str = "zh") -> Dict[str, Any]:
        if self.use_mock:
            lines = [ln.strip() for ln in (markdown_text or "").splitlines() if ln.strip()]
            title = lines[0].lstrip("# ").strip() if lines else "未命名"
            points = [ln.lstrip("- ").strip() for ln in lines if ln.startswith("-")][:6]
            description_lines = [f"页面标题：{title}", "", "页面文字："] + [f"- {x}" for x in points]
            return {"title": title, "points": points, "description": "\n".join(description_lines)}

        prompt = get_ppt_page_content_extraction_prompt(markdown_text, language=language)
        data = self._generate_json(prompt, temperature=0.2)
        if not isinstance(data, dict):
            raise RuntimeError("extract_page_content output is not object")
        return data

    def _generate_text_from_image(self, prompt: str, image_path: str) -> str:
        if self.use_mock:
            return "页面结构为顶部标题+中部内容区，布局清晰，留白均衡。"

        return self.client.chat_with_image_text(
            system_prompt="You are a helpful multimodal assistant for PPT analysis.",
            user_prompt=prompt,
            image_path=image_path,
            temperature=0.2,
        )

    def generate_layout_caption(self, image_path: str) -> str:
        return self._generate_text_from_image(get_layout_caption_prompt(), image_path)

    def extract_style_description(self, image_path: str) -> str:
        return self._generate_text_from_image(get_style_extraction_prompt(), image_path)

    @staticmethod
    def _extract_candidates(text: str) -> list[str]:
        candidates: list[str] = []
        for raw in (text or "").splitlines():
            line = raw.strip()
            if not line:
                continue
            line = re.sub(r"^[0-9一二三四五六七八九十]+[\.、\)）]\s*", "", line)
            if 3 <= len(line) <= 26 and line not in candidates:
                candidates.append(line)
        return candidates

    def _mock_generate_outline(self, project_context: BananaProjectContext) -> List[Dict]:
        idea = project_context.idea_prompt or ""
        candidates = self._extract_candidates(idea)

        if not candidates:
            candidates = ["项目背景", "现状分析", "关键问题", "目标与策略", "执行计划", "资源保障"]

        pages: list[dict[str, Any]] = [
            {"title": "封面", "points": []},
            {"title": "目录", "points": []},
        ]

        for item in candidates:
            pages.append(
                {
                    "title": item,
                    "points": [
                        f"{item}的核心结论",
                        f"{item}的关键数据",
                        f"{item}的落地建议",
                    ],
                }
            )

        while len(pages) < 8:
            idx = len(pages) - 1
            pages.append(
                {
                    "title": f"专题分析{idx}",
                    "points": ["关键结论", "主要证据", "下一步行动"],
                }
            )

        return pages

    def _mock_parse_outline_text(self, project_context: BananaProjectContext) -> List[Dict]:
        text = project_context.outline_text or ""
        titles = self._extract_candidates(text)
        if not titles:
            return self._mock_generate_outline(project_context)

        pages: list[dict[str, Any]] = []
        for idx, title in enumerate(titles, start=1):
            pages.append({"title": title, "points": [f"要点{idx}.1", f"要点{idx}.2"]})
        return pages

    @staticmethod
    def _mock_generate_page_description(
        project_context: BananaProjectContext,
        outline: List[Dict],
        page_outline: Dict,
        page_index: int,
    ) -> Dict:
        title = str(page_outline.get("title") or f"第{page_index}页")
        points = [str(x) for x in page_outline.get("points", []) if str(x).strip()]
        if not points:
            points = ["核心结论", "关键数据", "行动建议"]

        if page_index == 1:
            text = (
                f"页面标题：{title}\n\n"
                "页面文字：\n"
                f"- {project_context.idea_prompt or title}\n"
                "- 汇报人：自动生成\n"
                f"- 日期：{datetime.now().strftime('%Y-%m-%d')}"
            )
        else:
            text = (
                f"页面标题：{title}\n\n"
                "页面文字：\n"
                f"- {points[0]}\n"
                f"- {points[1] if len(points) > 1 else '补充说明'}\n"
                f"- {points[2] if len(points) > 2 else '下一步计划'}\n\n"
                "其他页面素材：\n"
                "- 建议使用相关图标或数据图"
            )

        return {"text": text}

    def _mock_description_split(self, project_context: BananaProjectContext, outline: List[Dict]) -> List[str]:
        pages = self.flatten_outline(outline)
        out: list[str] = []
        for idx, page in enumerate(pages, start=1):
            desc = self._mock_generate_page_description(project_context, pages, page, idx)
            out.append(str(desc.get("text", "")))
        return out

    @staticmethod
    def _rewrite_bullets(lines: list[str], mode: str) -> list[str]:
        while len(lines) < 3:
            lines.append("补充要点")

        if mode == "concise":
            return [line[:18] + ("..." if len(line) > 18 else "") for line in lines[:3]]
        if mode == "management":
            return [f"结果：{lines[0]}", f"风险：{lines[1]}", f"决策：{lines[2]}"]
        if mode == "technical":
            return [f"现状：{lines[0]}", f"细节：{lines[1]}", f"计划：{lines[2]}"]
        return lines[:3]

    def _mock_refine_descriptions(self, current_descriptions: List[Dict], user_requirement: str) -> List[str]:
        req = (user_requirement or "").lower()
        mode = "default"
        if "精简" in req or "concise" in req:
            mode = "concise"
        elif "管理" in req or "management" in req:
            mode = "management"
        elif "技术" in req or "technical" in req:
            mode = "technical"

        out: list[str] = []
        for desc in current_descriptions:
            raw = desc.get("description_content", "")
            if isinstance(raw, dict):
                raw_text = str(raw.get("text") or "")
            else:
                raw_text = str(raw)

            title_match = re.search(r"页面标题[：:]\s*(.+)", raw_text)
            title = title_match.group(1).strip() if title_match else str(desc.get("title") or "未命名")

            section = ""
            m = re.search(r"页面文字[：:]\s*(.+?)(?:\n\s*(?:图片素材|其他页面素材)[：:]|$)", raw_text, flags=re.S)
            if m:
                section = m.group(1)

            lines = [
                re.sub(r"^[-*•\d\.\)）\s]+", "", line).strip()
                for line in section.splitlines()
                if line.strip()
            ]
            lines = [line for line in lines if line]
            rewritten = self._rewrite_bullets(lines, mode)

            text = "\n".join(
                [
                    f"页面标题：{title}",
                    "",
                    "页面文字：",
                    f"- {rewritten[0]}",
                    f"- {rewritten[1]}",
                    f"- {rewritten[2]}",
                ]
            )
            out.append(text)

        return out


def make_project_context_from_row(project_row: Any) -> BananaProjectContext:
    title = str(project_row["title"])
    style = str(project_row["style"])
    material_text = str(project_row["material_text"] or "")

    idea_prompt = build_idea_prompt(title, style, material_text)
    outline_text = str(project_row["outline_text"] or "")

    return BananaProjectContext(
        idea_prompt=idea_prompt,
        outline_text=outline_text,
        description_text=None,
        creation_type=str(project_row.get("creation_type", "idea") if isinstance(project_row, dict) else project_row["creation_type"]),
        outline_requirements=None,
        description_requirements=None,
        reference_files_content=[],
    )





def _normalize_outline_title(raw: str) -> str:
    txt = str(raw or "").strip()
    txt = re.sub(r"^\s*第?\s*\d+\s*页?\s*[:：.、\)\]]\s*", "", txt)
    txt = re.sub(r"\s+", " ", txt).strip(" -:：\t\n")
    return txt


def _is_cover_title(title: str) -> bool:
    t = _normalize_outline_title(title).lower()
    if not t:
        return False
    keywords = (
        "封面", "标题页", "title", "cover", "front page", "opening",
    )
    return any(k in t for k in keywords)


def _is_toc_title(title: str) -> bool:
    t = _normalize_outline_title(title).lower()
    if not t:
        return False
    keywords = (
        "目录", "议程", "agenda", "contents", "table of contents", "toc",
    )
    return any(k in t for k in keywords)


def enforce_target_pages(pages: list[dict[str, Any]], target_pages: int) -> list[dict[str, Any]]:
    target = max(8, min(12, int(target_pages)))
    normalized_in = [dict(page) for page in pages if isinstance(page, dict)]

    # Keep only meaningful body pages; cover/toc are always re-inserted in fixed slots.
    body: list[dict[str, Any]] = []
    seen_titles: set[str] = set()

    for page in normalized_in:
        title = _normalize_outline_title(page.get("title") or "")
        if not title:
            continue
        if _is_cover_title(title) or _is_toc_title(title):
            continue

        key = title.lower()
        if key in seen_titles:
            continue
        seen_titles.add(key)

        points = [str(x).strip() for x in list(page.get("points") or []) if str(x).strip()]
        item: dict[str, Any] = {"title": title, "points": points[:5]}
        if page.get("part"):
            item["part"] = str(page.get("part"))
        body.append(item)

    desired_body = max(0, target - 2)

    seed = random.Random(len(body) * 9973 + target)
    while len(body) < desired_body:
        idx = len(body) + 1
        if body:
            ref = body[min(len(body) - 1, seed.randrange(len(body)))]
            base = str(ref.get("title") or f"专题分析{idx}")
            points = [str(x) for x in list(ref.get("points") or [])][:2]
        else:
            base = f"专题分析{idx}"
            points = []

        points.extend(["关键结论", "行动建议"])
        title = f"{base}扩展" if not base.endswith("扩展") else base
        body.append({"title": title, "points": points[:3]})

    body = body[:desired_body]

    # Canonical outline skeleton required by UI and rendering flow.
    out: list[dict[str, Any]] = [
        {"title": "封面", "points": []},
        {"title": "目录", "points": []},
    ]
    out.extend(body)

    # Safety fill if body filtering produced fewer pages than expected.
    while len(out) < target:
        i = len(out) + 1
        out.append({"title": f"第{i}页", "points": ["关键结论", "主要依据", "下一步行动"]})

    return out










