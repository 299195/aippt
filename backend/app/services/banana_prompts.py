"""Banana prompt templates used by the generation workflow."""

from __future__ import annotations

import json
import logging
import os
from typing import Dict, List, Optional, Protocol

logger = logging.getLogger(__name__)


class ProjectContextLike(Protocol):
    idea_prompt: str | None
    outline_text: str | None
    description_text: str | None
    creation_type: str
    outline_requirements: str | None
    description_requirements: str | None
    reference_files_content: list[dict[str, str]]


LANGUAGE_CONFIG = {
    "zh": {
        "instruction": "请使用全中文输出。",
        "ppt_text": "PPT文字请使用全中文。",
    },
    "ja": {
        "instruction": "すべて日本語で出力してください。",
        "ppt_text": "PPTのテキストは全て日本語で出力してください。",
    },
    "en": {
        "instruction": "Please output all in English.",
        "ppt_text": "Use English for PPT text.",
    },
    "auto": {
        "instruction": "",
        "ppt_text": "",
    },
}

DETAIL_LEVEL_SPECS = {
    "concise": "每页给出3-4条高密度要点，尽量短句但仍保留事实和结论。",
    "default": "每页给出4-5条完整要点，每条20-40字，覆盖背景、方法、结果与结论。",
    "detailed": "每页给出5-7条完整内容，优先引用输入资料中的事实/数据，形成可直接上屏的讲解文字。",
}

_OUTLINE_JSON_FORMAT = """\
1. Simple format (for short PPTs without major sections):
[{"title": "title1", "points": ["point1", "point2"]}, {"title": "title2", "points": ["point1", "point2"]}]

2. Part-based format (for longer PPTs with major sections):
[
    {
    "part": "Part 1: Introduction",
    "pages": [
        {"title": "Welcome", "points": ["point1", "point2"]},
        {"title": "Overview", "points": ["point1", "point2"]}
    ]
    },
    {
    "part": "Part 2: Main Content",
    "pages": [
        {"title": "Topic 1", "points": ["point1", "point2"]},
        {"title": "Topic 2", "points": ["point1", "point2"]}
    ]
    }
]"""


def _format_reference_files_xml(reference_files_content: Optional[List[Dict[str, str]]]) -> str:
    if not reference_files_content:
        return ""
    xml_parts = ["<uploaded_files>"]
    for file_info in reference_files_content:
        filename = file_info.get("filename", "unknown")
        content = file_info.get("content", "")
        xml_parts.append(f'  <file name="{filename}">')
        xml_parts.append("    <content>")
        xml_parts.append(content)
        xml_parts.append("    </content>")
        xml_parts.append("  </file>")
    xml_parts.append("</uploaded_files>")
    xml_parts.append("")
    return "\n".join(xml_parts)


def _build_prompt(prompt_text: str, reference_files_content=None) -> str:
    return _format_reference_files_xml(reference_files_content) + prompt_text


def _get_original_input(project_context: ProjectContextLike) -> str:
    if project_context.creation_type == "idea" and project_context.idea_prompt:
        return project_context.idea_prompt
    if project_context.creation_type == "outline" and project_context.outline_text:
        return f"用户提供的大纲：\n{project_context.outline_text}"
    if project_context.creation_type == "descriptions" and project_context.description_text:
        return f"用户提供的描述：\n{project_context.description_text}"
    return project_context.idea_prompt or ""


def _get_original_input_labeled(project_context: ProjectContextLike) -> str:
    text = "\n原始输入信息：\n"
    if project_context.creation_type == "idea" and project_context.idea_prompt:
        text += f"- PPT构想：{project_context.idea_prompt}\n"
    elif project_context.creation_type == "outline" and project_context.outline_text:
        text += f"- 用户提供的大纲文本：\n{project_context.outline_text}\n"
    elif project_context.creation_type == "descriptions" and project_context.description_text:
        text += f"- 用户提供的页面描述文本：\n{project_context.description_text}\n"
    elif project_context.idea_prompt:
        text += f"- 用户输入：{project_context.idea_prompt}\n"
    return text


def _get_previous_requirements_text(previous_requirements: Optional[List[str]]) -> str:
    if not previous_requirements:
        return ""
    prev_list = "\n".join([f"- {req}" for req in previous_requirements])
    return f"\n\n之前用户提出的修改要求：\n{prev_list}\n"


def _format_extra_field_instructions(extra_fields: list | None) -> str:
    if not extra_fields:
        return ""
    parts = [f"{f}：[关于{f}的建议]" for f in extra_fields]
    return "\n".join([""] + parts)


def _format_requirements(requirements: str | None, context: str = "outline") -> str:
    if requirements and requirements.strip():
        if context == "description":
            marker_example = (
                "For example, if the user asks to avoid certain symbols, "
                "do NOT use them in the page content, but still use structural markers "
                "like '页面文字：', '图片素材：', and '<!-- PAGE_END -->' as-is."
            )
        else:
            marker_example = (
                "For example, if the user asks to avoid '#' symbols, "
                "do NOT use '#' in the page content, but still use '## Title' as "
                "the structural heading delimiter between pages."
            )
        return (
            "<user_requirements>\n"
            f"{requirements.strip()}\n"
            "</user_requirements>\n"
            "Note: The requirements above apply to the generated content of each page and "
            "take precedence over other content-related instructions. The required output format "
            f"and structural markers must still be used as-is. {marker_example}\n\n"
        )
    return ""


def get_default_output_language() -> str:
    return os.getenv("OUTPUT_LANGUAGE", "zh")


def get_language_instruction(language: str | None = None) -> str:
    lang = language if language else get_default_output_language()
    config = LANGUAGE_CONFIG.get(lang, LANGUAGE_CONFIG["zh"])
    return config["instruction"]


def get_ppt_language_instruction(language: str | None = None) -> str:
    lang = language if language else get_default_output_language()
    config = LANGUAGE_CONFIG.get(lang, LANGUAGE_CONFIG["zh"])
    return config["ppt_text"]


def get_outline_generation_prompt(project_context: ProjectContextLike, language: str | None = None) -> str:
    idea_prompt = project_context.idea_prompt or ""

    prompt = f"""\
You are a helpful assistant that generates an outline for a ppt.

You can organize the content in two ways:

{_OUTLINE_JSON_FORMAT}

Choose the format that best fits the content. Use parts when the PPT has clear major sections.
Always include a cover page as page 1 and an agenda page as page 2. Page 1 should contain only polished presentation title/subtitle (not a raw copy of user prompt). Page 2 should be agenda and must align with subsequent page titles.

The user's request: {idea_prompt}.
{_format_requirements(project_context.outline_requirements)}Now generate the outline, don't include any other text.
{get_language_instruction(language)}
"""

    return _build_prompt(prompt, project_context.reference_files_content)


def get_outline_generation_prompt_markdown(project_context: ProjectContextLike, language: str | None = None) -> str:
    idea_prompt = project_context.idea_prompt or ""

    prompt = f"""\
You are a helpful assistant that generates an outline for a ppt.

You can organize the content in two ways:

1. Simple format (for short PPTs without major sections):
## title1
- point1
- point2

## title2
- point1
- point2

2. Part-based format (for longer PPTs with major sections):
# Part 1: Introduction
## Welcome
- point1
- point2

## Overview
- point1
- point2

# Part 2: Main Content
## Topic 1
- point1
- point2

## Topic 2
- point1
- point2

Constraints:
- Title should not contain page number.
- Choose the format that best fits the content. Use parts when the PPT has clear major sections.
- Always include a cover page as page 1 and an agenda page as page 2. Page 1 should contain only polished presentation title/subtitle (not a raw copy of user prompt). Page 2 should be agenda and must align with subsequent page titles.

The user's request: {idea_prompt}.
{_format_requirements(project_context.outline_requirements)}Now generate the outline, strictly follow the format provided above, don't include any other text. Output `<!-- END -->` on the last line when finished.
{get_language_instruction(language)}
"""

    return _build_prompt(prompt, project_context.reference_files_content)


def get_outline_parsing_prompt(project_context: ProjectContextLike, language: str | None = None) -> str:
    outline_text = project_context.outline_text or ""

    prompt = f"""\
You are a helpful assistant that parses a user-provided PPT outline text into a structured format.

The user has provided the following outline text:

{outline_text}

Your task is to analyze this text and convert it into a structured JSON format WITHOUT modifying any of the original text content.
You should only reorganize and structure the existing content, preserving all titles, points, and text exactly as provided.

You can organize the content in two ways:

{_OUTLINE_JSON_FORMAT}

Important rules:
- DO NOT modify, rewrite, or change any text from the original outline
- DO NOT add new content that wasn't in the original text
- DO NOT remove any content from the original text
- Only reorganize the existing content into the structured format
- Preserve all titles, bullet points, and text exactly as they appear
- If the text has clear sections/parts, use the part-based format
- Extract titles and points from the original text, keeping them exactly as written

Now parse the outline text above into the structured format. Return only the JSON, don't include any other text.
{get_language_instruction(language)}
"""

    return _build_prompt(prompt, project_context.reference_files_content)


def get_outline_parsing_prompt_markdown(project_context: ProjectContextLike, language: str | None = None) -> str:
    outline_text = project_context.outline_text or ""

    prompt = f"""\
You are a helpful assistant that parses a user-provided PPT outline text into a structured Markdown format.

The user has provided the following outline text:

{outline_text}

Your task is to analyze this text and convert it into a structured Markdown outline WITHOUT modifying any of the original text content.

Output rules:
- Use `# Part Name` for major sections (only if the text has clear parts/chapters)
- Use `## Page Title` for each page
- Use `- ` bullet points for key points under each page
- Preserve all titles, points, and text exactly as provided
- Do NOT wrap in code blocks or add any extra text

Now parse the outline text above into the Markdown format. Output `<!-- END -->` on the last line when finished.
{get_language_instruction(language)}
"""

    return _build_prompt(prompt, project_context.reference_files_content)


def get_description_to_outline_prompt(project_context: ProjectContextLike, language: str | None = None) -> str:
    description_text = project_context.description_text or ""

    prompt = f"""\
You are a helpful assistant that analyzes a user-provided PPT description text and extracts the outline structure from it.

The user has provided the following description text:

{description_text}

Your task is to analyze this text and extract the outline structure (titles and key points) for each page.
You should identify:
1. How many pages are described
2. The title for each page
3. The key points or content structure for each page

You can organize the content in two ways:

{_OUTLINE_JSON_FORMAT}

Important rules:
- Extract the outline structure from the description text
- Identify page titles and key points
- If the text has clear sections/parts, use the part-based format
- Preserve the logical structure and organization from the original text
- The points should be concise summaries of the main content for each page

Now extract the outline structure from the description text above. Return only the JSON, don't include any other text.
{get_language_instruction(language)}
"""

    return _build_prompt(prompt, project_context.reference_files_content)


def get_description_to_outline_prompt_markdown(project_context: ProjectContextLike, language: str | None = None) -> str:
    description_text = project_context.description_text or ""

    prompt = f"""\
You are a helpful assistant that analyzes a user-provided PPT description text and extracts the outline structure.

The user has provided the following description text:

{description_text}

Your task is to extract the outline structure (titles and key points) for each page.

Output rules:
- Use `# Part Name` for major sections (only if the text has clear parts/chapters)
- Use `## Page Title` for each page
- Use `- ` bullet points for key points under each page
- Preserve the logical structure from the original text
- Do NOT wrap in code blocks or add any extra text

Now extract the outline structure from the description text above. Output `<!-- END -->` on the last line when finished.
{get_language_instruction(language)}
"""

    return _build_prompt(prompt, project_context.reference_files_content)


def get_page_description_prompt(
    project_context: ProjectContextLike,
    outline: list,
    page_outline: dict,
    page_index: int,
    part_info: str = "",
    language: str | None = None,
    detail_level: str = "default",
    extra_fields: list | None = None,
) -> str:
    original_input = _get_original_input(project_context)
    detail_instruction = DETAIL_LEVEL_SPECS.get(detail_level, DETAIL_LEVEL_SPECS["default"])
    page_title = str(page_outline.get("title") or "").strip()
    page_points = [str(x).strip() for x in list(page_outline.get("points") or []) if str(x).strip()]
    page_points_text = "\n".join([f"- {p}" for p in page_points]) if page_points else "- （无显式要点，请按主题与资料补全）"

    outline_json = json.dumps(outline, ensure_ascii=False, indent=2)

    body_titles: list[str] = []
    for idx, item in enumerate(outline or []):
        if not isinstance(item, dict):
            continue
        title = str(item.get("title") or "").strip()
        if not title:
            continue
        title_low = title.lower()
        if idx == 0:
            continue
        if any(k in title_low for k in ("agenda", "contents", "toc", "目录", "议程")):
            continue
        body_titles.append(title)

    toc_reference = "\n".join([f"{i + 1}. {t}" for i, t in enumerate(body_titles)]) or "（后续标题待模型补全）"
    is_toc_page = page_index == 2 or any(k in page_title.lower() for k in ("agenda", "contents", "toc", "目录", "议程"))

    cover_rule = "- 当前为封面页：只输出标题、副标题、演讲人/单位/日期，不要正文段落、图表或数据结论。" if page_index == 1 else ""
    toc_rule = (
        "- 当前为目录页：页面文字必须是目录项，且顺序与后续页面标题完全一致。\n"
        f"  目录目标如下：\n{toc_reference}"
        if is_toc_page
        else ""
    )
    body_rule = (
        "- 普通内容页：页面文字至少4条，建议5-6条，覆盖背景/问题、方法/方案、关键证据、结论影响、下一步。"
        if (page_index != 1 and not is_toc_page)
        else ""
    )

    prompt = f"""\
你是资深咨询顾问 + 演示文稿撰写专家。请为单页PPT生成“可直接上屏”的完整文字内容。

用户原始需求：
{original_input}

完整大纲：
{outline_json}

当前页信息（第 {page_index} 页）{part_info}：
- 页面标题：{page_title}
- 页面要点：
{page_points_text}

{_format_requirements(project_context.description_requirements, "description")}## 生成要求
- 生成的“页面文字”会直接渲染到PPT，请只输出可展示内容，不要解释你在做什么。
- 正文第一条不得与页面标题重复或近似改写。
- 内容必须基于用户主题和输入资料，避免空泛套话。
- 若资料中没有明确数字，禁止编造数据、图表结论或百分比。
- {detail_instruction}
{cover_rule}
{toc_rule}
{body_rule}

## 输出格式（严格遵守）
页面标题：[与当前页一致或更自然的标题]

页面文字：
- [要点1]
- [要点2]
- [要点3]
- [要点4]
[可继续补充第5-6条]

图片素材：
[仅在参考文件中存在可用图片时，使用 markdown 图片链接；否则省略该字段]
{_format_extra_field_instructions(extra_fields)}

## 关于图片
如果参考文件中包含以 /files/ 开头的本地文件URL图片（例如 /files/mineru/xxx/image.png），请将这些图片以 markdown 格式输出，例如：![图片描述](/files/mineru/xxx/image.png)。
{get_language_instruction(language)}
"""

    return _build_prompt(prompt, project_context.reference_files_content)

def get_all_descriptions_stream_prompt(
    project_context: ProjectContextLike,
    outline: list,
    flat_pages: list,
    language: str | None = None,
    detail_level: str = "default",
    extra_fields: list | None = None,
) -> str:
    original_input = _get_original_input(project_context)
    detail_instruction = DETAIL_LEVEL_SPECS.get(detail_level, DETAIL_LEVEL_SPECS["default"])

    outline_lines = []
    for i, page in enumerate(flat_pages):
        part_str = f"  [章节: {page['part']}]" if page.get("part") else ""
        points_str = ", ".join(page.get("points", []))
        outline_lines.append(f"第 {i + 1} 页：{page.get('title', '')}{part_str}\n  要点：{points_str}")
    pages_outline_text = "\n".join(outline_lines)

    toc_targets: list[str] = []
    for idx, page in enumerate(flat_pages):
        title = str(page.get("title") or "").strip()
        if not title:
            continue
        title_low = title.lower()
        if idx == 0:
            continue
        if any(k in title_low for k in ("agenda", "contents", "toc", "目录", "议程")):
            continue
        toc_targets.append(title)
    toc_targets_text = "\n".join([f"{i + 1}. {t}" for i, t in enumerate(toc_targets)]) or "（后续标题待模型补全）"

    prompt = f"""\
你是资深咨询顾问 + 演示文稿撰写专家。请按大纲逐页生成“可直接渲染到PPT”的完整页面内容。

用户原始需求：
{original_input}

完整页级大纲：
{pages_outline_text}

{_format_requirements(project_context.description_requirements, "description")}请按页顺序输出，先输出 `<!-- BEGIN -->`，每页结尾输出 `<!-- PAGE_END -->`，最后输出 `<!-- END -->`。

## 统一硬性约束
- 页面文字会直接上屏，不要输出解释性旁白或注释。
- 每页正文第一条不得与该页标题重复或近似改写。
- 内容必须基于主题和资料，避免空泛套话。
- 若资料没有明确数字，禁止编造图表数据或百分比。
- 细致程度：{detail_instruction}

## 页面约束
- 第1页封面：只写标题、副标题、演讲人/单位/日期，不写正文段落、图表结论。
- 目录页（通常第2页）：目录项必须与后续页面标题严格对应、顺序一致。目录目标如下：
{toc_targets_text}
- 普通内容页：至少4条，建议5-6条，覆盖背景/问题、方法/方案、关键证据、结论影响、下一步。

## 输出格式（严格遵守）
```
<!-- BEGIN -->
页面标题：[第1页标题]

页面文字：
- [要点1]
- [要点2]
- [要点3]
- [要点4]
[可继续补充第5-6条]

图片素材：
[如果参考文件存在可用图片，使用 markdown 链接，如 ![描述](/files/xxx/image.png)；否则省略该字段]
{_format_extra_field_instructions(extra_fields)}
<!-- PAGE_END -->
...
<!-- END -->
```

现在开始生成，并严格遵守格式。
{get_language_instruction(language)}
"""

    return _build_prompt(prompt, project_context.reference_files_content)

def get_description_split_prompt(
    project_context: ProjectContextLike,
    outline: List[Dict],
    language: str | None = None,
) -> str:
    outline_json = json.dumps(outline, ensure_ascii=False, indent=2)
    description_text = project_context.description_text or ""

    prompt = f"""\
You are a helpful assistant that splits a complete PPT description text into individual page descriptions.

The user has provided a complete description text:

{description_text}

We have already extracted the outline structure:

{outline_json}

Your task is to split the description text into individual page descriptions based on the outline structure.
For each page in the outline, extract the corresponding description from the original text.

Return a JSON array where each element corresponds to a page in the outline (in the same order).
Each element should be a string containing the page description in the following format:

页面标题：[页面标题]

页面文字：
- [要点1]
- [要点2]
...

其他页面素材（如果有排版、风格、素材等细节）

Important rules:
- Split the description text according to the outline structure
- Each page description should match the corresponding page in the outline
- Preserve all important content from the original text, including layout details, style requirements and material descriptions
- If a page in the outline doesn't have a clear description in the text, create a reasonable description based on the outline

Now split the description text into individual page descriptions. Return only the JSON array, don't include any other text.
{get_language_instruction(language)}
"""

    return _build_prompt(prompt, project_context.reference_files_content)


def get_descriptions_refinement_prompt(
    current_descriptions: List[Dict],
    user_requirement: str,
    project_context: ProjectContextLike,
    outline: List[Dict] | None = None,
    previous_requirements: Optional[List[str]] = None,
    language: str | None = None,
) -> str:
    outline_text = ""
    if outline:
        outline_json = json.dumps(outline, ensure_ascii=False, indent=2)
        outline_text = f"\n\n完整的 PPT 大纲：\n{outline_json}\n"

    all_descriptions_text = "当前所有页面的描述：\n\n"
    has_any_description = False
    for desc in current_descriptions:
        page_num = desc.get("index", 0) + 1
        title = desc.get("title", "未命名")
        content = desc.get("description_content", "")
        if isinstance(content, dict):
            content = content.get("text", "")

        if content:
            has_any_description = True
            all_descriptions_text += f"--- 第 {page_num} 页：{title} ---\n{content}\n\n"
        else:
            all_descriptions_text += f"--- 第 {page_num} 页：{title} ---\n(当前没有内容)\n\n"

    if not has_any_description:
        all_descriptions_text = "当前所有页面的描述：\n\n(当前没有内容，需要基于大纲生成新的描述)\n\n"

    prompt = f"""\
You are a helpful assistant that modifies PPT page descriptions based on user requirements.
{_get_original_input_labeled(project_context)}{outline_text}
{all_descriptions_text}
{_get_previous_requirements_text(previous_requirements)}
**用户现在提出新的要求：{user_requirement}**

请根据用户的要求修改和调整所有页面的描述。你可以：
- 修改页面标题和内容
- 调整页面文字的详细程度
- 添加或删除要点
- 调整描述的结构和表达
- 确保所有页面描述都符合用户的要求
- 如果当前没有内容，请根据大纲和用户要求创建新的描述

请为每个页面生成修改后的描述，格式如下：

页面标题：[页面标题]

页面文字：
- [要点1]
- [要点2]
...
其他页面素材（如果有请加上，包括markdown图片链接等）

请返回一个 JSON 数组，每个元素是一个字符串，对应每个页面的修改后描述（按页面顺序）。

现在请根据用户要求修改所有页面描述，只输出 JSON 数组，不要包含其他文字。
{get_language_instruction(language)}
"""

    return _build_prompt(prompt, project_context.reference_files_content)


def get_image_generation_prompt(
    page_desc: str,
    outline_text: str,
    current_section: str,
    has_material_images: bool = False,
    extra_requirements: str | None = None,
    language: str | None = None,
    has_template: bool = True,
    page_index: int = 1,
    aspect_ratio: str = "16:9",
) -> str:
    material_images_note = ""
    if has_material_images:
        material_images_note = (
            "\n\n提示："
            + ("除了模板参考图片（用于风格参考）外，还提供了额外的素材图片。" if has_template else "用户提供了额外的素材图片。")
            + "这些素材图片是可供挑选和使用的元素，请根据页面内容智能选择并融合。"
        )

    extra_req_text = ""
    if extra_requirements and extra_requirements.strip():
        extra_req_text = f"\n\n额外要求（请务必遵循）：\n{extra_requirements.strip()}\n"

    template_style_guideline = "- 配色和设计语言和模板图片严格相似。" if has_template else "- 严格按照风格描述进行设计。"
    forbidden_template_text_guidline = "- 只参考风格设计，禁止出现模板中的文字。\n" if has_template else ""

    prompt = f"""\
你是一位专家级UI UX演示设计师，专注于生成设计良好的PPT页面。
当前PPT页面的页面描述如下:
<page_description>
{page_desc}
</page_description>

<outline>
{outline_text}
</outline>

<current_section>
{current_section}
</current_section>

<design_guidelines>
- 要求文字清晰锐利, 画面为4K分辨率，{aspect_ratio}比例。
{template_style_guideline}
- 根据内容和要求自动设计最完美的构图，不重不漏地渲染"页面文字"段落中的文本。
- 如非必要，禁止出现 markdown 格式符号（如 # 和 * 等）。
{forbidden_template_text_guidline}
</design_guidelines>
{get_ppt_language_instruction(language)}
{material_images_note}{extra_req_text}

{"**注意：当前页面为ppt的封面页，请你采用专业的封面设计美学技巧，务必凸显出页面标题，分清主次，确保一下就能抓住观众的注意力。**" if page_index == 1 else ""}
"""

    return prompt


def get_image_edit_prompt(edit_instruction: str, original_description: str | None = None) -> str:
    if original_description:
        if "其他页面素材" in original_description:
            original_description = original_description.split("其他页面素材")[0].strip()

        return f"""\
该PPT页面的原始页面描述为：
{original_description}

现在，根据以下指令修改这张PPT页面：{edit_instruction}

要求维持原有的文字内容和设计风格，只按照指令进行修改。
"""

    return (
        f"根据以下指令修改这张PPT页面：{edit_instruction}\n"
        "保持原有的内容结构和设计风格，只按照指令进行修改。"
    )


def get_ppt_page_content_extraction_prompt(markdown_text: str, language: str | None = None) -> str:
    prompt = f"""\
You are a helpful assistant that extracts structured PPT page content from parsed document text.

The following markdown text was extracted from a single PPT slide:

<slide_content>
{markdown_text}
</slide_content>

Your task is to extract the following structured information from this slide:

1. **title**: The main title/heading of the slide
2. **points**: A list of key bullet points or content items on the slide
3. **description**: A complete page description suitable for regenerating this slide, following this format:

页面标题：[title]

页面文字：
- [point 1]
- [point 2]
...

其他页面素材（如果有图表、表格、公式等描述，保留原文中的markdown图片完整形式）

Return a JSON object with exactly these three fields: "title", "points" (array of strings), "description" (string).
Return only the JSON, no other text.
{get_language_instruction(language)}
"""
    return prompt


def get_layout_caption_prompt() -> str:
    return """\
You are a professional PPT layout analyst. Describe the visual layout and composition of this PPT slide image in detail.

Focus on:
1. **Overall layout**: How elements are arranged
2. **Text placement**: Where text blocks are positioned
3. **Visual elements**: Position and size of images/charts/icons
4. **Spacing and proportions**: How space is distributed

Output a concise layout description in Chinese. Format:

排版布局：
- 整体结构：[描述]
- 标题位置：[描述]
- 内容区域：[描述]
- 视觉元素：[描述]

Only describe the layout and spatial arrangement. Do not describe colors, text content, or style.
"""


def get_style_extraction_prompt() -> str:
    return """\
You are a professional PPT design analyst. Analyze this image and extract a detailed style description that can be used to generate PPT slides with a similar visual style.

Focus on:
1. Color palette
2. Typography style
3. Design elements
4. Overall mood
5. Layout tendencies

Output a concise style description in Chinese that can be directly used as a style prompt for PPT generation.
Only output the style description text, no other content.
"""



