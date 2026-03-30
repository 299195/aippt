from __future__ import annotations

from dataclasses import dataclass
import random
import re
from typing import Dict, List

from app.services.model_client import ModelClient


STYLE_PREFIX = {
    "management": ("结果", "风险", "决策"),
    "technical": ("现状", "细节", "计划"),
}
STOP_PREFIXES = ("主题:", "用户提纲:", "上传资料:")
ALLOWED_SLIDE_TYPES = {"title", "summary", "risk", "timeline", "data"}
RISK_HINTS = ("风险", "问题", "挑战", "阻塞", "缺口", "隐患")
TIMELINE_HINTS = ("计划", "路线", "进展", "里程碑", "阶段", "排期", "节奏")
DATA_HINTS = (
    "数据",
    "指标",
    "同比",
    "环比",
    "增长",
    "下降",
    "转化",
    "效率",
    "成本",
    "收入",
    "预算",
    "gmv",
    "roi",
    "%",
)


@dataclass
class LLMAdapter:
    use_mock: bool = True

    def __post_init__(self) -> None:
        self.client = ModelClient()
        if not self.use_mock and not self.client.enabled():
            raise RuntimeError(
                "模型未配置完整。请设置 model_provider.json 或 MODEL_BASE_URL / MODEL_API_KEY / MODEL_NAME (或 MODEL_ENDPOINT_ID)，或将 USE_MOCK_LLM=true。"
            )

    def _normalize(self, text: str) -> str:
        return re.sub(r"\s+", " ", text).strip(" -\t\r")

    def _material_sentences(self, material: str) -> List[str]:
        lines = []
        for raw in material.splitlines():
            line = self._normalize(raw)
            if not line:
                continue
            if any(line.startswith(prefix) for prefix in STOP_PREFIXES):
                continue
            for seg in re.split(r"[。；;！!？?]", line):
                s = self._normalize(seg)
                if len(s) >= 8:
                    lines.append(s)
        return lines

    def _extract_outline_candidates(self, material: str) -> List[str]:
        candidates: List[str] = []
        for raw in material.splitlines():
            line = self._normalize(raw)
            if not line:
                continue
            if any(line.startswith(prefix) for prefix in STOP_PREFIXES):
                continue
            line = re.sub(r"^[0-9一二三四五六七八九十]+[\.、\)]\s*", "", line)
            if 2 <= len(line) <= 22 and line not in candidates:
                candidates.append(line)
        return candidates

    def _guess_slide_type(self, section: str, page: int, total_pages: int) -> str:
        s = section.lower()
        if page == 1 or any(x in section for x in ("封面", "标题", "主题")):
            return "title"
        if any(x in section for x in RISK_HINTS):
            return "risk"
        if any(x in section for x in TIMELINE_HINTS):
            return "timeline"
        if any(x in s for x in DATA_HINTS):
            return "data"
        # Give the middle of deck a higher chance to be data slide when no explicit hint.
        if total_pages >= 8 and page in (4, 5):
            return "data"
        return "summary"

    def _llm_generate_outline(self, title: str, style: str, material: str, target_pages: int) -> List[str]:
        pages = max(8, min(12, target_pages))
        sys = "你是PPT策划专家。根据主题和资料直接规划8~12页汇报大纲，不需要六要素拆解。"
        usr = (
            "只返回JSON：{\"outline\": [\"第1页：...\", ...]}，每页标题要具体可执行。"
            f"\n主题: {title}\n风格: {style}\n页数要求: {pages}\n资料:\n{material}"
        )
        data = self.client.chat_json(sys, usr, temperature=0.5)
        outline = [self._normalize(str(x)) for x in data.get("outline", []) if self._normalize(str(x))]
        if len(outline) < 8:
            raise RuntimeError("模型返回大纲页数不足8")
        return outline[:pages]

    def generate_outline(self, title: str, style: str, material: str, target_pages: int) -> List[str]:
        pages = max(8, min(12, target_pages))
        if not self.use_mock:
            return self._llm_generate_outline(title, style, material, pages)

        candidates = self._extract_outline_candidates(material)
        outline = ["第1页：封面", "第2页：目录"]
        for i, item in enumerate(candidates, start=3):
            outline.append(f"第{i}页：{item}")
            if len(outline) >= pages:
                break
        while len(outline) < pages:
            outline.append(f"第{len(outline)+1}页：专题分析{len(outline)-1}")
        return outline[:pages]

    def _llm_plan_slide_types(self, title: str, outline: List[str], style: str, material: str) -> List[str]:
        allowed = "|".join(sorted(ALLOWED_SLIDE_TYPES))
        sys = "你是企业演示架构师。为每个大纲页分配最合适的页型，只返回JSON。"
        usr = (
            f"输出格式: {{\"slide_types\": [\"{allowed}\", ...]}}，长度必须与outline一致。"
            "\n要求: 至少包含1页timeline，若资料里有数字或指标，至少包含1页data。"
            f"\n主题: {title}\n风格: {style}\noutline: {outline}\n资料:\n{material}"
        )
        data = self.client.chat_json(sys, usr, temperature=0.2)
        raw = [self._normalize(str(x)).lower() for x in data.get("slide_types", [])]
        if len(raw) != len(outline):
            raise RuntimeError("slide_types length mismatch")
        out = [x if x in ALLOWED_SLIDE_TYPES else "summary" for x in raw]
        return out

    def plan_slide_types(self, title: str, outline: List[str], style: str, material: str) -> List[str]:
        total = len(outline)
        if not outline:
            return []

        if not self.use_mock:
            try:
                planned = self._llm_plan_slide_types(title, outline, style, material)
            except Exception:
                planned = [self._guess_slide_type(sec, i, total) for i, sec in enumerate(outline, start=1)]
        else:
            planned = [self._guess_slide_type(sec, i, total) for i, sec in enumerate(outline, start=1)]

        # Baseline constraints for more stable deck structure.
        if total >= 3 and "timeline" not in planned:
            planned[min(total - 1, 5)] = "timeline"
        if total >= 4 and "data" not in planned and re.search(r"\d|%|同比|环比|增长|下降", material):
            planned[min(total - 1, 4)] = "data"
        planned[0] = "title"
        return planned

    def _pick_evidence(self, material: str, section: str) -> List[str]:
        sentences = self._material_sentences(material)
        if not sentences:
            sentences = ["待补充资料内容"]

        tokens = [t for t in re.split(r"[\s/、,，]+", section) if len(t) >= 2]
        scored = []
        for s in sentences:
            score = 0
            for t in tokens:
                if t in s:
                    score += 1
            scored.append((score, s))

        scored.sort(key=lambda x: x[0], reverse=True)
        evidence = [s for _, s in scored[:4]]
        while len(evidence) < 4:
            evidence.append(sentences[len(evidence) % len(sentences)])
        return evidence[:4]

    def _ensure_conclusion_bullet(self, bullets: List[str]) -> List[str]:
        clean = [self._normalize(x) for x in bullets if self._normalize(x)]
        while len(clean) < 3:
            clean.append("待补充要点")

        first = clean[0]
        if not re.match(r"^(结论|takeaway|conclusion)\s*[:：]", first, flags=re.IGNORECASE):
            clean[0] = f"结论：{first}"
        return clean[:3]

    def _build_chart_data(self, section: str, evidence: List[str], material: str) -> Dict | None:
        text = "\n".join(evidence + self._material_sentences(material)[:12])
        nums = []
        units = []
        for m in re.finditer(r"(-?\d+(?:\.\d+)?)\s*(%|万元|万|亿|k|m|ms|天|人|次)?", text, flags=re.IGNORECASE):
            try:
                v = float(m.group(1))
            except ValueError:
                continue
            if abs(v) > 1_000_000_000:
                continue
            nums.append(abs(v))
            units.append((m.group(2) or "").strip())
            if len(nums) >= 4:
                break

        if len(nums) < 3:
            return None

        unit = ""
        for u in units:
            if u:
                unit = u
                break

        labels = [f"指标{i}" for i in range(1, len(nums) + 1)]
        if "同比" in section:
            labels[0] = "同比"
        if "环比" in section and len(labels) > 1:
            labels[1] = "环比"

        return {
            "labels": labels,
            "values": nums,
            "unit": unit,
        }

    def _llm_generate_slide(
        self,
        title: str,
        section: str,
        style: str,
        material: str,
        page: int,
        slide_type_hint: str = "summary",
    ) -> Dict:
        sys = "你是企业汇报PPT写作助手。返回JSON，不要解释。"
        usr = (
            "请为单页PPT生成内容，返回JSON："
            "{\"title\":\"\",\"bullets\":[\"\",\"\",\"\"],\"notes\":\"\","
            "\"slide_type\":\"title|risk|timeline|summary|data\","
            "\"evidence\":[\"\",\"\",\"\"],"
            "\"chart_data\":{\"labels\":[\"\"],\"values\":[0],\"unit\":\"\"}}"
            f"\n主题: {title}\n页面: 第{page}页\n小节: {section}\n风格: {style}\n"
            f"建议页型: {slide_type_hint}\n资料:\n{material}"
        )
        data = self.client.chat_json(sys, usr, temperature=0.6)
        bullets = [self._normalize(str(x)) for x in data.get("bullets", []) if self._normalize(str(x))]
        bullets = self._ensure_conclusion_bullet(bullets)
        evidence = [self._normalize(str(x)) for x in data.get("evidence", []) if self._normalize(str(x))]

        while len(evidence) < 3:
            evidence.append(bullets[len(evidence) % len(bullets)])

        slide_type = self._normalize(str(data.get("slide_type", slide_type_hint))).lower() or slide_type_hint
        if slide_type not in ALLOWED_SLIDE_TYPES:
            slide_type = slide_type_hint if slide_type_hint in ALLOWED_SLIDE_TYPES else "summary"

        chart_data = data.get("chart_data") if isinstance(data.get("chart_data"), dict) else None
        if slide_type == "data" and not chart_data:
            chart_data = self._build_chart_data(section, evidence, material)

        return {
            "title": self._normalize(str(data.get("title", f"{title} - {section}"))) or f"{title} - {section}",
            "bullets": bullets,
            "notes": self._normalize(str(data.get("notes", ""))) or "基于资料自动生成",
            "slide_type": slide_type,
            "evidence": evidence[:3],
            "chart_data": chart_data,
        }

    def generate_slide(
        self,
        title: str,
        section: str,
        style: str,
        material: str,
        page: int,
        slide_type_hint: str = "summary",
    ) -> Dict:
        if not self.use_mock:
            return self._llm_generate_slide(title, section, style, material, page, slide_type_hint)

        evidence = self._pick_evidence(material, section)
        prefixes = STYLE_PREFIX["management" if style == "management" else "technical"]
        seed = abs(hash(f"{title}|{section}|{style}|{page}|{evidence[0]}")) % (2**32)
        rng = random.Random(seed)

        candidates = [
            f"{prefixes[0]}：{evidence[0]}",
            f"{prefixes[1]}：{evidence[1]}",
            f"{prefixes[2]}：{evidence[2]}",
            f"补充：{evidence[3]}",
        ]
        rng.shuffle(candidates)

        inferred_type = slide_type_hint if slide_type_hint in ALLOWED_SLIDE_TYPES else self._guess_slide_type(section, page, 10)
        if page == 1:
            inferred_type = "title"

        chart_data = None
        if inferred_type == "data":
            chart_data = self._build_chart_data(section, evidence, material)
            if chart_data:
                vals = chart_data["values"]
                unit = chart_data.get("unit", "")
                max_v = max(vals)
                min_v = min(vals)
                delta = max_v - min_v
                candidates = [
                    f"结论：核心指标区间为 {min_v:.1f}{unit} ~ {max_v:.1f}{unit}",
                    f"波动：区间跨度 {delta:.1f}{unit}",
                    f"建议：聚焦高值指标并复用有效动作",
                ]

        bullets = self._ensure_conclusion_bullet(candidates[:3])
        return {
            "title": f"{title} - {section}",
            "bullets": bullets,
            "notes": f"资料依据：{evidence[0]}；{evidence[1]}",
            "slide_type": inferred_type,
            "evidence": evidence[:3],
            "chart_data": chart_data,
        }

    def _llm_rewrite_slide(self, slide: Dict, action: str) -> Dict:
        sys = "你是企业汇报改写助手。只返回JSON。"
        usr = (
            "请根据action改写当前PPT页文案，保持事实不变。返回JSON格式："
            "{\"title\":\"\",\"bullets\":[\"\",\"\",\"\"],\"notes\":\"\"}"
            f"\naction: {action}\nslide: {slide}"
        )
        data = self.client.chat_json(sys, usr, temperature=0.4)
        bullets = [self._normalize(str(x)) for x in data.get("bullets", []) if self._normalize(str(x))]
        bullets = self._ensure_conclusion_bullet(bullets)

        out = dict(slide)
        out["title"] = self._normalize(str(data.get("title", slide.get("title", "")))) or slide.get("title", "")
        out["bullets"] = bullets
        out["notes"] = self._normalize(str(data.get("notes", ""))) or slide.get("notes", "")
        return out

    def rewrite_slide(self, slide: Dict, action: str) -> Dict:
        if not self.use_mock:
            return self._llm_rewrite_slide(slide, action)

        result = dict(slide)
        bullets = list(slide.get("bullets", []))
        while len(bullets) < 3:
            bullets.append("补充要点")

        if action == "concise":
            result["bullets"] = [b[:26] + "..." if len(b) > 26 else b for b in bullets[:3]]
            result["notes"] = "更精简表达"
        elif action == "management":
            result["bullets"] = [
                f"结果：{bullets[0]}",
                f"风险：{bullets[1]}",
                f"决策：{bullets[2]}",
            ]
            result["notes"] = "切换为管理口径"
        elif action == "technical":
            result["bullets"] = [
                f"现状：{bullets[0]}",
                f"细节：{bullets[1]}",
                f"计划：{bullets[2]}",
            ]
            result["notes"] = "切换为技术口径"

        result["bullets"] = self._ensure_conclusion_bullet(result.get("bullets", bullets[:3]))
        return result
