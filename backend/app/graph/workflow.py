from __future__ import annotations

import json
import re
from datetime import datetime
from typing import Dict, List, TypedDict
from uuid import uuid4

from langgraph.graph import END, StateGraph

from app.config import settings
from app.services.llm import LLMAdapter
from app.services.pptx_exporter import export_slides_to_pptx
from app.storage.db import upsert_job


class PPTState(TypedDict, total=False):
    job_id: str
    title: str
    style: str
    template_id: str
    target_pages: int
    material: str
    outline: List[str]
    slides: List[Dict]
    rewrite_action: str
    pptx_url: str
    status: str
    created_at: str
    qc_issues: List[str]
    is_rewrite: bool
    has_outline: bool


llm = LLMAdapter(use_mock=settings.use_mock_llm)


def _section_name(raw_outline: str) -> str:
    parts = re.split(r"[:：]", raw_outline, maxsplit=1)
    if len(parts) == 2:
        return parts[1].strip()
    return raw_outline.strip()


def _has_conclusion_lead(bullets: List[str]) -> bool:
    if not bullets:
        return False
    first = str(bullets[0]).strip().lower()
    return first.startswith("结论:") or first.startswith("结论：") or first.startswith("conclusion:") or first.startswith("takeaway:")


def parse_input_node(state: PPTState) -> PPTState:
    is_rewrite = bool(state.get("rewrite_action") and state.get("slides"))
    has_outline = bool(state.get("outline"))
    return {"status": "parsed", "is_rewrite": is_rewrite, "has_outline": has_outline}


def route_after_parse(state: PPTState) -> str:
    if state.get("is_rewrite"):
        return "style"
    if state.get("has_outline"):
        return "fill"
    return "outline"


def outline_node(state: PPTState) -> PPTState:
    outline = llm.generate_outline(state["title"], state["style"], state.get("material", ""), state["target_pages"])
    return {"outline": outline, "status": "outlined"}


def content_fill_node(state: PPTState) -> PPTState:
    outline = state.get("outline", [])
    slide_types = llm.plan_slide_types(
        state["title"],
        outline,
        state["style"],
        state.get("material", ""),
    )

    slides: List[Dict] = []
    for i, section in enumerate(outline, start=1):
        section_name = _section_name(section)
        slide_type_hint = slide_types[i - 1] if i - 1 < len(slide_types) else "summary"
        payload = llm.generate_slide(
            state["title"],
            section_name,
            state["style"],
            state.get("material", ""),
            i,
            slide_type_hint,
        )
        payload["page"] = i
        payload["slide_type"] = slide_type_hint if payload.get("slide_type") == "title" and i != 1 else payload.get("slide_type", slide_type_hint)
        slides.append(payload)

    return {"slides": slides, "status": "filled"}


def quality_gate_node(state: PPTState) -> PPTState:
    issues: List[str] = []
    slides = state.get("slides", [])
    if len(slides) < 8:
        issues.append("页数不足8页")

    for idx, slide in enumerate(slides, start=1):
        bullets = slide.get("bullets", [])
        if len(bullets) < 3:
            issues.append("第%d页要点不足3条" % idx)
        if not slide.get("title"):
            issues.append("第%d页缺少标题" % idx)
        if not _has_conclusion_lead(bullets):
            issues.append("第%d页缺少结论句" % idx)

        if slide.get("slide_type") == "data":
            chart_data = slide.get("chart_data")
            if not isinstance(chart_data, dict) or not chart_data.get("values"):
                issues.append("第%d页数据页缺少图表数据" % idx)

    status = "qc_pass" if not issues else "qc_failed"
    return {"status": status, "qc_issues": issues}


def route_after_quality(state: PPTState) -> str:
    return "style" if state.get("status") == "qc_pass" else "repair"


def repair_node(state: PPTState) -> PPTState:
    repaired = []
    for slide in state.get("slides", []):
        fixed = dict(slide)
        bullets = [str(x).strip() for x in list(fixed.get("bullets", [])) if str(x).strip()]
        while len(bullets) < 3:
            bullets.append("补充要点：请补充指标、风险与下一步")
        if not _has_conclusion_lead(bullets):
            bullets[0] = f"结论：{bullets[0]}"
        fixed["bullets"] = bullets[:3]

        if fixed.get("slide_type") == "data":
            chart_data = fixed.get("chart_data")
            if not isinstance(chart_data, dict) or not chart_data.get("values"):
                fixed["chart_data"] = {
                    "labels": ["指标1", "指标2", "指标3"],
                    "values": [72.0, 64.0, 81.0],
                    "unit": "%",
                }
        repaired.append(fixed)
    return {"slides": repaired, "status": "repaired"}


def style_adapt_node(state: PPTState) -> PPTState:
    action = state.get("rewrite_action", "")
    if not action:
        return {"status": "styled"}

    slides = []
    for slide in state.get("slides", []):
        rewritten = llm.rewrite_slide(slide, action)
        rewritten["page"] = slide.get("page")
        rewritten["slide_type"] = slide.get("slide_type", rewritten.get("slide_type", "summary"))
        rewritten["evidence"] = slide.get("evidence", [])
        rewritten["chart_data"] = slide.get("chart_data")
        slides.append(rewritten)

    style = state.get("style", "management")
    if action in ("management", "technical"):
        style = action
    return {"slides": slides, "style": style, "status": "styled"}


def export_node(state: PPTState) -> PPTState:
    job_id = state.get("job_id") or str(uuid4())
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{job_id}_{ts}.pptx"
    out_path = settings.export_dir / filename
    exported = export_slides_to_pptx(
        state.get("slides", []),
        out_path,
        state.get("template_id", "no_template"),
        state.get("title", ""),
        state.get("outline", []),
    )
    return {
        "job_id": job_id,
        "pptx_url": f"/exports/{exported}",
        "status": "exported",
        "created_at": datetime.utcnow().isoformat(),
    }


def persist_node(state: PPTState) -> PPTState:
    upsert_job(
        {
            "job_id": state["job_id"],
            "title": state["title"],
            "style": state["style"],
            "template_id": state.get("template_id", "no_template"),
            "status": "done",
            "outline_json": json.dumps(state.get("outline", []), ensure_ascii=False),
            "slides_json": json.dumps(state.get("slides", []), ensure_ascii=False),
            "parsed_json": json.dumps({}, ensure_ascii=False),
            "material_text": state.get("material", ""),
            "pptx_url": state.get("pptx_url", ""),
            "created_at": state["created_at"],
        }
    )
    return {"status": "done"}


def build_graph():
    graph = StateGraph(PPTState)
    graph.add_node("parse", parse_input_node)
    graph.add_node("outline", outline_node)
    graph.add_node("fill", content_fill_node)
    graph.add_node("quality", quality_gate_node)
    graph.add_node("repair", repair_node)
    graph.add_node("style", style_adapt_node)
    graph.add_node("export", export_node)
    graph.add_node("persist", persist_node)

    graph.set_entry_point("parse")
    graph.add_conditional_edges("parse", route_after_parse, {"style": "style", "outline": "outline", "fill": "fill"})
    graph.add_edge("outline", "fill")
    graph.add_edge("fill", "quality")
    graph.add_conditional_edges("quality", route_after_quality, {"style": "style", "repair": "repair"})
    graph.add_edge("repair", "style")
    graph.add_edge("style", "export")
    graph.add_edge("export", "persist")
    graph.add_edge("persist", END)
    return graph.compile()


workflow = build_graph()


def run_generation(state: PPTState) -> PPTState:
    return workflow.invoke(state)

