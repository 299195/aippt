
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Any, List
from uuid import uuid4

from fastapi import APIRouter, File, HTTPException, UploadFile
from fastapi.responses import StreamingResponse

from app.config import settings
from app.schemas import (
    GenerateRequest,
    GenerateResponse,
    HistoryItem,
    JobDetailResponse,
    ModelConfigResponse,
    OutlinePreviewRequest,
    OutlinePreviewResponse,
    PageDTO,
    ProjectCreateRequest,
    ProjectCreateResponse,
    ProjectDetailDTO,
    ProjectListItemDTO,
    ProjectOutlineGenerateRequest,
    SlideDTO,
    TaskDTO,
    TaskProgressDTO,
    TaskStartResponse,
    TemplateItem,
    UploadParseResponse,
)
from app.services.new_backend_workflow import (
    clean_outline_items,
    generate_descriptions_task,
    generate_ppt_task,
    get_outline_for_project,
    llm,
    rebuild_project_pages,
    stream_generate_descriptions_events,
    stream_outline_preview_events,
    utc_now_iso,
)
from app.services.parser import parse_text_input, read_uploaded_file
from app.services.task_manager import task_manager
from app.services.template_catalog import list_templates, template_exists
from app.storage.db import (
    create_project as db_create_project,
    create_task,
    get_project,
    get_project_task,
    get_task,
    list_pages,
    list_projects,
    make_progress,
)


router = APIRouter()


def _parse_json(raw: str | None, fallback: Any) -> Any:
    if raw is None:
        return fallback
    try:
        return json.loads(raw)
    except Exception:
        return fallback


def _as_dt(raw: str | None) -> datetime:
    if not raw:
        return datetime.utcnow()
    try:
        return datetime.fromisoformat(raw)
    except Exception:
        return datetime.utcnow()


def _normalize_style(style: str | None) -> str:
    return "technical" if str(style or "").lower() == "technical" else "management"


def _row_to_page_dto(row: Any) -> PageDTO:
    outline_content = _parse_json(row["outline_content"], {})
    description_content = _parse_json(row["description_content"], None)
    return PageDTO(
        page_id=str(row["page_id"]),
        order_index=int(row["order_index"]),
        outline_content={
            "title": str(outline_content.get("title") or ""),
            "points": [str(x) for x in list(outline_content.get("points") or [])],
        },
        description_content=description_content,
        status=str(row["status"]),
        created_at=_as_dt(row["created_at"]),
        updated_at=_as_dt(row["updated_at"]),
    )


def _get_project_detail_or_404(project_id: str) -> ProjectDetailDTO:
    project = get_project(project_id)
    if not project:
        raise HTTPException(status_code=404, detail="project not found")

    pages = [_row_to_page_dto(x) for x in list_pages(project_id)]
    return ProjectDetailDTO(
        project_id=str(project["project_id"]),
        title=str(project["title"]),
        creation_type=str(project["creation_type"]),
        idea_prompt=str(project["idea_prompt"] or ""),
        outline_text=str(project["outline_text"] or ""),
        material_text=str(project["material_text"] or ""),
        style=_normalize_style(str(project["style"])),
        template_id=str(project["template_id"]),
        status=str(project["status"]),
        pptx_url=project["pptx_url"],
        pages=pages,
        created_at=_as_dt(project["created_at"]),
        updated_at=_as_dt(project["updated_at"]),
    )


def _project_to_job_detail(project_id: str) -> JobDetailResponse:
    detail = _get_project_detail_or_404(project_id)
    slides: list[SlideDTO] = []
    outline: list[str] = []

    for page in detail.pages:
        outline.append(page.outline_content.title)
        if page.description_content:
            dc = page.description_content
            slides.append(
                SlideDTO(
                    page=page.order_index + 1,
                    title=str(dc.title),
                    bullets=[str(x) for x in dc.bullets],
                    notes=str(dc.notes),
                    slide_type=dc.slide_type,
                    evidence=dc.evidence,
                )
            )

    return JobDetailResponse(
        job_id=detail.project_id,
        status=detail.status,
        style=detail.style,
        template_id=detail.template_id,
        title=detail.title,
        outline=outline,
        slides=slides,
        pptx_url=detail.pptx_url,
        created_at=detail.created_at,
    )


def _create_project_row(req: ProjectCreateRequest) -> str:
    if not template_exists(req.template_id):
        raise HTTPException(status_code=400, detail="template not found")

    project_id = str(uuid4())
    now = utc_now_iso()
    db_create_project(
        {
            "project_id": project_id,
            "title": req.title,
            "creation_type": req.creation_type,
            "idea_prompt": req.title,
            "outline_text": req.outline_text,
            "material_text": req.material_text,
            "style": _normalize_style(req.style),
            "template_id": req.template_id,
            "target_pages": 0,
            "status": "DRAFT",
            "pptx_url": None,
            "created_at": now,
            "updated_at": now,
        }
    )
    return project_id


def _create_task(project_id: str, task_type: str, total: int = 0) -> str:
    task_id = str(uuid4())
    create_task(
        {
            "task_id": task_id,
            "project_id": project_id,
            "task_type": task_type,
            "status": "PENDING",
            "progress_json": make_progress(total, 0, 0, "queued"),
            "error_message": None,
            "result_json": None,
            "created_at": utc_now_iso(),
            "completed_at": None,
        }
    )
    return task_id


@router.get("/model/config", response_model=ModelConfigResponse)
def model_config() -> ModelConfigResponse:
    endpoint_id = str(getattr(settings, "model_endpoint_id", "") or "").strip()
    model = endpoint_id or settings.model_name
    return ModelConfigResponse(
        provider=settings.model_provider,
        model=model,
        endpoint_id=endpoint_id or None,
        use_mock=settings.use_mock_llm,
        configured=bool(settings.model_base_url and settings.model_api_key and model),
        base_url=settings.model_base_url,
    )


@router.get("/templates", response_model=List[TemplateItem])
def templates() -> List[TemplateItem]:
    return [TemplateItem(**item) for item in list_templates()]


@router.post("/parse-upload", response_model=UploadParseResponse)
async def parse_upload(file: UploadFile = File(...)) -> UploadParseResponse:
    suffix = Path(file.filename or "").suffix.lower()
    if suffix not in {".md", ".docx"}:
        raise HTTPException(status_code=400, detail="only .md/.docx are supported")

    with NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(await file.read())
        tmp_path = Path(tmp.name)

    try:
        extracted = read_uploaded_file(tmp_path)
        return UploadParseResponse(extracted_text=extracted)
    finally:
        tmp_path.unlink(missing_ok=True)

@router.post("/outline/preview", response_model=OutlinePreviewResponse)
def preview_outline(req: OutlinePreviewRequest) -> OutlinePreviewResponse:
    material = parse_text_input(req.title, req.outline_text, req.material_text)
    bundle = llm.generate_outline_bundle(req.title, _normalize_style(req.style), material)
    return OutlinePreviewResponse(
        outline=[str(x) for x in list(bundle.get("outline_titles") or [])],
        outline_markdown=str(bundle.get("outline_markdown") or ""),
    )


@router.post("/outline/preview/stream")
def preview_outline_stream(req: OutlinePreviewRequest) -> StreamingResponse:
    material = parse_text_input(req.title, req.outline_text, req.material_text)

    def _iter() -> Any:
        try:
            for event in stream_outline_preview_events(
                title=req.title,
                style=_normalize_style(req.style),
                material_text=material,
            ):
                yield f"data: {json.dumps(event, ensure_ascii=False)}\n\n"
        except Exception as exc:
            payload = {"type": "error", "message": str(exc)}
            yield f"data: {json.dumps(payload, ensure_ascii=False)}\n\n"

    return StreamingResponse(_iter(), media_type="text/event-stream")


@router.post("/projects", response_model=ProjectCreateResponse)
def create_project(req: ProjectCreateRequest) -> ProjectCreateResponse:
    project_id = _create_project_row(req)
    return ProjectCreateResponse(project_id=project_id, status="DRAFT")


@router.get("/projects", response_model=List[ProjectListItemDTO])
def project_history() -> List[ProjectListItemDTO]:
    rows = list_projects(100)
    return [
        ProjectListItemDTO(
            project_id=str(r["project_id"]),
            title=str(r["title"]),
            style=_normalize_style(str(r["style"])),
            template_id=str(r["template_id"]),
            status=str(r["status"]),
            created_at=_as_dt(r["created_at"]),
            updated_at=_as_dt(r["updated_at"]),
        )
        for r in rows
    ]


@router.get("/projects/{project_id}", response_model=ProjectDetailDTO)
def project_detail(project_id: str) -> ProjectDetailDTO:
    return _get_project_detail_or_404(project_id)


@router.post("/projects/{project_id}/generate/outline", response_model=ProjectDetailDTO)
def generate_project_outline(project_id: str, req: ProjectOutlineGenerateRequest) -> ProjectDetailDTO:
    project = get_project(project_id)
    if not project:
        raise HTTPException(status_code=404, detail="project not found")

    requested_outline = clean_outline_items(req.outline or [])
    outline_pages, outline_markdown = get_outline_for_project(
        project,
        requested_outline if requested_outline else None,
        str(req.outline_markdown or "").strip() or None,
    )
    if not outline_pages:
        raise HTTPException(status_code=400, detail="outline is empty")

    rebuild_project_pages(project_id, outline_pages, outline_markdown)
    return _get_project_detail_or_404(project_id)


@router.post("/projects/{project_id}/generate/descriptions", response_model=TaskStartResponse)
def start_descriptions(project_id: str) -> TaskStartResponse:
    project = get_project(project_id)
    if not project:
        raise HTTPException(status_code=404, detail="project not found")

    pages = list_pages(project_id)
    if not pages:
        raise HTTPException(status_code=400, detail="please generate outline first")

    task_id = _create_task(project_id, "GENERATE_DESCRIPTIONS", len(pages))
    task_manager.submit_task(task_id, generate_descriptions_task, project_id)
    return TaskStartResponse(task_id=task_id)


@router.post("/projects/{project_id}/generate/descriptions/stream")
def stream_descriptions(project_id: str) -> StreamingResponse:
    project = get_project(project_id)
    if not project:
        raise HTTPException(status_code=404, detail="project not found")

    pages = list_pages(project_id)
    if not pages:
        raise HTTPException(status_code=400, detail="please generate outline first")

    def _iter() -> Any:
        try:
            for event in stream_generate_descriptions_events(project_id):
                yield f"data: {json.dumps(event, ensure_ascii=False)}\n\n"
        except Exception as exc:
            payload = {"type": "error", "message": str(exc)}
            yield f"data: {json.dumps(payload, ensure_ascii=False)}\n\n"

    return StreamingResponse(_iter(), media_type="text/event-stream")


@router.post("/projects/{project_id}/generate/ppt", response_model=TaskStartResponse)
def start_generate_ppt(project_id: str) -> TaskStartResponse:
    project = get_project(project_id)
    if not project:
        raise HTTPException(status_code=404, detail="project not found")

    pages = list_pages(project_id)
    if not pages:
        raise HTTPException(status_code=400, detail="please generate outline first")

    task_id = _create_task(project_id, "GENERATE_PPT", len(pages))
    task_manager.submit_task(task_id, generate_ppt_task, project_id)
    return TaskStartResponse(task_id=task_id)


@router.get("/projects/{project_id}/tasks/{task_id}", response_model=TaskDTO)
def project_task_detail(project_id: str, task_id: str) -> TaskDTO:
    row = get_project_task(project_id, task_id)
    if not row:
        raise HTTPException(status_code=404, detail="task not found")

    progress = _parse_json(row["progress_json"], {})
    return TaskDTO(
        task_id=str(row["task_id"]),
        project_id=str(row["project_id"]),
        task_type=str(row["task_type"]),
        status=str(row["status"]),
        progress=TaskProgressDTO(
            total=int(progress.get("total", 0)),
            completed=int(progress.get("completed", 0)),
            failed=int(progress.get("failed", 0)),
            current_step=progress.get("current_step"),
        ),
        error_message=row["error_message"],
        result=_parse_json(row["result_json"], None),
        created_at=_as_dt(row["created_at"]),
        completed_at=_as_dt(row["completed_at"]) if row["completed_at"] else None,
    )


@router.get("/tasks/{task_id}", response_model=TaskDTO)
def global_task_detail(task_id: str) -> TaskDTO:
    row = get_task(task_id)
    if not row:
        raise HTTPException(status_code=404, detail="task not found")

    progress = _parse_json(row["progress_json"], {})
    return TaskDTO(
        task_id=str(row["task_id"]),
        project_id=str(row["project_id"]),
        task_type=str(row["task_type"]),
        status=str(row["status"]),
        progress=TaskProgressDTO(
            total=int(progress.get("total", 0)),
            completed=int(progress.get("completed", 0)),
            failed=int(progress.get("failed", 0)),
            current_step=progress.get("current_step"),
        ),
        error_message=row["error_message"],
        result=_parse_json(row["result_json"], None),
        created_at=_as_dt(row["created_at"]),
        completed_at=_as_dt(row["completed_at"]) if row["completed_at"] else None,
    )

@router.post("/jobs", response_model=GenerateResponse)
def create_job(req: GenerateRequest) -> GenerateResponse:
    project_id = _create_project_row(
        ProjectCreateRequest(
            title=req.title,
            material_text=req.material_text,
            outline_text=req.outline_text,
            style=req.style,
            template_id=req.template_id,
            creation_type="idea",
        )
    )

    project = get_project(project_id)
    if not project:
        raise HTTPException(status_code=500, detail="project create failed")

    requested_outline = clean_outline_items(req.outline or [])
    outline_pages, outline_markdown = get_outline_for_project(
        project,
        requested_outline if requested_outline else None,
        req.outline_text.strip() or None,
    )
    rebuild_project_pages(project_id, outline_pages, outline_markdown)

    desc_task_id = _create_task(project_id, "GENERATE_DESCRIPTIONS", len(outline_pages))
    generate_descriptions_task(desc_task_id, project_id)
    desc_task = get_task(desc_task_id)
    if desc_task and str(desc_task["status"]) == "FAILED":
        raise HTTPException(status_code=500, detail=str(desc_task["error_message"] or "description generation failed"))

    ppt_task_id = _create_task(project_id, "GENERATE_PPT", len(outline_pages))
    generate_ppt_task(ppt_task_id, project_id)
    ppt_task = get_task(ppt_task_id)
    if ppt_task and str(ppt_task["status"]) == "FAILED":
        raise HTTPException(status_code=500, detail=str(ppt_task["error_message"] or "ppt export failed"))

    return GenerateResponse(job_id=project_id)


@router.get("/jobs/{job_id}", response_model=JobDetailResponse)
def job_detail(job_id: str) -> JobDetailResponse:
    return _project_to_job_detail(job_id)


@router.get("/jobs", response_model=List[HistoryItem])
def history() -> List[HistoryItem]:
    rows = list_projects(100)
    return [
        HistoryItem(
            job_id=str(r["project_id"]),
            title=str(r["title"]),
            style=_normalize_style(str(r["style"])),
            template_id=str(r["template_id"]),
            status=str(r["status"]),
            created_at=_as_dt(r["created_at"]),
        )
        for r in rows
    ]

