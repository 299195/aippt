from __future__ import annotations

from datetime import datetime
from typing import Any, List, Literal, Optional

from pydantic import BaseModel, Field


StyleType = Literal["management", "technical"]
TemplateId = str
CreationType = Literal["idea", "outline"]


class GenerateResponse(BaseModel):
    job_id: str


class SlideDTO(BaseModel):
    page: int
    title: str
    bullets: List[str]
    notes: str
    slide_type: Optional[str] = None
    evidence: Optional[List[str]] = None


class JobDetailResponse(BaseModel):
    job_id: str
    status: str
    style: StyleType
    template_id: TemplateId = Field(default="a2p_2", pattern=r"^[a-z0-9_\\-]+$", min_length=2, max_length=80)
    title: str
    outline: List[str]
    slides: List[SlideDTO]
    pptx_url: Optional[str] = None
    created_at: datetime


class HistoryItem(BaseModel):
    job_id: str
    title: str
    style: StyleType
    template_id: TemplateId = Field(default="a2p_2", pattern=r"^[a-z0-9_\\-]+$", min_length=2, max_length=80)
    status: str
    created_at: datetime


class ModelConfigResponse(BaseModel):
    provider: str
    model: str
    endpoint_id: Optional[str] = None
    use_mock: bool
    configured: bool
    base_url: str


class GenerateRequest(BaseModel):
    title: str = Field(min_length=2, max_length=200)
    material_text: str = Field(default="")
    outline_text: str = Field(default="")
    outline: Optional[List[str]] = None
    style: StyleType = "management"
    template_id: TemplateId = Field(default="a2p_2", pattern=r"^[a-z0-9_\\-]+$", min_length=2, max_length=80)


class UploadParseResponse(BaseModel):
    extracted_text: str


class OutlinePreviewRequest(BaseModel):
    title: str = Field(min_length=2, max_length=200)
    material_text: str = Field(default="")
    outline_text: str = Field(default="")
    style: StyleType = "management"


class OutlinePreviewResponse(BaseModel):
    outline: List[str]
    outline_markdown: Optional[str] = None


class TemplateItem(BaseModel):
    id: TemplateId
    name: str
    subtitle: str
    summary: str
    preview_bg: str
    preview_fg: str
    preview_accent: str
    preview_image_url: Optional[str] = None


class ProjectCreateRequest(BaseModel):
    title: str = Field(min_length=2, max_length=200)
    material_text: str = Field(default="")
    outline_text: str = Field(default="")
    style: StyleType = "management"
    template_id: TemplateId = Field(default="a2p_2", pattern=r"^[a-z0-9_\\-]+$", min_length=2, max_length=80)
    creation_type: CreationType = "idea"


class ProjectCreateResponse(BaseModel):
    project_id: str
    status: str


class ProjectOutlineGenerateRequest(BaseModel):
    outline: Optional[List[str]] = None
    outline_markdown: Optional[str] = None


class TaskStartResponse(BaseModel):
    task_id: str


class TaskProgressDTO(BaseModel):
    total: int = 0
    completed: int = 0
    failed: int = 0
    current_step: Optional[str] = None


class TaskDTO(BaseModel):
    task_id: str
    project_id: str
    task_type: str
    status: str
    progress: TaskProgressDTO
    error_message: Optional[str] = None
    result: Optional[dict[str, Any]] = None
    created_at: datetime
    completed_at: Optional[datetime] = None


class OutlineContentDTO(BaseModel):
    title: str
    points: List[str] = []


class DescriptionContentDTO(BaseModel):
    title: str
    bullets: List[str]
    notes: str
    slide_type: Optional[str] = None
    evidence: Optional[List[str]] = None
    chart_data: Optional[dict[str, Any]] = None


class PageDTO(BaseModel):
    page_id: str
    order_index: int
    outline_content: OutlineContentDTO
    description_content: Optional[DescriptionContentDTO] = None
    status: str
    created_at: datetime
    updated_at: datetime


class ProjectDetailDTO(BaseModel):
    project_id: str
    title: str
    creation_type: CreationType
    idea_prompt: str
    outline_text: str
    material_text: str
    style: StyleType
    template_id: TemplateId
    status: str
    pptx_url: Optional[str] = None
    pages: List[PageDTO]
    created_at: datetime
    updated_at: datetime


class ProjectListItemDTO(BaseModel):
    project_id: str
    title: str
    style: StyleType
    template_id: TemplateId
    status: str
    created_at: datetime
    updated_at: datetime


