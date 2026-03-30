from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Literal


StyleType = Literal["management", "technical"]


@dataclass
class Slide:
    page: int
    title: str
    bullets: list[str]
    notes: str


@dataclass
class JobResult:
    job_id: str
    style: StyleType
    outline: list[str]
    slides: list[Slide]
    pptx_path: str
    created_at: datetime

