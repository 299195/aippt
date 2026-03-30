from __future__ import annotations

import json
import sqlite3
from typing import Any

from app.config import settings


JOBS_DDL = """
CREATE TABLE IF NOT EXISTS jobs (
    job_id TEXT PRIMARY KEY,
    title TEXT NOT NULL,
    style TEXT NOT NULL,
    template_id TEXT NOT NULL DEFAULT 'no_template',
    status TEXT NOT NULL,
    outline_json TEXT NOT NULL,
    slides_json TEXT NOT NULL,
    parsed_json TEXT NOT NULL DEFAULT '{}',
    material_text TEXT NOT NULL DEFAULT '',
    pptx_url TEXT,
    created_at TEXT NOT NULL
);
"""

PROJECTS_DDL = """
CREATE TABLE IF NOT EXISTS projects (
    project_id TEXT PRIMARY KEY,
    title TEXT NOT NULL,
    creation_type TEXT NOT NULL DEFAULT 'idea',
    idea_prompt TEXT NOT NULL DEFAULT '',
    outline_text TEXT NOT NULL DEFAULT '',
    material_text TEXT NOT NULL DEFAULT '',
    style TEXT NOT NULL DEFAULT 'management',
    template_id TEXT NOT NULL DEFAULT 'no_template',
    target_pages INTEGER NOT NULL DEFAULT 8,
    status TEXT NOT NULL DEFAULT 'DRAFT',
    pptx_url TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);
"""

PAGES_DDL = """
CREATE TABLE IF NOT EXISTS pages (
    page_id TEXT PRIMARY KEY,
    project_id TEXT NOT NULL,
    order_index INTEGER NOT NULL,
    outline_content TEXT NOT NULL,
    description_content TEXT,
    status TEXT NOT NULL DEFAULT 'DRAFT',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    FOREIGN KEY(project_id) REFERENCES projects(project_id) ON DELETE CASCADE
);
"""

TASKS_DDL = """
CREATE TABLE IF NOT EXISTS tasks (
    task_id TEXT PRIMARY KEY,
    project_id TEXT NOT NULL,
    task_type TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'PENDING',
    progress_json TEXT NOT NULL DEFAULT '{"total":0,"completed":0,"failed":0}',
    error_message TEXT,
    result_json TEXT,
    created_at TEXT NOT NULL,
    completed_at TEXT,
    FOREIGN KEY(project_id) REFERENCES projects(project_id) ON DELETE CASCADE
);
"""

PAGES_INDEX_DDL = "CREATE INDEX IF NOT EXISTS idx_pages_project_order ON pages(project_id, order_index)"
TASKS_INDEX_DDL = "CREATE INDEX IF NOT EXISTS idx_tasks_project_created ON tasks(project_id, datetime(created_at) DESC)"


def get_conn() -> sqlite3.Connection:
    settings.data_dir.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(settings.database_path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def _ensure_column(conn: sqlite3.Connection, table: str, column: str, ddl: str) -> None:
    rows = conn.execute("PRAGMA table_info(%s)" % table).fetchall()
    cols = [r[1] for r in rows]
    if column not in cols:
        conn.execute(ddl)


def init_db() -> None:
    with get_conn() as conn:
        conn.execute(JOBS_DDL)
        conn.execute(PROJECTS_DDL)
        conn.execute(PAGES_DDL)
        conn.execute(TASKS_DDL)
        conn.execute(PAGES_INDEX_DDL)
        conn.execute(TASKS_INDEX_DDL)
        _ensure_column(
            conn,
            "jobs",
            "template_id",
            "ALTER TABLE jobs ADD COLUMN template_id TEXT NOT NULL DEFAULT 'no_template'",
        )
        _ensure_column(
            conn,
            "jobs",
            "material_text",
            "ALTER TABLE jobs ADD COLUMN material_text TEXT NOT NULL DEFAULT ''",
        )
        _ensure_column(
            conn,
            "jobs",
            "parsed_json",
            "ALTER TABLE jobs ADD COLUMN parsed_json TEXT NOT NULL DEFAULT '{}'",
        )
        conn.commit()


def upsert_job(row: dict[str, Any]) -> None:
    sql = """
    INSERT INTO jobs (job_id, title, style, template_id, status, outline_json, slides_json, parsed_json, material_text, pptx_url, created_at)
    VALUES (:job_id, :title, :style, :template_id, :status, :outline_json, :slides_json, :parsed_json, :material_text, :pptx_url, :created_at)
    ON CONFLICT(job_id) DO UPDATE SET
        title=excluded.title,
        style=excluded.style,
        template_id=excluded.template_id,
        status=excluded.status,
        outline_json=excluded.outline_json,
        slides_json=excluded.slides_json,
        parsed_json=excluded.parsed_json,
        material_text=excluded.material_text,
        pptx_url=excluded.pptx_url,
        created_at=excluded.created_at;
    """
    with get_conn() as conn:
        conn.execute(sql, row)
        conn.commit()


def get_job(job_id: str) -> sqlite3.Row | None:
    with get_conn() as conn:
        return conn.execute("SELECT * FROM jobs WHERE job_id = ?", (job_id,)).fetchone()


def list_jobs(limit: int = 50) -> list[sqlite3.Row]:
    with get_conn() as conn:
        return conn.execute(
            "SELECT * FROM jobs ORDER BY datetime(created_at) DESC LIMIT ?",
            (limit,),
        ).fetchall()


def create_project(row: dict[str, Any]) -> None:
    sql = """
    INSERT INTO projects (
        project_id, title, creation_type, idea_prompt, outline_text, material_text,
        style, template_id, target_pages, status, pptx_url, created_at, updated_at
    )
    VALUES (
        :project_id, :title, :creation_type, :idea_prompt, :outline_text, :material_text,
        :style, :template_id, :target_pages, :status, :pptx_url, :created_at, :updated_at
    );
    """
    with get_conn() as conn:
        conn.execute(sql, row)
        conn.commit()


def update_project(project_id: str, fields: dict[str, Any]) -> None:
    if not fields:
        return
    assignments = ", ".join([f"{key} = :{key}" for key in fields.keys()])
    sql = f"UPDATE projects SET {assignments} WHERE project_id = :project_id"
    params = dict(fields)
    params["project_id"] = project_id
    with get_conn() as conn:
        conn.execute(sql, params)
        conn.commit()


def get_project(project_id: str) -> sqlite3.Row | None:
    with get_conn() as conn:
        return conn.execute("SELECT * FROM projects WHERE project_id = ?", (project_id,)).fetchone()


def list_projects(limit: int = 50) -> list[sqlite3.Row]:
    with get_conn() as conn:
        return conn.execute(
            "SELECT * FROM projects ORDER BY datetime(updated_at) DESC LIMIT ?",
            (limit,),
        ).fetchall()


def delete_project(project_id: str) -> None:
    with get_conn() as conn:
        conn.execute("DELETE FROM projects WHERE project_id = ?", (project_id,))
        conn.commit()


def replace_pages(project_id: str, pages: list[dict[str, Any]]) -> None:
    sql = """
    INSERT INTO pages (
        page_id, project_id, order_index, outline_content, description_content,
        status, created_at, updated_at
    )
    VALUES (
        :page_id, :project_id, :order_index, :outline_content, :description_content,
        :status, :created_at, :updated_at
    );
    """
    with get_conn() as conn:
        conn.execute("DELETE FROM pages WHERE project_id = ?", (project_id,))
        if pages:
            conn.executemany(sql, pages)
        conn.commit()


def list_pages(project_id: str) -> list[sqlite3.Row]:
    with get_conn() as conn:
        return conn.execute(
            "SELECT * FROM pages WHERE project_id = ? ORDER BY order_index ASC",
            (project_id,),
        ).fetchall()


def get_page(page_id: str) -> sqlite3.Row | None:
    with get_conn() as conn:
        return conn.execute("SELECT * FROM pages WHERE page_id = ?", (page_id,)).fetchone()


def update_page(page_id: str, fields: dict[str, Any]) -> None:
    if not fields:
        return
    assignments = ", ".join([f"{key} = :{key}" for key in fields.keys()])
    sql = f"UPDATE pages SET {assignments} WHERE page_id = :page_id"
    params = dict(fields)
    params["page_id"] = page_id
    with get_conn() as conn:
        conn.execute(sql, params)
        conn.commit()


def create_task(row: dict[str, Any]) -> None:
    sql = """
    INSERT INTO tasks (
        task_id, project_id, task_type, status, progress_json, error_message,
        result_json, created_at, completed_at
    )
    VALUES (
        :task_id, :project_id, :task_type, :status, :progress_json, :error_message,
        :result_json, :created_at, :completed_at
    );
    """
    with get_conn() as conn:
        conn.execute(sql, row)
        conn.commit()


def update_task(task_id: str, fields: dict[str, Any]) -> None:
    if not fields:
        return
    assignments = ", ".join([f"{key} = :{key}" for key in fields.keys()])
    sql = f"UPDATE tasks SET {assignments} WHERE task_id = :task_id"
    params = dict(fields)
    params["task_id"] = task_id
    with get_conn() as conn:
        conn.execute(sql, params)
        conn.commit()


def get_task(task_id: str) -> sqlite3.Row | None:
    with get_conn() as conn:
        return conn.execute("SELECT * FROM tasks WHERE task_id = ?", (task_id,)).fetchone()


def get_project_task(project_id: str, task_id: str) -> sqlite3.Row | None:
    with get_conn() as conn:
        return conn.execute(
            "SELECT * FROM tasks WHERE task_id = ? AND project_id = ?",
            (task_id, project_id),
        ).fetchone()


def list_project_tasks(project_id: str, limit: int = 20) -> list[sqlite3.Row]:
    with get_conn() as conn:
        return conn.execute(
            "SELECT * FROM tasks WHERE project_id = ? ORDER BY datetime(created_at) DESC LIMIT ?",
            (project_id, limit),
        ).fetchall()


def make_progress(total: int, completed: int = 0, failed: int = 0, current_step: str = "") -> str:
    payload = {
        "total": total,
        "completed": completed,
        "failed": failed,
    }
    if current_step:
        payload["current_step"] = current_step
    return json.dumps(payload, ensure_ascii=False)

