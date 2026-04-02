const DEFAULT_API_BASE = "http://127.0.0.1:8001/api";
const DEFAULT_FILE_BASE = "http://127.0.0.1:8001";
function normalizeBase(rawValue, fallbackBase, fallbackPath = "") {
    if (!rawValue) {
        return fallbackBase;
    }
    const raw = String(rawValue).trim();
    try {
        const parsed = new URL(raw);
        if (parsed.protocol !== "http:" && parsed.protocol !== "https:") {
            throw new Error("Unsupported protocol");
        }
        return `${parsed.origin}${parsed.pathname}`.replace(/\/+$/, "");
    }
    catch {
        const prefix = raw.match(/^(https?:\/\/[A-Za-z0-9.-]+:\d{2,5})/i);
        if (prefix) {
            return `${prefix[1]}${fallbackPath}`;
        }
        return fallbackBase;
    }
}
const API_BASE = normalizeBase(import.meta.env.VITE_API_BASE, DEFAULT_API_BASE, "/api");
const FILE_BASE = normalizeBase(import.meta.env.VITE_FILE_BASE, DEFAULT_FILE_BASE, "");
export const fileUrl = (url) => (url ? `${FILE_BASE}${url}` : "");
async function readError(res, fallback) {
    try {
        const data = (await res.json());
        throw new Error(data.detail || fallback);
    }
    catch {
        throw new Error(fallback);
    }
}
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
async function pollTask(projectId, taskId, timeoutMs = 8 * 60 * 1000) {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
        const res = await fetch(`${API_BASE}/projects/${projectId}/tasks/${taskId}`);
        if (!res.ok)
            return readError(res, `任务状态查询失败: HTTP ${res.status}`);
        const task = (await res.json());
        if (task.status === "COMPLETED")
            return task;
        if (task.status === "FAILED") {
            throw new Error(task.error_message || "任务执行失败");
        }
        await sleep(1200);
    }
    throw new Error("任务超时，请稍后在历史记录中查看");
}
function mapProjectToJob(project) {
    const sorted = [...(project.pages || [])].sort((a, b) => a.order_index - b.order_index);
    return {
        job_id: project.project_id,
        status: project.status,
        style: project.style,
        template_id: project.template_id,
        title: project.title,
        outline: sorted.map((p) => p.outline_content?.title || `第${p.order_index + 1}页`),
        slides: sorted
            .filter((p) => p.description_content)
            .map((p) => ({
            page: p.order_index + 1,
            title: p.description_content?.title || p.outline_content?.title || `第${p.order_index + 1}页`,
            bullets: p.description_content?.bullets || [],
            notes: p.description_content?.notes || "",
            slide_type: p.description_content?.slide_type,
            evidence: p.description_content?.evidence,
        })),
        pptx_url: project.pptx_url ?? null,
        created_at: project.created_at,
    };
}
export async function getModelConfig() {
    const res = await fetch(`${API_BASE}/model/config`);
    if (!res.ok)
        return readError(res, `模型配置获取失败: HTTP ${res.status}`);
    return (await res.json());
}
export async function getTemplates() {
    const res = await fetch(`${API_BASE}/templates`);
    if (!res.ok)
        return readError(res, `模板列表获取失败: HTTP ${res.status}`);
    return (await res.json());
}
export async function previewOutline(payload) {
    const res = await fetch(`${API_BASE}/outline/preview`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
    });
    if (!res.ok)
        return readError(res, `大纲预览失败: HTTP ${res.status}`);
    return (await res.json());
}
export async function createJob(payload) {
    const createRes = await fetch(`${API_BASE}/projects`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
            title: payload.title,
            material_text: payload.material_text,
            outline_text: payload.outline_text,
            style: payload.style,
            template_id: payload.template_id,
            target_pages: payload.target_pages,
            creation_type: "idea",
        }),
    });
    if (!createRes.ok)
        return readError(createRes, `创建项目失败: HTTP ${createRes.status}`);
    const project = (await createRes.json());
    const projectId = project.project_id;
    const outlineRes = await fetch(`${API_BASE}/projects/${projectId}/generate/outline`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
            outline: payload.outline || [],
            outline_markdown: payload.outline_markdown || payload.outline_text || "",
        }),
    });
    if (!outlineRes.ok)
        return readError(outlineRes, `生成大纲失败: HTTP ${outlineRes.status}`);
    const descRes = await fetch(`${API_BASE}/projects/${projectId}/generate/descriptions`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({}),
    });
    if (!descRes.ok)
        return readError(descRes, `描述任务创建失败: HTTP ${descRes.status}`);
    const descTask = (await descRes.json());
    await pollTask(projectId, descTask.task_id);
    const pptRes = await fetch(`${API_BASE}/projects/${projectId}/generate/ppt`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({}),
    });
    if (!pptRes.ok)
        return readError(pptRes, `导出任务创建失败: HTTP ${pptRes.status}`);
    const pptTask = (await pptRes.json());
    await pollTask(projectId, pptTask.task_id);
    return { job_id: projectId };
}
export async function getJob(jobId) {
    const res = await fetch(`${API_BASE}/projects/${jobId}`);
    if (!res.ok)
        return readError(res, `查询任务失败: HTTP ${res.status}`);
    const project = (await res.json());
    return mapProjectToJob(project);
}
export async function rewriteJob(jobId, action) {
    const res = await fetch(`${API_BASE}/projects/${jobId}/rewrite`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action }),
    });
    if (!res.ok)
        return readError(res, `重写失败: HTTP ${res.status}`);
    return (await res.json());
}
export async function getHistory() {
    const res = await fetch(`${API_BASE}/projects`);
    if (!res.ok)
        return readError(res, `获取历史失败: HTTP ${res.status}`);
    const list = (await res.json());
    return list.map((item) => ({
        job_id: item.project_id,
        title: item.title,
        style: item.style,
        template_id: item.template_id,
        status: item.status,
        created_at: item.created_at,
    }));
}
export async function parseUpload(file) {
    const fd = new FormData();
    fd.append("file", file);
    const res = await fetch(`${API_BASE}/parse-upload`, {
        method: "POST",
        body: fd,
    });
    if (!res.ok)
        return readError(res, `文件解析失败: HTTP ${res.status}`);
    const payload = (await res.json());
    return payload.extracted_text;
}
