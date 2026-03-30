import { jsx as _jsx, jsxs as _jsxs, Fragment as _Fragment } from "react/jsx-runtime";
import { useEffect, useMemo, useState } from "react";
import { createJob, fileUrl, getHistory, getJob, getModelConfig, getTemplates, parseUpload, previewOutline, rewriteJob, } from "./api";
const QUICK_TOPICS = [
    "Q2 经营复盘与增长策略",
    "互联网竞品分析与应对策略",
    "制造业产线效率提升方案",
    "AI 应用落地阶段汇报",
    "融资路演核心故事线",
    "年度项目里程碑回顾",
];
const STYLE_LABEL = {
    management: "管理版",
    technical: "技术版",
};
const STATUS_LABEL = {
    pending: "排队中",
    queued: "排队中",
    running: "生成中",
    processing: "处理中",
    done: "已完成",
    succeeded: "已完成",
    success: "已完成",
    draft: "草稿",
    outline_generated: "大纲已生成",
    descriptions_generated: "文案已生成",
    completed: "已完成",
    failed: "失败",
    error: "失败",
};
const formatOutlineText = (outline) => outline.map((item, idx) => `${idx + 1}. ${item}`).join("\n");
const parseOutlineText = (text) => text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => line.replace(/^(\d+[\.|、\)|）]|[-*])\s*/, "").trim())
    .filter(Boolean);
function App() {
    const [title, setTitle] = useState("");
    const [materialText, setMaterialText] = useState("");
    const [outlineText, setOutlineText] = useState("");
    const [style, setStyle] = useState("management");
    const [templateId, setTemplateId] = useState("no_template");
    const [pages, setPages] = useState(8);
    const [loading, setLoading] = useState(false);
    const [outlineLoading, setOutlineLoading] = useState(false);
    const [uploading, setUploading] = useState(false);
    const [jobId, setJobId] = useState("");
    const [job, setJob] = useState(null);
    const [history, setHistory] = useState([]);
    const [templates, setTemplates] = useState([]);
    const [bottomTab, setBottomTab] = useState("templates");
    const [error, setError] = useState("");
    const [modelConfig, setModelConfig] = useState(null);
    const [rewriteOpen, setRewriteOpen] = useState(false);
    const errorText = (e, fallback) => {
        if (e instanceof Error && e.message)
            return e.message;
        return fallback;
    };
    const editorValue = outlineText || title;
    const modelText = useMemo(() => {
        if (!modelConfig)
            return "模型配置读取中";
        const provider = modelConfig.provider || "未知Provider";
        const model = modelConfig.model || "(未设置 model)";
        const mode = modelConfig.use_mock ? "Mock" : "Real";
        const cfg = modelConfig.configured ? "已配置" : "未配置";
        return `${provider}/${model} · ${mode} · ${cfg}`;
    }, [modelConfig]);
    const statusText = useMemo(() => {
        if (uploading)
            return "正在解析资料文件...";
        if (outlineLoading)
            return "正在生成大纲...";
        if (loading)
            return "正在生成PPT...";
        return `模型配置信息：${modelText}`;
    }, [uploading, outlineLoading, loading, modelText]);
    const downloadUrl = useMemo(() => fileUrl(job?.pptx_url ?? null), [job?.pptx_url]);
    const statusLabel = (status) => STATUS_LABEL[status.toLowerCase()] ?? "进行中";
    const formatHistoryTime = (value) => {
        const date = new Date(value);
        if (Number.isNaN(date.getTime()))
            return value;
        return date.toLocaleString("zh-CN", { hour12: false });
    };
    async function refreshHistory() {
        try {
            setHistory(await getHistory());
        }
        catch (e) {
            setError(errorText(e, "历史加载失败"));
        }
    }
    async function refreshModelConfig() {
        try {
            setModelConfig(await getModelConfig());
        }
        catch (e) {
            setError(errorText(e, "模型配置读取失败"));
        }
    }
    async function refreshTemplates() {
        try {
            const list = await getTemplates();
            setTemplates(list);
            setTemplateId((prev) => {
                if (list.length === 0)
                    return prev;
                return list.some((t) => t.id === prev) ? prev : list[0].id;
            });
        }
        catch (e) {
            setError(errorText(e, "Template list load failed"));
        }
    }
    useEffect(() => {
        refreshHistory();
        refreshModelConfig();
        refreshTemplates();
    }, []);
    async function handleGenerateOutline() {
        if (!title.trim()) {
            setError("请输入您的主题");
            return;
        }
        setError("");
        setOutlineLoading(true);
        try {
            const res = await previewOutline({
                title: title.trim(),
                material_text: materialText,
                outline_text: "",
                style,
                target_pages: pages,
            });
            setOutlineText(formatOutlineText(res.outline));
        }
        catch (e) {
            setError(errorText(e, "大纲生成失败"));
        }
        finally {
            setOutlineLoading(false);
        }
    }
    async function handleGenerate() {
        if (!title.trim()) {
            setError("请输入您的主题");
            return;
        }
        const normalizedOutline = parseOutlineText(outlineText);
        if (!normalizedOutline.length) {
            setError("请先生成并编辑大纲");
            return;
        }
        setError("");
        setLoading(true);
        try {
            const res = await createJob({
                title: title.trim(),
                material_text: materialText,
                outline_text: "",
                outline: normalizedOutline,
                style,
                template_id: templateId,
                target_pages: pages,
            });
            setJobId(res.job_id);
            const detail = await getJob(res.job_id);
            setJob(detail);
            setOutlineText(formatOutlineText(detail.outline));
            refreshHistory();
            refreshModelConfig();
        }
        catch (e) {
            setError(errorText(e, "生成失败"));
        }
        finally {
            setLoading(false);
        }
    }
    async function handleRewrite(action) {
        if (!jobId)
            return;
        setError("");
        setLoading(true);
        try {
            await rewriteJob(jobId, action);
            const detail = await getJob(jobId);
            setJob(detail);
            setOutlineText(formatOutlineText(detail.outline));
            refreshHistory();
            refreshModelConfig();
        }
        catch (e) {
            setError(errorText(e, "复写失败"));
        }
        finally {
            setLoading(false);
        }
    }
    async function loadJob(id) {
        setLoading(true);
        setError("");
        try {
            setJobId(id);
            const detail = await getJob(id);
            setJob(detail);
            setTitle(detail.title);
            setOutlineText(formatOutlineText(detail.outline));
            setTemplateId(templates.some((tpl) => tpl.id === detail.template_id) ? detail.template_id : "no_template");
            setStyle(detail.style);
            if (detail.slides.length) {
                setPages(detail.slides.length);
            }
        }
        catch (e) {
            setError(errorText(e, "任务读取失败"));
        }
        finally {
            setLoading(false);
        }
    }
    async function handleUpload(file) {
        if (!file)
            return;
        setUploading(true);
        setError("");
        try {
            const parsed = await parseUpload(file);
            setMaterialText(parsed);
        }
        catch (e) {
            setError(errorText(e, "上传解析失败，仅支持 .md/.docx"));
        }
        finally {
            setUploading(false);
        }
    }
    return (_jsxs("div", { className: "app-shell", children: [_jsx("span", { className: "shape shape-circle" }), _jsx("span", { className: "shape shape-cross" }), _jsx("span", { className: "shape shape-drop" }), _jsxs("header", { className: "hero-center", children: [_jsxs("h1", { children: [_jsx("span", { className: "hero-badge", children: "LivePPT" }), "\u8F93\u5165 PPT \u4E3B\u9898", _jsx("br", {}), "AI \u751F\u6210\u9AD8\u8D28\u91CF PPT"] }), _jsxs("div", { className: "hero-card", children: [_jsx("div", { className: "hero-meta", children: _jsx("span", { className: "status-text", children: statusText }) }), _jsxs("div", { className: "composer-card", children: [_jsx("textarea", { className: "composer-editor", value: editorValue, onChange: (e) => {
                                            const value = e.target.value;
                                            if (outlineText) {
                                                setOutlineText(value);
                                            }
                                            else {
                                                setTitle(value);
                                            }
                                        }, placeholder: "\u8F93\u5165\u60A8\u60F3\u521B\u4F5C\u7684 PPT \u4E3B\u9898" }), _jsxs("div", { className: "composer-toolbar", children: [_jsxs("label", { className: "upload-plus", title: "\u4E0A\u4F20\u8D44\u6599", children: [_jsx("input", { type: "file", accept: ".md,.docx", onChange: (e) => handleUpload(e.target.files?.[0] ?? null) }), "+"] }), _jsx("div", { className: "style-switch", role: "group", "aria-label": "\u98CE\u683C\u9009\u62E9", children: [
                                                    ["management", "管理版"],
                                                    ["technical", "技术版"],
                                                ].map(([styleKey, label]) => (_jsx("button", { type: "button", className: style === styleKey ? "switch-seg active" : "switch-seg", onClick: () => {
                                                        setStyle(styleKey);
                                                        setOutlineText("");
                                                    }, children: label }, styleKey))) }), _jsxs("div", { className: "page-stepper-wrap", children: [_jsx("span", { className: "page-stepper-label", children: "\u9875" }), _jsx("input", { className: "page-stepper", type: "number", min: 8, max: 12, step: 1, value: pages, onChange: (e) => {
                                                            const value = Number(e.target.value);
                                                            const clamped = Number.isFinite(value) ? Math.min(12, Math.max(8, value)) : 8;
                                                            setPages(clamped);
                                                            setOutlineText("");
                                                        } })] }), _jsx("div", { className: "toolbar-spacer" }), _jsx("button", { className: "btn btn-main", disabled: outlineLoading, onClick: handleGenerateOutline, children: outlineLoading ? "生成中..." : "生成大纲" })] })] }), materialText && _jsx("p", { className: "upload-tip", children: "\u8D44\u6599\u5DF2\u5BFC\u5165\uFF0C\u53EF\u76F4\u63A5\u751F\u6210\u5927\u7EB2" }), _jsxs("div", { className: "outline-actions", children: [_jsx("button", { className: "btn btn-primary", disabled: loading || !parseOutlineText(outlineText).length, onClick: handleGenerate, children: loading ? "处理中..." : "生成PPT" }), downloadUrl ? (_jsx("a", { className: "btn text-action", href: downloadUrl, target: "_blank", rel: "noreferrer", children: "\u4E0B\u8F7DPPTX" })) : (_jsx("span", { className: "btn text-action disabled", children: "\u4E0B\u8F7DPPTX" })), _jsxs("div", { className: `rewrite-menu ${loading || !jobId ? "disabled" : ""} ${rewriteOpen ? "open" : ""}`, onMouseEnter: () => {
                                            if (!loading && jobId)
                                                setRewriteOpen(true);
                                        }, onMouseLeave: () => setRewriteOpen(false), children: [_jsx("button", { className: "btn text-action", type: "button", disabled: loading || !jobId, onMouseEnter: () => {
                                                    if (!loading && jobId)
                                                        setRewriteOpen(true);
                                                }, children: "\u590D\u5199" }), _jsxs("div", { className: "rewrite-options", children: [_jsx("button", { className: "btn", type: "button", disabled: loading || !jobId, onClick: () => handleRewrite("concise"), children: "\u66F4\u7CBE\u7B80" }), _jsx("button", { className: "btn", type: "button", disabled: loading || !jobId, onClick: () => handleRewrite("management"), children: "\u66F4\u7BA1\u7406\u53E3\u5F84" }), _jsx("button", { className: "btn", type: "button", disabled: loading || !jobId, onClick: () => handleRewrite("technical"), children: "\u66F4\u6280\u672F\u7EC6\u8282" })] })] })] })] }), _jsx("div", { className: "quick-topics", children: QUICK_TOPICS.map((topic) => (_jsx("button", { type: "button", onClick: () => {
                                setTitle(topic);
                                setOutlineText("");
                            }, children: topic }, topic))) }), error && _jsx("p", { className: "error-banner", children: error })] }), _jsxs("main", { className: "workspace workspace-bottom", children: [_jsxs("div", { className: "bottom-switch", role: "tablist", "aria-label": "\u5E95\u90E8\u9762\u677F\u5207\u6362", children: [_jsx("button", { className: `switch-btn ${bottomTab === "templates" ? "active" : ""}`, type: "button", onClick: () => setBottomTab("templates"), children: "PPT\u6A21\u677F" }), _jsx("button", { className: `switch-btn ${bottomTab === "history" ? "active" : ""}`, type: "button", onClick: () => setBottomTab("history"), children: "\u5386\u53F2\u8BB0\u5F55" })] }), bottomTab === "templates" ? (_jsx("section", { className: "panel-block templates-block", children: _jsx("div", { className: "template-list side-template-list", children: templates.map((tpl) => (_jsxs("button", { className: `template-card ${templateId === tpl.id ? "active" : ""}`, onClick: () => setTemplateId(tpl.id), type: "button", children: [_jsx("div", { className: "template-preview", style: { background: tpl.preview_bg, color: tpl.preview_fg }, children: tpl.preview_image_url ? (_jsx("img", { className: "template-preview-image", src: fileUrl(tpl.preview_image_url), alt: `${tpl.name} preview`, loading: "lazy" })) : (_jsxs(_Fragment, { children: [_jsx("span", { className: "template-preview-head", style: { background: tpl.preview_fg } }), _jsx("span", { className: "template-preview-block", style: { borderColor: tpl.preview_accent } }), _jsx("span", { className: "template-preview-block", style: { borderColor: tpl.preview_accent } })] })) }), _jsx("strong", { children: tpl.name })] }, tpl.id))) }) })) : (_jsx("section", { className: "panel-block history-block", children: _jsx("ul", { className: "history-list", children: history.map((item) => (_jsx("li", { children: _jsxs("button", { className: "history-item", onClick: () => loadJob(item.job_id), children: [_jsx("span", { className: "history-title", children: item.title }), _jsx("span", { className: "history-meta", children: `${STYLE_LABEL[item.style]}\u3000\u3000${formatHistoryTime(item.created_at)}` })] }) }, item.job_id))) }) }))] })] }));
}
export default App;
