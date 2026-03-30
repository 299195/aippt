import { useEffect, useMemo, useState } from "react";

import {
  createJob,
  fileUrl,
  getHistory,
  getJob,
  getModelConfig,
  getTemplates,
  parseUpload,
  previewOutline,
  rewriteJob,
} from "./api";
import type {
  HistoryItem,
  JobDetail,
  ModelConfig,
  RewriteAction,
  StyleType,
  TemplateItem,
} from "./types";

const QUICK_TOPICS = [
  "Q2 经营复盘与增长策略",
  "互联网竞品分析与应对策略",
  "制造业产线效率提升方案",
  "AI 应用落地阶段汇报",
  "融资路演核心故事线",
  "年度项目里程碑回顾",
];

const STYLE_LABEL: Record<StyleType, string> = {
  management: "管理版",
  technical: "技术版",
};

const STATUS_LABEL: Record<string, string> = {
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

const formatOutlineText = (outline: string[]) => outline.map((item, idx) => `${idx + 1}. ${item}`).join("\n");

const parseOutlineText = (text: string): string[] =>
  text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => line.replace(/^(\d+[\.|、\)|）]|[-*])\s*/, "").trim())
    .filter(Boolean);

function App() {
  const [title, setTitle] = useState("");
  const [materialText, setMaterialText] = useState("");
  const [outlineText, setOutlineText] = useState("");
  const [style, setStyle] = useState<StyleType>("management");
  const [templateId, setTemplateId] = useState<string>("no_template");
  const [pages, setPages] = useState(8);

  const [loading, setLoading] = useState(false);
  const [outlineLoading, setOutlineLoading] = useState(false);
  const [uploading, setUploading] = useState(false);

  const [jobId, setJobId] = useState("");
  const [job, setJob] = useState<JobDetail | null>(null);
  const [history, setHistory] = useState<HistoryItem[]>([]);
  const [templates, setTemplates] = useState<TemplateItem[]>([]);
  const [bottomTab, setBottomTab] = useState<"templates" | "history">("templates");

  const [error, setError] = useState("");
  const [modelConfig, setModelConfig] = useState<ModelConfig | null>(null);
  const [rewriteOpen, setRewriteOpen] = useState(false);

  const errorText = (e: unknown, fallback: string): string => {
    if (e instanceof Error && e.message) return e.message;
    return fallback;
  };

  const editorValue = outlineText || title;

  const modelText = useMemo(() => {
    if (!modelConfig) return "模型配置读取中";
    const provider = modelConfig.provider || "未知Provider";
    const model = modelConfig.model || "(未设置 model)";
    const mode = modelConfig.use_mock ? "Mock" : "Real";
    const cfg = modelConfig.configured ? "已配置" : "未配置";
    return `${provider}/${model} · ${mode} · ${cfg}`;
  }, [modelConfig]);

  const statusText = useMemo(() => {
    if (uploading) return "正在解析资料文件...";
    if (outlineLoading) return "正在生成大纲...";
    if (loading) return "正在生成PPT...";
    return `模型配置信息：${modelText}`;
  }, [uploading, outlineLoading, loading, modelText]);

  const downloadUrl = useMemo(() => fileUrl(job?.pptx_url ?? null), [job?.pptx_url]);

  const statusLabel = (status: string): string => STATUS_LABEL[status.toLowerCase()] ?? "进行中";
  const formatHistoryTime = (value: string): string => {
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return value;
    return date.toLocaleString("zh-CN", { hour12: false });
  };
  async function refreshHistory() {
    try {
      setHistory(await getHistory());
    } catch (e) {
      setError(errorText(e, "历史加载失败"));
    }
  }

  async function refreshModelConfig() {
    try {
      setModelConfig(await getModelConfig());
    } catch (e) {
      setError(errorText(e, "模型配置读取失败"));
    }
  }

  async function refreshTemplates() {
    try {
      const list = await getTemplates();
      setTemplates(list);
      setTemplateId((prev) => {
        if (list.length === 0) return prev;
        return list.some((t) => t.id === prev) ? prev : list[0].id;
      });
    } catch (e) {
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
    } catch (e) {
      setError(errorText(e, "大纲生成失败"));
    } finally {
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
    } catch (e) {
      setError(errorText(e, "生成失败"));
    } finally {
      setLoading(false);
    }
  }

  async function handleRewrite(action: RewriteAction) {
    if (!jobId) return;
    setError("");
    setLoading(true);
    try {
      await rewriteJob(jobId, action);
      const detail = await getJob(jobId);
      setJob(detail);
      setOutlineText(formatOutlineText(detail.outline));
      refreshHistory();
      refreshModelConfig();
    } catch (e) {
      setError(errorText(e, "复写失败"));
    } finally {
      setLoading(false);
    }
  }

  async function loadJob(id: string) {
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
    } catch (e) {
      setError(errorText(e, "任务读取失败"));
    } finally {
      setLoading(false);
    }
  }

  async function handleUpload(file: File | null) {
    if (!file) return;
    setUploading(true);
    setError("");
    try {
      const parsed = await parseUpload(file);
      setMaterialText(parsed);
    } catch (e) {
      setError(errorText(e, "上传解析失败，仅支持 .md/.docx"));
    } finally {
      setUploading(false);
    }
  }

  return (
    <div className="app-shell">
      <span className="shape shape-circle" />
      <span className="shape shape-cross" />
      <span className="shape shape-drop" />

      <header className="hero-center">
        <h1>
          <span className="hero-badge">LivePPT</span>
          输入 PPT 主题
          <br />
          AI 生成高质量 PPT
        </h1>

        <div className="hero-card">
          <div className="hero-meta">
            <span className="status-text">{statusText}</span>
          </div>

          <div className="composer-card">
            <textarea
              className="composer-editor"
              value={editorValue}
              onChange={(e) => {
                const value = e.target.value;
                if (outlineText) {
                  setOutlineText(value);
                } else {
                  setTitle(value);
                }
              }}
              placeholder="输入您想创作的 PPT 主题"
            />

            <div className="composer-toolbar">
              <label className="upload-plus" title="上传资料">
                <input
                  type="file"
                  accept=".md,.docx"
                  onChange={(e) => handleUpload(e.target.files?.[0] ?? null)}
                />
                +
              </label>

              <div className="style-switch" role="group" aria-label="风格选择">
                {(
                  [
                    ["management", "管理版"],
                    ["technical", "技术版"],
                  ] as Array<[StyleType, string]>
                ).map(([styleKey, label]) => (
                  <button
                    key={styleKey}
                    type="button"
                    className={style === styleKey ? "switch-seg active" : "switch-seg"}
                    onClick={() => {
                      setStyle(styleKey);
                      setOutlineText("");
                    }}
                  >
                    {label}
                  </button>
                ))}
              </div>

              <div className="page-stepper-wrap">
                <span className="page-stepper-label">页</span>
                <input
                  className="page-stepper"
                  type="number"
                  min={8}
                  max={12}
                  step={1}
                  value={pages}
                  onChange={(e) => {
                    const value = Number(e.target.value);
                    const clamped = Number.isFinite(value) ? Math.min(12, Math.max(8, value)) : 8;
                    setPages(clamped);
                    setOutlineText("");
                  }}
                />
              </div>

              <div className="toolbar-spacer" />

              <button className="btn btn-main" disabled={outlineLoading} onClick={handleGenerateOutline}>
                {outlineLoading ? "生成中..." : "生成大纲"}
              </button>
            </div>
          </div>

          {materialText && <p className="upload-tip">资料已导入，可直接生成大纲</p>}

          <div className="outline-actions">
            <button className="btn btn-primary" disabled={loading || !parseOutlineText(outlineText).length} onClick={handleGenerate}>
              {loading ? "处理中..." : "生成PPT"}
            </button>

            {downloadUrl ? (
              <a className="btn text-action" href={downloadUrl} target="_blank" rel="noreferrer">
                下载PPTX
              </a>
            ) : (
              <span className="btn text-action disabled">下载PPTX</span>
            )}

            <div
              className={`rewrite-menu ${loading || !jobId ? "disabled" : ""} ${rewriteOpen ? "open" : ""}`}
              onMouseEnter={() => {
                if (!loading && jobId) setRewriteOpen(true);
              }}
              onMouseLeave={() => setRewriteOpen(false)}
            >
              <button
                className="btn text-action"
                type="button"
                disabled={loading || !jobId}
                onMouseEnter={() => {
                  if (!loading && jobId) setRewriteOpen(true);
                }}
              >
                复写
              </button>
              <div className="rewrite-options">
                <button className="btn" type="button" disabled={loading || !jobId} onClick={() => handleRewrite("concise")}>
                  更精简
                </button>
                <button className="btn" type="button" disabled={loading || !jobId} onClick={() => handleRewrite("management")}>
                  更管理口径
                </button>
                <button className="btn" type="button" disabled={loading || !jobId} onClick={() => handleRewrite("technical")}>
                  更技术细节
                </button>
              </div>
            </div>
          </div>
        </div>

        <div className="quick-topics">
          {QUICK_TOPICS.map((topic) => (
            <button
              key={topic}
              type="button"
              onClick={() => {
                setTitle(topic);
                setOutlineText("");
              }}
            >
              {topic}
            </button>
          ))}
        </div>

        {error && <p className="error-banner">{error}</p>}
      </header>

      <main className="workspace workspace-bottom">
        <div className="bottom-switch" role="tablist" aria-label="底部面板切换">
          <button
            className={`switch-btn ${bottomTab === "templates" ? "active" : ""}`}
            type="button"
            onClick={() => setBottomTab("templates")}
          >
            PPT模板
          </button>
          <button
            className={`switch-btn ${bottomTab === "history" ? "active" : ""}`}
            type="button"
            onClick={() => setBottomTab("history")}
          >
            历史记录
          </button>
        </div>

        {bottomTab === "templates" ? (
          <section className="panel-block templates-block">
            <div className="template-list side-template-list">
              {templates.map((tpl) => (
                <button
                  key={tpl.id}
                  className={`template-card ${templateId === tpl.id ? "active" : ""}`}
                  onClick={() => setTemplateId(tpl.id)}
                  type="button"
                >
                  <div className="template-preview" style={{ background: tpl.preview_bg, color: tpl.preview_fg }}>
                    {tpl.preview_image_url ? (
                      <img className="template-preview-image" src={fileUrl(tpl.preview_image_url)} alt={`${tpl.name} preview`} loading="lazy" />
                    ) : (
                      <>
                        <span className="template-preview-head" style={{ background: tpl.preview_fg }} />
                        <span className="template-preview-block" style={{ borderColor: tpl.preview_accent }} />
                        <span className="template-preview-block" style={{ borderColor: tpl.preview_accent }} />
                      </>
                    )}
                  </div>
                  <strong>{tpl.name}</strong>
                </button>
              ))}
            </div>
          </section>
        ) : (
          <section className="panel-block history-block">
            <ul className="history-list">
              {history.map((item) => (
                <li key={item.job_id}>
                  <button className="history-item" onClick={() => loadJob(item.job_id)}>
                    <span className="history-title">{item.title}</span>
                    <span className="history-meta">{`${STYLE_LABEL[item.style]}\u3000\u3000${formatHistoryTime(item.created_at)}`}</span>
                  </button>
                </li>
              ))}
            </ul>
          </section>
        )}
      </main>
    </div>
  );
}

export default App;


















