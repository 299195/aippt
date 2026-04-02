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
} from "./api";

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

const formatOutlineText = (outline) => outline.map((item, idx) => `${idx + 1}. ${item}`).join("\n");
const parseOutlineText = (text) =>
  text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => line.replace(/^(\d+[\.、\)）]?|[-*])\s*/, "").trim())
    .filter(Boolean);

function App() {
  const [title, setTitle] = useState("");
  const [materialText, setMaterialText] = useState("");
  const [outlineText, setOutlineText] = useState("");
  const [style, setStyle] = useState("management");
  const [templateId, setTemplateId] = useState("no_template");

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

  const errorText = (e, fallback) => {
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

  const formatHistoryTime = (value) => {
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
      });
      setOutlineText((res.outline_markdown || "").trim() || formatOutlineText(res.outline));
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
        outline_text: outlineText,
        outline_markdown: outlineText,
        outline: normalizedOutline,
        style,
        template_id: templateId,
      });
      setJobId(res.job_id);

      const detail = await getJob(res.job_id);
      setJob(detail);
      setOutlineText((detail.outline_text || "").trim() || formatOutlineText(detail.outline));

      refreshHistory();
      refreshModelConfig();
    } catch (e) {
      setError(errorText(e, "生成失败"));
    } finally {
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
      setOutlineText((detail.outline_text || "").trim() || formatOutlineText(detail.outline));
      setTemplateId(templates.some((tpl) => tpl.id === detail.template_id) ? detail.template_id : "no_template");
      setStyle(detail.style);
    } catch (e) {
      setError(errorText(e, "任务读取失败"));
    } finally {
      setLoading(false);
    }
  }

  async function handleUpload(file) {
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
          <span className="hero-badge">AI PPT</span>
          输入 PPT 主题
          <br />
          生成高质量PPT
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
                <input type="file" accept=".md,.docx" onChange={(e) => handleUpload(e.target.files?.[0] ?? null)} />+
              </label>

              <div className="style-switch" role="group" aria-label="风格选择">
                {[
                  ["management", "管理版"],
                  ["technical", "技术版"],
                ].map(([styleKey, label]) => (
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
          <button className={`switch-btn ${bottomTab === "templates" ? "active" : ""}`} type="button" onClick={() => setBottomTab("templates")}>
            PPT模板
          </button>
          <button className={`switch-btn ${bottomTab === "history" ? "active" : ""}`} type="button" onClick={() => setBottomTab("history")}>
            历史记录
          </button>
        </div>

        {bottomTab === "templates" ? (
          <section className="panel-block templates-block">
            <div className="template-list side-template-list">
              {templates.map((tpl) => (
                <button key={tpl.id} className={`template-card ${templateId === tpl.id ? "active" : ""}`} onClick={() => setTemplateId(tpl.id)} type="button">
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
                    <span className="history-meta">{`${STYLE_LABEL[item.style]}  ${formatHistoryTime(item.created_at)}`}</span>
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

