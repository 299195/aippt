# AI生成PPT助手（主题+资料文件直生）

技术栈：
- 后端：Python + FastAPI + LangGraph + python-pptx
- 前端：React + Vite + TypeScript

## 目录

- `backend/`：API、LangGraph流程、可切换模型适配、PPT导出、历史记录
- `frontend/`：任务创建、结果预览、快速重写、历史查看

## 后端启动

```powershell
cd backend
conda activate aippt
pip install -r requirements.txt
uvicorn app.main:app --reload --host 127.0.0.1 --port 8001
```

## 与 SmartSchoolAI 新后端同款导出引擎

默认导出引擎已切换为 `Ai-To-PPTX`（与 `third_party/ai-to-pptx-backend` 同款）。

要求：
- 安装 `PHP >= 7.4`，并启用 `zip` 扩展
- 如 `php` 不在 PATH，可设置 `AIPPT_PHP_BIN` 指向 `php.exe`

可选环境变量：
- `AIPPT_EXPORT_ENGINE=ai_to_pptx|auto|legacy`（默认 `ai_to_pptx`）
- `AIPPT_PHP_BIN=C:\path\to\php.exe`

## 独立模型配置（可随时换模型）

编辑 `backend/model_provider.json`：

```json
{
  "provider": "doubao",
  "use_mock": false,
  "base_url": "https://ark.cn-beijing.volces.com/api",
  "api_key": "YOUR_API_KEY",
  "model": "YOUR_MODEL_NAME",
  "chat_path": "/v3/chat/completions",
  "timeout": 60
}
```

你也可以用环境变量覆盖：
- `MODEL_PROVIDER`
- `MODEL_BASE_URL`
- `MODEL_API_KEY`
- `MODEL_NAME`
- `MODEL_CHAT_PATH`
- `MODEL_TIMEOUT`
- `USE_MOCK_LLM`

## 前端启动

```powershell
cd frontend
npm install
npm run dev
```

## 当前Agent流程（不再拆六要素）

`parse -> outline -> fill -> quality -> repair(可选) -> style -> export -> persist`

输入：
- PPT主题
- 资料文件（md/docx）或资料文本
- 风格（管理/技术）
- 页数（8~12）

输出：
- 8~12页PPT初稿
- 快速重写（更精简/更管理口径/更技术细节）
- 导出pptx与历史记录
