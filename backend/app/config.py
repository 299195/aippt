import json
import os
from pathlib import Path

from pydantic import BaseModel


def _load_model_config(config_path: Path) -> dict:
    if not config_path.exists():
        return {}
    try:
        return json.loads(config_path.read_text(encoding="utf-8-sig"))
    except Exception:
        return {}


def _env_or_default(name: str, default):
    value = os.getenv(name)
    if value is None:
        return default
    if isinstance(value, str) and value.strip() == "":
        return default
    return value


def _env_bool(name: str, default: bool) -> bool:
    value = os.getenv(name)
    if value is None or value.strip() == "":
        return default
    return value.strip().lower() in ("1", "true", "yes", "on")


def _env_int(name: str, default: int) -> int:
    value = os.getenv(name)
    if value is None or value.strip() == "":
        return default
    try:
        return int(value)
    except ValueError:
        return default


MODEL_CONFIG_PATH = Path(__file__).resolve().parents[1] / "model_provider.json"
MODEL_CONFIG = _load_model_config(MODEL_CONFIG_PATH)


class Settings(BaseModel):
    project_name: str = "AI PPT Assistant"
    data_dir: Path = Path("data")
    export_dir: Path = Path("exports")
    database_path: Path = Path("data/history.db")

    model_provider: str = _env_or_default("MODEL_PROVIDER", MODEL_CONFIG.get("provider", "doubao"))
    use_mock_llm: bool = _env_bool("USE_MOCK_LLM", bool(MODEL_CONFIG.get("use_mock", True)))

    model_base_url: str = _env_or_default("MODEL_BASE_URL", MODEL_CONFIG.get("base_url", ""))
    model_api_key: str = _env_or_default("MODEL_API_KEY", MODEL_CONFIG.get("api_key", ""))
    model_endpoint_id: str = _env_or_default(
        "MODEL_ENDPOINT_ID",
        MODEL_CONFIG.get("endpoint_id", MODEL_CONFIG.get("model", "")),
    )
    model_name: str = _env_or_default(
        "MODEL_NAME",
        _env_or_default("MODEL_ENDPOINT_ID", MODEL_CONFIG.get("model", MODEL_CONFIG.get("endpoint_id", ""))),
    )
    model_chat_path: str = _env_or_default("MODEL_CHAT_PATH", MODEL_CONFIG.get("chat_path", "/v1/chat/completions"))
    request_timeout_sec: int = _env_int("MODEL_TIMEOUT", int(MODEL_CONFIG.get("timeout", 60)))

    enable_image_generation: bool = _env_bool(
        "ENABLE_IMAGE_GENERATION",
        bool(MODEL_CONFIG.get("enable_image_generation", True)),
    )
    use_mock_image: bool = _env_bool("USE_MOCK_IMAGE", bool(MODEL_CONFIG.get("use_mock_image", True)))
    image_fallback_mock: bool = _env_bool(
        "IMAGE_FALLBACK_MOCK",
        bool(MODEL_CONFIG.get("image_fallback_mock", True)),
    )

    image_base_url: str = _env_or_default("IMAGE_BASE_URL", MODEL_CONFIG.get("image_base_url", MODEL_CONFIG.get("base_url", "")))
    image_api_key: str = _env_or_default("IMAGE_API_KEY", MODEL_CONFIG.get("image_api_key", MODEL_CONFIG.get("api_key", "")))
    image_model: str = _env_or_default("IMAGE_MODEL", MODEL_CONFIG.get("image_model", ""))
    image_gen_path: str = _env_or_default("IMAGE_GEN_PATH", MODEL_CONFIG.get("image_gen_path", "/v1/images/generations"))
    image_size: str = _env_or_default("IMAGE_SIZE", MODEL_CONFIG.get("image_size", "1536x1024"))
    image_timeout_sec: int = _env_int("IMAGE_TIMEOUT", int(MODEL_CONFIG.get("image_timeout", 90)))
    generated_image_dir: Path = Path("exports/generated_images")


settings = Settings()
settings.data_dir.mkdir(parents=True, exist_ok=True)
settings.export_dir.mkdir(parents=True, exist_ok=True)
settings.generated_image_dir.mkdir(parents=True, exist_ok=True)


