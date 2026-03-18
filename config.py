import os
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()


def _env_bool(name: str, default: bool = False) -> bool:
    raw = os.environ.get(name)
    if raw is None:
        return default
    return raw.strip().lower() in {"1", "true", "yes", "on"}


OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "")
OPENAI_BASE_URL = os.environ.get("OPENAI_BASE_URL", "https://openrouter.ai/api/v1")
OPENAI_MODEL = os.environ.get("OPENAI_MODEL", "anthropic/claude-sonnet-4-6")
DATA_DIR = Path(os.environ.get("DATA_DIR", "./data"))
RECORD_TEMPLATE_PATH = Path(os.environ.get("RECORD_TEMPLATE_PATH", "")).expanduser() if os.environ.get("RECORD_TEMPLATE_PATH") else None
PORT = int(os.environ.get("PORT", "5000"))
FLASK_DEBUG = _env_bool("FLASK_DEBUG", default=False)
MAX_UPLOAD_MB = 10

SECRET_KEY = os.environ.get("SECRET_KEY", "dev-change-me-in-production")
SQLALCHEMY_DATABASE_URI = f"sqlite:///{DATA_DIR.resolve() / 'app.db'}"
ADMIN_USERNAME = os.environ.get("ADMIN_USERNAME", "")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "")
