from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv


BASE_DIR = Path(__file__).resolve().parent.parent
ENV_FILE = BASE_DIR / ".env"

if ENV_FILE.exists():
    load_dotenv(ENV_FILE)


def _optional_int(name: str) -> int | None:
    value = os.getenv(name, "").strip()
    if not value:
        return None
    try:
        return int(value)
    except ValueError:
        return None


@dataclass(frozen=True)
class Settings:
    bot_token: str
    allowed_chat_id: int | None
    allowed_user_id: int | None
    env: str
    log_level: str
    base_dir: Path
    logs_dir: Path
    data_dir: Path
    session_dir: Path
    report_base_url: str


def load_settings() -> Settings:
    return Settings(
        bot_token=os.getenv("TELEGRAM_BOT_TOKEN", "").strip(),
        allowed_chat_id=_optional_int("TELEGRAM_ALLOWED_CHAT_ID"),
        allowed_user_id=_optional_int("TELEGRAM_ALLOWED_USER_ID"),
        env=os.getenv("BOT_ENV", "dev").strip() or "dev",
        log_level=os.getenv("BOT_LOG_LEVEL", "INFO").strip() or "INFO",
        base_dir=BASE_DIR,
        logs_dir=BASE_DIR / "logs",
        data_dir=BASE_DIR / "data",
        session_dir=BASE_DIR / "session",
        report_base_url=os.getenv("REPORT_BASE_URL", "").strip().rstrip("/"),
    )



