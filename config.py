import os
import sys
from pathlib import Path

from dotenv import load_dotenv


def get_base_dir() -> Path:

    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent

    return Path(__file__).resolve().parent


BASE_DIR = get_base_dir()
ENV_PATH = BASE_DIR / ".env"

load_dotenv(ENV_PATH, override=True)

APP_NAME = os.getenv("APP_NAME", "Procurement_app")
APP_ENV = os.getenv("APP_ENV", "development")

SQLALCHEMY_DATABASE_URI = os.getenv("SQLALCHEMY_DATABASE_URI")

if not SQLALCHEMY_DATABASE_URI:
    raise RuntimeError(
        f"SQLALCHEMY_DATABASE_URI не найден. Проверяемый файл: {ENV_PATH}"
    )