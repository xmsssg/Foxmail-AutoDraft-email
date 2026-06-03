from __future__ import annotations

import json
import os
import shutil
import sys
from pathlib import Path
from typing import Any


def _app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[2]


def _install_root(app_dir: Path) -> Path:
    env_root = os.environ.get("AUTOEMAIL_HOME")
    if env_root:
        return Path(env_root).resolve()
    if getattr(sys, "frozen", False):
        return app_dir.parent if app_dir.name.lower() == "app" else app_dir
    return app_dir


def _default_config_dir(app_dir: Path) -> Path:
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", app_dir)) / "app" / "config"
    return app_dir / "app" / "config"


BASE_DIR = _app_dir()
APP_DIR = BASE_DIR
INSTALL_ROOT = _install_root(APP_DIR)
DATA_DIR = INSTALL_ROOT / "data"
LOG_DIR = INSTALL_ROOT / "logs"
OUTPUT_DIR = INSTALL_ROOT / "output"
CONFIG_DIR = INSTALL_ROOT / "config"
DEFAULT_CONFIG_DIR = _default_config_dir(APP_DIR)


def expand_path(value: str) -> str:
    replacements = {
        "{APP_DIR}": str(APP_DIR),
        "{INSTALL_ROOT}": str(INSTALL_ROOT),
        "{CONFIG_DIR}": str(CONFIG_DIR),
        "{DATA_DIR}": str(DATA_DIR),
        "{LOG_DIR}": str(LOG_DIR),
        "{OUTPUT_DIR}": str(OUTPUT_DIR),
    }
    result = value
    for token, replacement in replacements.items():
        result = result.replace(token, replacement)
    return os.path.expandvars(result)


def ensure_runtime_files() -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    (OUTPUT_DIR / "generated_eml").mkdir(parents=True, exist_ok=True)
    (OUTPUT_DIR / "generated_mapi_xml").mkdir(parents=True, exist_ok=True)
    for name in ("settings.json", "customers.json"):
        target = CONFIG_DIR / name
        source = DEFAULT_CONFIG_DIR / name
        if not target.exists() and source.exists():
            shutil.copy2(source, target)


def load_json(path: Path) -> Any:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def load_settings() -> dict[str, Any]:
    ensure_runtime_files()
    return load_json(CONFIG_DIR / "settings.json")


def load_customers() -> list[dict[str, Any]]:
    ensure_runtime_files()
    return load_json(CONFIG_DIR / "customers.json")
