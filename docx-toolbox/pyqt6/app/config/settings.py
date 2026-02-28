"""配置持久化 — 最近路径、默认参数"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

_CONFIG_DIR = Path.home() / ".docx-toolbox"
_CONFIG_FILE = _CONFIG_DIR / "pyqt6_settings.json"

_DEFAULTS: dict[str, Any] = {
    "last_input_path": "",
    "last_output_dir": "",
    "excel_options": {
        "word_table": True,
        "extract_excel": True,
        "image": True,
        "keep_attachment": False,
        "remove_watermark": False,
        "a3": False,
    },
    "image_options": {
        "remove_images": False,
        "optimize_images": True,
        "jpeg_quality": 85,
    },
    "table_options": {
        "include_marked": False,
    },
    "workers": 1,
}


def load_settings() -> dict[str, Any]:
    """加载配置，不存在则返回默认值"""
    if _CONFIG_FILE.exists():
        try:
            with open(_CONFIG_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
            merged = {**_DEFAULTS, **saved}
            return merged
        except Exception:
            pass
    return dict(_DEFAULTS)


def save_settings(data: dict[str, Any]) -> None:
    """持久化配置"""
    _CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    with open(_CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
