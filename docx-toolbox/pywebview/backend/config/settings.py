"""配置持久化 — JSON 文件读写"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

_DEFAULT_CONFIG = {
    "output_dir": "",
    "workers": 1,
    "theme": "light",
}

_CONFIG_FILE = Path(__file__).resolve().parent.parent.parent / "config.json"


def _load() -> dict[str, Any]:
    if _CONFIG_FILE.exists():
        try:
            return json.loads(_CONFIG_FILE.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            pass
    return dict(_DEFAULT_CONFIG)


def _save(data: dict[str, Any]):
    _CONFIG_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def get_all() -> dict[str, Any]:
    """读取全部配置"""
    return _load()


def get(key: str, default: Any = None) -> Any:
    """读取单项配置"""
    return _load().get(key, default)


def set_value(key: str, value: Any):
    """写入单项配置"""
    data = _load()
    data[key] = value
    _save(data)


def set_all(updates: dict[str, Any]):
    """批量更新配置"""
    data = _load()
    data.update(updates)
    _save(data)
