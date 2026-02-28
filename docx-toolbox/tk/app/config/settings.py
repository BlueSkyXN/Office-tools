"""JSON 配置读写 — 记住上次路径与偏好"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

_DEFAULT_PATH = Path.home() / ".docx-toolbox" / "tk_settings.json"

_DEFAULTS: dict[str, Any] = {
    "last_input_path": "",
    "last_output_dir": "",
    "last_task_type": "excel_allinone",
    "window_geometry": "",
    "options": {},
}


class Settings:
    """简单 JSON 配置管理器"""

    def __init__(self, path: Path | None = None):
        self._path = path or _DEFAULT_PATH
        self._data: dict[str, Any] = dict(_DEFAULTS)
        self._load()

    def _load(self) -> None:
        if self._path.exists():
            try:
                with open(self._path, "r", encoding="utf-8") as f:
                    stored = json.load(f)
                self._data.update(stored)
            except (json.JSONDecodeError, OSError):
                pass

    def save(self) -> None:
        self._path.parent.mkdir(parents=True, exist_ok=True)
        with open(self._path, "w", encoding="utf-8") as f:
            json.dump(self._data, f, ensure_ascii=False, indent=2)

    def get(self, key: str, default: Any = None) -> Any:
        return self._data.get(key, default)

    def set(self, key: str, value: Any) -> None:
        self._data[key] = value

    def __getitem__(self, key: str) -> Any:
        return self._data[key]

    def __setitem__(self, key: str, value: Any) -> None:
        self._data[key] = value
