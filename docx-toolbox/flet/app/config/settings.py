"""配置持久化 — JSON 文件读写"""

from __future__ import annotations

import json
from pathlib import Path
from dataclasses import dataclass, field, asdict

_CONFIG_DIR = Path.home() / ".docx-toolbox"
_CONFIG_FILE = _CONFIG_DIR / "settings.json"


@dataclass
class AppSettings:
    default_output_dir: str = ""
    worker_count: int = 2
    log_max_lines: int = 500

    def save(self) -> None:
        _CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        _CONFIG_FILE.write_text(json.dumps(asdict(self), ensure_ascii=False, indent=2), encoding="utf-8")

    @classmethod
    def load(cls) -> "AppSettings":
        if _CONFIG_FILE.exists():
            try:
                data = json.loads(_CONFIG_FILE.read_text(encoding="utf-8"))
                return cls(**{k: v for k, v in data.items() if k in cls.__dataclass_fields__})
            except Exception:
                pass
        return cls()
