"""集中式应用状态"""

from __future__ import annotations

import threading
from dataclasses import dataclass, field
from typing import Callable


@dataclass
class TaskState:
    running: bool = False
    cancelled: bool = False
    progress: float = 0.0
    input_path: str = ""
    output_dir: str = ""


class AppState:
    """全局状态容器，线程安全"""

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self.current_page: int = 0
        self.task_states: dict[str, TaskState] = {
            "excel": TaskState(),
            "image": TaskState(),
            "table": TaskState(),
            "batch": TaskState(),
        }
        self.log_lines: list[str] = []
        self._log_listeners: list[Callable[[str], None]] = []

    def add_log(self, line: str) -> None:
        with self._lock:
            self.log_lines.append(line)
            if len(self.log_lines) > 500:
                self.log_lines = self.log_lines[-500:]
            listeners = list(self._log_listeners)
        for cb in listeners:
            try:
                cb(line)
            except Exception:
                pass

    def on_log(self, callback: Callable[[str], None]) -> None:
        with self._lock:
            self._log_listeners.append(callback)

    def get_task(self, key: str) -> TaskState:
        return self.task_states[key]
