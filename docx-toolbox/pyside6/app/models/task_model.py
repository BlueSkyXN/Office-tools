"""Task state model for UI binding"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any


class TaskStatus(str, Enum):
    IDLE = "idle"
    PENDING = "pending"
    RUNNING = "running"
    SUCCESS = "success"
    FAILED = "failed"
    CANCELLED = "cancelled"


@dataclass
class TaskState:
    """Holds the current task state that the UI observes."""
    task_type: str = ""
    input_path: str = ""
    output_dir: str = ""
    options: dict[str, Any] = field(default_factory=dict)
    status: TaskStatus = TaskStatus.IDLE
    progress: int = 0          # 0-100
    progress_text: str = ""
    error_message: str = ""

    def reset(self):
        self.status = TaskStatus.IDLE
        self.progress = 0
        self.progress_text = ""
        self.error_message = ""
