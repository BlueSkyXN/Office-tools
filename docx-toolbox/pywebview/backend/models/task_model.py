"""任务数据模型"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any


class TaskStatus(str, Enum):
    PENDING = "pending"
    RUNNING = "running"
    SUCCESS = "success"
    FAILED = "failed"
    CANCELLED = "cancelled"


@dataclass
class TaskRecord:
    """本地任务记录，用于跟踪后台执行状态"""

    task_id: str
    task_type: str
    input_path: str
    output_dir: str | None = None
    options: dict[str, Any] = field(default_factory=dict)
    status: TaskStatus = TaskStatus.PENDING
    summary: dict | None = None
    error: dict | None = None
    log_path: str | None = None
    created_at: str = ""

    def to_dict(self) -> dict:
        return {
            "task_id": self.task_id,
            "task_type": self.task_type,
            "input_path": self.input_path,
            "output_dir": self.output_dir,
            "status": self.status.value,
            "summary": self.summary,
            "error": self.error,
            "log_path": self.log_path,
            "created_at": self.created_at,
        }
