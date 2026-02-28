"""任务数据模型 — dataclass 定义"""

from __future__ import annotations

import uuid
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
class TaskItem:
    """GUI 任务队列中的单项"""
    task_type: str
    input_path: str
    output_dir: str = ""
    options: dict[str, Any] = field(default_factory=dict)
    workers: int = 1
    task_id: str = field(default_factory=lambda: uuid.uuid4().hex[:12])
    status: TaskStatus = TaskStatus.PENDING
    progress: int = 0
    message: str = ""
