"""Core 适配层 — 将 GUI TaskItem 映射为 core.api.TaskRequest 并调用"""

from __future__ import annotations

import sys
from pathlib import Path

# 将 docx-toolbox 根目录加入 sys.path 以便导入共享 core
_TOOLBOX_ROOT = str(Path(__file__).resolve().parent.parent.parent.parent)
if _TOOLBOX_ROOT not in sys.path:
    sys.path.insert(0, _TOOLBOX_ROOT)

from core.api import TaskRequest, RuntimeOptions, TaskResponse, run_task  # noqa: E402
from pyqt6.app.models.task_model import TaskItem  # noqa: E402


def build_request(item: TaskItem) -> TaskRequest:
    """从 GUI TaskItem 构造 core TaskRequest"""
    return TaskRequest(
        task_type=item.task_type,
        input_path=item.input_path,
        output_dir=item.output_dir or None,
        options=dict(item.options),
        runtime=RuntimeOptions(workers=item.workers),
        task_id=item.task_id,
    )


def execute_task(item: TaskItem) -> TaskResponse:
    """执行任务并返回响应"""
    request = build_request(item)
    return run_task(request)
