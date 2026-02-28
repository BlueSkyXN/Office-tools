"""GUI → Core 适配层 — 将 UI 参数映射为 CORE-INTERFACE 请求模型"""

from __future__ import annotations

import uuid
from core.api import TaskRequest, TaskResponse, RuntimeOptions, run_task


def build_request(
    task_type: str,
    input_path: str,
    output_dir: str | None = None,
    options: dict | None = None,
    workers: int = 1,
) -> TaskRequest:
    """从 GUI 参数构建 TaskRequest"""
    return TaskRequest(
        task_type=task_type,
        input_path=input_path,
        output_dir=output_dir or None,
        options=options or {},
        runtime=RuntimeOptions(workers=workers),
        task_id=uuid.uuid4().hex[:12],
    )


def execute(request: TaskRequest) -> TaskResponse:
    """调用 core 统一入口"""
    return run_task(request)
