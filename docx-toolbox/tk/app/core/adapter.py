"""GUI 参数 → core.TaskRequest 适配层"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

# 确保能导入顶层 core 包
_DOCX_ROOT = str(Path(__file__).resolve().parent.parent.parent.parent)
if _DOCX_ROOT not in sys.path:
    sys.path.insert(0, _DOCX_ROOT)

from core.api import TaskRequest, RuntimeOptions  # noqa: E402


def build_request(
    task_type: str,
    input_path: str,
    output_dir: str | None = None,
    options: dict[str, Any] | None = None,
    workers: int = 1,
) -> TaskRequest:
    """将 GUI 收集的参数转化为 core.TaskRequest"""
    return TaskRequest(
        task_type=task_type,
        input_path=input_path,
        output_dir=output_dir or None,
        options=options or {},
        runtime=RuntimeOptions(workers=workers),
    )
