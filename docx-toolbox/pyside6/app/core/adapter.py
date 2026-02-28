"""Thin adapter: GUI params -> core.TaskRequest -> core.run_task"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

# Ensure the top-level docx-toolbox dir is on sys.path so `import core` works
_DOCX_TOOLBOX_ROOT = str(Path(__file__).resolve().parent.parent.parent.parent)
if _DOCX_TOOLBOX_ROOT not in sys.path:
    sys.path.insert(0, _DOCX_TOOLBOX_ROOT)

from core.api import TaskRequest, TaskResponse, RuntimeOptions, run_task  # noqa: E402


def build_task_request(
    task_type: str,
    input_path: str,
    output_dir: str = "",
    options: dict[str, Any] | None = None,
    workers: int = 1,
) -> TaskRequest:
    """Map GUI form values to a core TaskRequest."""
    return TaskRequest(
        task_type=task_type,
        input_path=input_path,
        output_dir=output_dir or None,
        options=options or {},
        runtime=RuntimeOptions(workers=workers),
    )


def execute_task(request: TaskRequest) -> TaskResponse:
    """Convenience wrapper around core.run_task."""
    return run_task(request)
