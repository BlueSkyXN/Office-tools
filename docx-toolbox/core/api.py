"""请求/响应模型与统一调度入口 — 遵循 CORE-INTERFACE.md"""

from __future__ import annotations

import threading
import uuid
from dataclasses import dataclass, field
from typing import Any

from core.errors import TaskError, ErrorCode, CancelledError


# ---------------------------------------------------------------------------
# 请求 / 响应模型
# ---------------------------------------------------------------------------

@dataclass
class RuntimeOptions:
    workers: int = 1
    dry_run: bool = False
    cancel_event: threading.Event | None = field(default=None, repr=False, compare=False)


@dataclass
class TaskRequest:
    task_type: str
    input_path: str
    output_dir: str | None = None
    options: dict[str, Any] = field(default_factory=dict)
    runtime: RuntimeOptions = field(default_factory=RuntimeOptions)
    task_id: str = field(default_factory=lambda: uuid.uuid4().hex[:12])


@dataclass
class TaskSummary:
    processed: int = 0
    failed: int = 0
    skipped: int = 0
    outputs: list[str] = field(default_factory=list)


@dataclass
class TaskResponse:
    ok: bool
    task_id: str
    status: str  # success | failed
    summary: TaskSummary | None = None
    error: dict[str, str] | None = None

    def to_dict(self) -> dict:
        d: dict[str, Any] = {
            "ok": self.ok,
            "task_id": self.task_id,
            "status": self.status,
        }
        if self.summary:
            d["summary"] = {
                "processed": self.summary.processed,
                "failed": self.summary.failed,
                "skipped": self.summary.skipped,
                "outputs": self.summary.outputs,
            }
        if self.error:
            d["error"] = self.error
        return d


# ---------------------------------------------------------------------------
# 调度入口
# ---------------------------------------------------------------------------

_ADAPTERS: dict[str, Any] = {}


def _ensure_adapters():
    if _ADAPTERS:
        return
    from core.adapters.excel_allinone import ExcelAllinoneAdapter
    from core.adapters.image_extract import ImageExtractAdapter
    from core.adapters.table_extract import TableExtractAdapter

    _ADAPTERS["excel_allinone"] = ExcelAllinoneAdapter()
    _ADAPTERS["image_extract"] = ImageExtractAdapter()
    _ADAPTERS["table_extract"] = TableExtractAdapter()


def run_task(
    request: TaskRequest,
    cancel_event: threading.Event | None = None,
) -> TaskResponse:
    """统一任务调度入口"""
    _ensure_adapters()

    if cancel_event is not None:
        request.runtime.cancel_event = cancel_event

    adapter = _ADAPTERS.get(request.task_type)
    if adapter is None:
        return TaskResponse(
            ok=False,
            task_id=request.task_id,
            status="failed",
            error=TaskError(
                ErrorCode.INVALID_INPUT,
                f"未知任务类型: {request.task_type}",
            ).to_dict(),
        )

    try:
        if request.runtime.cancel_event and request.runtime.cancel_event.is_set():
            raise CancelledError()
        summary = adapter.execute(request)
        return TaskResponse(
            ok=True,
            task_id=request.task_id,
            status="success",
            summary=summary,
        )
    except TaskError as e:
        return TaskResponse(
            ok=False,
            task_id=request.task_id,
            status="failed",
            error=e.to_dict(),
        )
    except Exception as e:
        return TaskResponse(
            ok=False,
            task_id=request.task_id,
            status="failed",
            error={
                "code": ErrorCode.INTERNAL.value,
                "message": str(e),
                "detail": type(e).__name__,
            },
        )
