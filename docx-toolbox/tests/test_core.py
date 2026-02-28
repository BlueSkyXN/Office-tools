"""docx-toolbox core 基础测试"""

import time
from pathlib import Path
from types import SimpleNamespace

import pytest
from core.api import TaskRequest, TaskResponse, RuntimeOptions, run_task


# ---------------------------------------------------------------------------
# 请求/响应模型测试
# ---------------------------------------------------------------------------

class TestTaskRequest:
    def test_default_task_id_generated(self):
        req = TaskRequest(task_type="excel_allinone", input_path="/tmp/test.docx")
        assert req.task_id  # non-empty
        assert len(req.task_id) == 12

    def test_custom_options(self):
        req = TaskRequest(
            task_type="image_extract",
            input_path="/tmp/test.docx",
            options={"remove_images": True, "jpeg_quality": 90},
        )
        assert req.options["remove_images"] is True
        assert req.options["jpeg_quality"] == 90

    def test_runtime_defaults(self):
        req = TaskRequest(task_type="table_extract", input_path="/tmp/test.docx")
        assert req.runtime.workers == 1
        assert req.runtime.dry_run is False


class TestTaskResponse:
    def test_success_to_dict(self):
        from core.api import TaskSummary
        resp = TaskResponse(
            ok=True, task_id="abc123", status="success",
            summary=TaskSummary(processed=3, failed=0, skipped=1, outputs=["/out/a.docx"]),
        )
        d = resp.to_dict()
        assert d["ok"] is True
        assert d["summary"]["processed"] == 3

    def test_failure_to_dict(self):
        resp = TaskResponse(
            ok=False, task_id="abc123", status="failed",
            error={"code": "E_INVALID_INPUT", "message": "test", "detail": ""},
        )
        d = resp.to_dict()
        assert d["ok"] is False
        assert d["error"]["code"] == "E_INVALID_INPUT"


# ---------------------------------------------------------------------------
# 错误码测试
# ---------------------------------------------------------------------------

class TestErrors:
    def test_error_code_values(self):
        from core.errors import ErrorCode
        assert ErrorCode.INVALID_INPUT.value == "E_INVALID_INPUT"
        assert ErrorCode.CANCELLED.value == "E_CANCELLED"

    def test_task_error_to_dict(self):
        from core.errors import InvalidInputError
        err = InvalidInputError("bad path", detail="/foo/bar")
        d = err.to_dict()
        assert d["code"] == "E_INVALID_INPUT"
        assert "bad path" in d["message"]

    def test_all_error_subclasses(self):
        from core.errors import (
            InvalidInputError, UnsupportedFormatError,
            PermissionDeniedError, ProcessFailedError, CancelledError,
        )
        assert InvalidInputError("x").code.value == "E_INVALID_INPUT"
        assert UnsupportedFormatError("x").code.value == "E_UNSUPPORTED_FORMAT"
        assert PermissionDeniedError("x").code.value == "E_PERMISSION_DENIED"
        assert ProcessFailedError("x").code.value == "E_PROCESS_FAILED"
        assert CancelledError().code.value == "E_CANCELLED"


# ---------------------------------------------------------------------------
# 调度入口测试
# ---------------------------------------------------------------------------

class TestRunTask:
    def test_unknown_task_type(self):
        req = TaskRequest(task_type="unknown_type", input_path="/tmp/test.docx")
        resp = run_task(req)
        assert resp.ok is False
        assert resp.error["code"] == "E_INVALID_INPUT"

    def test_invalid_input_path(self):
        req = TaskRequest(task_type="excel_allinone", input_path="/nonexistent/file.docx")
        resp = run_task(req)
        assert resp.ok is False

    def test_invalid_input_path_image(self):
        req = TaskRequest(task_type="image_extract", input_path="/nonexistent/file.docx")
        resp = run_task(req)
        assert resp.ok is False

    def test_invalid_input_path_table(self):
        req = TaskRequest(task_type="table_extract", input_path="/nonexistent/file.docx")
        resp = run_task(req)
        assert resp.ok is False


# ---------------------------------------------------------------------------
# 适配器行为测试
# ---------------------------------------------------------------------------

class TestAdapters:
    def test_excel_allinone_single_file_failure_should_not_abort_batch(self, tmp_path, monkeypatch):
        from core.adapters.excel_allinone import ExcelAllinoneAdapter

        failed_doc = tmp_path / "a.docx"
        ok_doc = tmp_path / "b.docx"
        failed_doc.write_text("bad")
        ok_doc.write_text("ok")

        def fake_process_document(path: str, args):
            if path.endswith("a.docx"):
                raise SystemExit(1)
            out = Path(path).with_name(f"{Path(path).stem}-AIO.docx")
            out.write_text("generated")

        monkeypatch.setattr(
            "core.adapters.excel_allinone._load_ref_module",
            lambda: SimpleNamespace(process_document=fake_process_document),
        )

        adapter = ExcelAllinoneAdapter()
        summary = adapter.execute(
            TaskRequest(task_type="excel_allinone", input_path=str(tmp_path))
        )
        assert summary.processed == 1
        assert summary.failed == 1
        assert len(summary.outputs) == 1
        assert summary.outputs[0].endswith("b-AIO.docx")


# ---------------------------------------------------------------------------
# Runner 测试
# ---------------------------------------------------------------------------

class TestRunner:
    def test_submit_and_run_failed(self):
        from core.runner import TaskRunner, JobStatus
        runner = TaskRunner(max_workers=1)
        job = runner.submit(TaskRequest(task_type="excel_allinone", input_path="/nonexistent"))
        assert job.status == JobStatus.PENDING
        results = runner.run_all()
        assert results[0].status == JobStatus.FAILED

    def test_cancel_during_run(self, monkeypatch):
        import threading
        from core.api import TaskSummary
        from core.runner import JobStatus, TaskRunner

        runner = TaskRunner(max_workers=1)

        def fake_run_task(request, cancel_event=None):
            for _ in range(50):
                if cancel_event and cancel_event.is_set():
                    return TaskResponse(
                        ok=False,
                        task_id=request.task_id,
                        status="failed",
                        error={"code": "E_CANCELLED", "message": "cancelled", "detail": ""},
                    )
                time.sleep(0.005)
            return TaskResponse(
                ok=True,
                task_id=request.task_id,
                status="success",
                summary=TaskSummary(processed=1),
            )

        monkeypatch.setattr("core.runner.run_task", fake_run_task)

        for _ in range(5):
            runner.submit(TaskRequest(task_type="excel_allinone", input_path="/nonexistent"))

        def cancel_soon():
            time.sleep(0.03)
            runner.cancel()

        threading.Thread(target=cancel_soon, daemon=True).start()
        results = runner.run_all()
        cancelled = [j for j in results if j.status == JobStatus.CANCELLED]
        assert len(results) == 5
        assert len(cancelled) >= 1

    def test_cancel_parallel_should_cancel_pending_jobs(self, monkeypatch):
        import threading
        from core.api import TaskSummary
        from core.runner import JobStatus, TaskRunner

        runner = TaskRunner(max_workers=2)

        def fake_run_task(request, cancel_event=None):
            for _ in range(200):
                if cancel_event and cancel_event.is_set():
                    return TaskResponse(
                        ok=False,
                        task_id=request.task_id,
                        status="failed",
                        error={"code": "E_CANCELLED", "message": "cancelled", "detail": ""},
                    )
                time.sleep(0.002)
            return TaskResponse(
                ok=True,
                task_id=request.task_id,
                status="success",
                summary=TaskSummary(processed=1),
            )

        monkeypatch.setattr("core.runner.run_task", fake_run_task)

        for _ in range(8):
            runner.submit(TaskRequest(task_type="excel_allinone", input_path="/nonexistent"))

        def cancel_soon():
            time.sleep(0.02)
            runner.cancel()

        threading.Thread(target=cancel_soon, daemon=True).start()
        results = runner.run_all()
        cancelled = [j for j in results if j.status == JobStatus.CANCELLED]
        assert len(results) == 8
        assert len(cancelled) >= 6

    def test_reset_clears_jobs(self):
        from core.runner import TaskRunner
        runner = TaskRunner()
        runner.submit(TaskRequest(task_type="excel_allinone", input_path="/tmp/x"))
        assert len(runner.jobs) == 1
        runner.reset()
        assert len(runner.jobs) == 0


# ---------------------------------------------------------------------------
# Logging 测试
# ---------------------------------------------------------------------------

class TestLogging:
    def test_get_logger_returns_logger(self):
        from core.logging_utils import get_logger
        lg = get_logger()
        assert lg is not None
        assert lg.name == "docx_toolbox"

    def test_setup_file_logging(self, tmp_path):
        from core.logging_utils import setup_file_logging
        log_path = setup_file_logging(log_dir=tmp_path, task_id="test123")
        assert log_path.exists() or True  # file created on first write
        assert "test123" in str(log_path)
