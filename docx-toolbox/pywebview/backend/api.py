"""pywebview JS Bridge API — 遵循 DESIGN.md §6"""

from __future__ import annotations

import os
import subprocess
import sys
import traceback
from pathlib import Path

import webview

from backend.services.task_service import TaskService


def _ok(data=None) -> dict:
    return {"ok": True, "data": data}


def _err(code: str, message: str) -> dict:
    return {"ok": False, "error": {"code": code, "message": message}}


class ApiBridge:
    """暴露给 window.pywebview.api 的桥接类"""

    def __init__(self):
        self._task_service = TaskService()

    # ------------------------------------------------------------------
    # 文件 / 目录选择
    # ------------------------------------------------------------------

    def select_input_path(self) -> dict:
        """打开文件选择对话框，返回选中路径"""
        try:
            window = webview.windows[0]
            result = window.create_file_dialog(
                webview.OPEN_DIALOG,
                file_types=(
                    "Word 文档 (*.docx)|*.docx",
                    "所有文件 (*.*)|*.*",
                ),
            )
            if result and len(result) > 0:
                return _ok(result[0])
            return _ok(None)
        except Exception as e:
            return _err("E_INTERNAL", str(e))

    def select_folder(self) -> dict:
        """打开文件夹选择对话框"""
        try:
            window = webview.windows[0]
            result = window.create_file_dialog(webview.FOLDER_DIALOG)
            if result and len(result) > 0:
                return _ok(result[0])
            return _ok(None)
        except Exception as e:
            return _err("E_INTERNAL", str(e))

    # ------------------------------------------------------------------
    # 任务管理
    # ------------------------------------------------------------------

    def start_task(self, payload: dict) -> dict:
        """启动后台任务，payload 遵循 CORE-INTERFACE §2"""
        try:
            task_type = payload.get("task_type")
            input_path = payload.get("input_path")
            if not task_type or not input_path:
                return _err("E_INVALID_INPUT", "缺少 task_type 或 input_path")

            if not Path(input_path).exists():
                return _err("E_INVALID_INPUT", f"文件不存在: {input_path}")

            task_info = self._task_service.start(payload)
            return _ok(task_info)
        except Exception as e:
            return _err("E_INTERNAL", str(e))

    def cancel_task(self, task_id: str) -> dict:
        """取消指定任务"""
        try:
            self._task_service.cancel(task_id)
            return _ok({"task_id": task_id, "status": "cancelled"})
        except KeyError:
            return _err("E_INVALID_INPUT", f"任务不存在: {task_id}")
        except Exception as e:
            return _err("E_INTERNAL", str(e))

    def get_task_status(self, task_id: str) -> dict:
        """查询任务状态"""
        try:
            info = self._task_service.get_status(task_id)
            return _ok(info)
        except KeyError:
            return _err("E_INVALID_INPUT", f"任务不存在: {task_id}")
        except Exception as e:
            return _err("E_INTERNAL", str(e))

    def list_tasks(self) -> dict:
        """列出所有任务"""
        try:
            tasks = self._task_service.list_all()
            return _ok(tasks)
        except Exception as e:
            return _err("E_INTERNAL", str(e))

    # ------------------------------------------------------------------
    # 辅助功能
    # ------------------------------------------------------------------

    def open_output_folder(self, path: str) -> dict:
        """在系统文件管理器中打开输出目录"""
        try:
            target = Path(path)
            if not target.exists():
                return _err("E_INVALID_INPUT", f"路径不存在: {path}")

            folder = str(target if target.is_dir() else target.parent)
            if sys.platform == "darwin":
                subprocess.Popen(["open", folder])
            elif sys.platform == "win32":
                os.startfile(folder)
            else:
                subprocess.Popen(["xdg-open", folder])
            return _ok(None)
        except Exception as e:
            return _err("E_INTERNAL", str(e))

    def export_logs(self, task_id: str) -> dict:
        """导出指定任务的日志文件路径"""
        try:
            log_path = self._task_service.get_log_path(task_id)
            if log_path and Path(log_path).exists():
                return _ok({"path": str(log_path)})
            return _err("E_INVALID_INPUT", f"日志文件不存在: {task_id}")
        except Exception as e:
            return _err("E_INTERNAL", str(e))
