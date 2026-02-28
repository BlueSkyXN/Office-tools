"""Tab4: 批处理任务"""

from __future__ import annotations

import tkinter as tklib
from tkinter import ttk
from pathlib import Path
from typing import TYPE_CHECKING

from .widgets.file_picker import FilePicker

if TYPE_CHECKING:
    from ..runner.worker import TaskWorker
    from .widgets.log_text import LogText

TASK_TYPES = {
    "Excel 嵌入处理": "excel_allinone",
    "图片分离": "image_extract",
    "表格提取": "table_extract",
}


class BatchTab(ttk.Frame):
    """批处理标签页 — 选择文件夹，扫描 .docx 并逐文件处理"""

    def __init__(self, master: tklib.Misc, log: LogText, worker: TaskWorker, **kwargs):
        super().__init__(master, **kwargs)
        self._log = log
        self._worker = worker
        self._jobs: list[dict] = []
        self._current_idx = 0
        self._build_ui()

    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 4}

        self._folder = FilePicker(self, label="输入文件夹", mode="folder")
        self._folder.pack(fill="x", **pad)

        self._output = FilePicker(self, label="输出目录", mode="folder")
        self._output.pack(fill="x", **pad)

        # 任务类型选择
        type_frame = ttk.Frame(self)
        type_frame.pack(fill="x", **pad)
        ttk.Label(type_frame, text="任务类型:", width=10, anchor="e").pack(side="left", padx=(0, 6))
        self._task_combo = ttk.Combobox(type_frame, values=list(TASK_TYPES.keys()),
                                        state="readonly", width=20)
        self._task_combo.set("Excel 嵌入处理")
        self._task_combo.pack(side="left")

        ttk.Button(type_frame, text="扫描文件", command=self._scan).pack(side="left", padx=(12, 0))

        # 文件列表
        tree_frame = ttk.Frame(self)
        tree_frame.pack(fill="both", expand=True, **pad)

        cols = ("file", "status")
        self._tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=8)
        self._tree.heading("file", text="文件名")
        self._tree.heading("status", text="状态")
        self._tree.column("file", width=400)
        self._tree.column("status", width=100)

        tree_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=tree_scroll.set)
        self._tree.pack(side="left", fill="both", expand=True)
        tree_scroll.pack(side="right", fill="y")

        # 进度条 + 按钮
        bottom = ttk.Frame(self)
        bottom.pack(fill="x", **pad)

        self._progress = ttk.Progressbar(bottom, mode="determinate")
        self._progress.pack(fill="x", pady=(0, 6))

        btn_frame = ttk.Frame(bottom)
        btn_frame.pack(fill="x")

        self._start_btn = ttk.Button(btn_frame, text="开始批处理",
                                     style="Primary.TButton", command=self._start)
        self._start_btn.pack(side="left", padx=(0, 8))

        self._stop_btn = ttk.Button(btn_frame, text="停止",
                                    style="Danger.TButton", command=self._stop,
                                    state="disabled")
        self._stop_btn.pack(side="left")

    def _scan(self) -> None:
        folder = self._folder.path
        if not folder:
            self._log.append("请选择输入文件夹", "WARNING")
            return

        p = Path(folder)
        if not p.is_dir():
            self._log.append("所选路径不是有效文件夹", "ERROR")
            return

        # 清空旧数据
        self._tree.delete(*self._tree.get_children())
        self._jobs.clear()

        files = sorted(p.glob("*.docx"))
        if not files:
            self._log.append("未找到 .docx 文件", "WARNING")
            return

        for f in files:
            iid = self._tree.insert("", "end", values=(f.name, "等待"))
            self._jobs.append({"path": str(f), "iid": iid, "status": "pending"})

        self._progress["maximum"] = len(self._jobs)
        self._progress["value"] = 0
        self._log.append(f"扫描到 {len(self._jobs)} 个 .docx 文件")

    def _start(self) -> None:
        if not self._jobs:
            self._log.append("请先扫描文件", "WARNING")
            return

        self._start_btn.configure(state="disabled")
        self._stop_btn.configure(state="normal")
        self._current_idx = 0
        self._log.append("批处理开始")
        self._process_next()

    def _process_next(self) -> None:
        if self._current_idx >= len(self._jobs):
            self._log.append("✅ 批处理全部完成", "SUCCESS")
            self._reset_buttons()
            return

        job = self._jobs[self._current_idx]
        task_label = self._task_combo.get()
        task_type = TASK_TYPES.get(task_label, "excel_allinone")

        self._tree.set(job["iid"], "status", "处理中")

        from ..core.adapter import build_request
        request = build_request(
            task_type=task_type,
            input_path=job["path"],
            output_dir=self._output.path or None,
        )

        self._worker.start(
            request,
            on_progress=lambda msg: self._log.append(msg),
            on_done=self._on_job_done,
            on_error=self._on_job_error,
        )

    def _on_job_done(self, response) -> None:
        job = self._jobs[self._current_idx]
        if response.ok:
            self._tree.set(job["iid"], "status", "✅ 成功")
            job["status"] = "success"
        else:
            self._tree.set(job["iid"], "status", "❌ 失败")
            job["status"] = "failed"
            err = response.error or {}
            self._log.append(
                f"  {Path(job['path']).name}: {err.get('message', '未知错误')}", "ERROR"
            )
        self._advance()

    def _on_job_error(self, msg: str) -> None:
        job = self._jobs[self._current_idx]
        self._tree.set(job["iid"], "status", "❌ 异常")
        job["status"] = "failed"
        self._log.append(f"  {Path(job['path']).name}: {msg}", "ERROR")
        self._advance()

    def _advance(self) -> None:
        self._current_idx += 1
        self._progress["value"] = self._current_idx
        if self._worker._cancel_event.is_set():
            # 标记剩余为已取消
            for j in self._jobs[self._current_idx:]:
                self._tree.set(j["iid"], "status", "已取消")
            self._log.append("批处理已取消", "WARNING")
            self._reset_buttons()
            return
        self._process_next()

    def _stop(self) -> None:
        self._worker.cancel()
        self._log.append("正在停止批处理…", "WARNING")

    def _reset_buttons(self) -> None:
        self._start_btn.configure(state="normal")
        self._stop_btn.configure(state="disabled")
