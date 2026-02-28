"""Tab1: Excel 嵌入处理"""

from __future__ import annotations

import tkinter as tklib
from tkinter import ttk
from typing import TYPE_CHECKING

from .widgets.file_picker import FilePicker

if TYPE_CHECKING:
    from ..runner.worker import TaskWorker
    from .widgets.log_text import LogText


class ExcelTab(ttk.Frame):
    """Excel 嵌入对象 all-in-one 处理标签页"""

    def __init__(self, master: tklib.Misc, log: LogText, worker: TaskWorker, **kwargs):
        super().__init__(master, **kwargs)
        self._log = log
        self._worker = worker
        self._build_ui()

    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 4}

        # 输入文件
        self._input = FilePicker(self, label="输入文件", mode="file")
        self._input.pack(fill="x", **pad)

        # 输出目录
        self._output = FilePicker(self, label="输出目录", mode="folder")
        self._output.pack(fill="x", **pad)

        # 选项区域
        opts_frame = ttk.LabelFrame(self, text="处理选项", padding=10)
        opts_frame.pack(fill="x", **pad)

        self._word_table = tklib.BooleanVar(value=True)
        self._extract_excel = tklib.BooleanVar(value=True)
        self._image = tklib.BooleanVar(value=False)
        self._keep_attachment = tklib.BooleanVar(value=False)
        self._remove_watermark = tklib.BooleanVar(value=False)
        self._a3 = tklib.BooleanVar(value=False)

        checks = [
            ("转换为 Word 原生表格", self._word_table),
            ("提取嵌入 Excel 文件", self._extract_excel),
            ("渲染为图片", self._image),
            ("保留附件入口", self._keep_attachment),
            ("移除水印", self._remove_watermark),
            ("设置 A3 横向", self._a3),
        ]
        for i, (text, var) in enumerate(checks):
            row, col = divmod(i, 3)
            ttk.Checkbutton(opts_frame, text=text, variable=var).grid(
                row=row, column=col, sticky="w", padx=8, pady=3)

        # 按钮区域
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill="x", **pad)

        self._start_btn = ttk.Button(btn_frame, text="开始处理",
                                     style="Primary.TButton", command=self._start)
        self._start_btn.pack(side="left", padx=(0, 8))

        self._stop_btn = ttk.Button(btn_frame, text="停止",
                                    style="Danger.TButton", command=self._stop,
                                    state="disabled")
        self._stop_btn.pack(side="left")

    def _get_options(self) -> dict:
        return {
            "word_table": self._word_table.get(),
            "extract_excel": self._extract_excel.get(),
            "image": self._image.get(),
            "keep_attachment": self._keep_attachment.get(),
            "remove_watermark": self._remove_watermark.get(),
            "a3": self._a3.get(),
        }

    def _start(self) -> None:
        input_path = self._input.path
        if not input_path:
            self._log.append("请选择输入文件", "WARNING")
            return

        from ..core.adapter import build_request
        request = build_request(
            task_type="excel_allinone",
            input_path=input_path,
            output_dir=self._output.path or None,
            options=self._get_options(),
        )

        self._start_btn.configure(state="disabled")
        self._stop_btn.configure(state="normal")
        self._log.append(f"Excel 处理任务已提交: {input_path}")

        self._worker.start(
            request,
            on_progress=lambda msg: self._log.append(msg),
            on_done=self._on_done,
            on_error=self._on_error,
        )

    def _stop(self) -> None:
        self._worker.cancel()
        self._log.append("正在停止任务…", "WARNING")
        self._reset_buttons()

    def _on_done(self, response) -> None:
        if response.ok:
            s = response.summary
            self._log.append(
                f"✅ 完成: 处理 {s.processed} 个, 失败 {s.failed} 个, 跳过 {s.skipped} 个",
                "SUCCESS",
            )
        else:
            err = response.error or {}
            self._log.append(f"❌ 失败: {err.get('message', '未知错误')}", "ERROR")
        self._reset_buttons()

    def _on_error(self, msg: str) -> None:
        self._log.append(f"❌ 异常: {msg}", "ERROR")
        self._reset_buttons()

    def _reset_buttons(self) -> None:
        self._start_btn.configure(state="normal")
        self._stop_btn.configure(state="disabled")
