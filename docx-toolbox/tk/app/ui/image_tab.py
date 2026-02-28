"""Tab2: 图片分离"""

from __future__ import annotations

import tkinter as tklib
from tkinter import ttk
from typing import TYPE_CHECKING

from .widgets.file_picker import FilePicker

if TYPE_CHECKING:
    from ..runner.worker import TaskWorker
    from .widgets.log_text import LogText


class ImageTab(ttk.Frame):
    """图片分离标签页"""

    def __init__(self, master: tklib.Misc, log: LogText, worker: TaskWorker, **kwargs):
        super().__init__(master, **kwargs)
        self._log = log
        self._worker = worker
        self._build_ui()

    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 4}

        self._input = FilePicker(self, label="输入文件", mode="file")
        self._input.pack(fill="x", **pad)

        self._output = FilePicker(self, label="输出目录", mode="folder")
        self._output.pack(fill="x", **pad)

        # 选项
        opts_frame = ttk.LabelFrame(self, text="处理选项", padding=10)
        opts_frame.pack(fill="x", **pad)

        self._remove_images = tklib.BooleanVar(value=False)
        self._optimize_images = tklib.BooleanVar(value=False)
        self._jpeg_quality = tklib.IntVar(value=85)

        ttk.Checkbutton(opts_frame, text="删除原图仅保留标记",
                        variable=self._remove_images).grid(
            row=0, column=0, sticky="w", padx=8, pady=3)
        ttk.Checkbutton(opts_frame, text="启用图片优化",
                        variable=self._optimize_images).grid(
            row=0, column=1, sticky="w", padx=8, pady=3)

        quality_frame = ttk.Frame(opts_frame)
        quality_frame.grid(row=1, column=0, columnspan=2, sticky="w", padx=8, pady=3)
        ttk.Label(quality_frame, text="JPEG 质量:").pack(side="left")
        ttk.Spinbox(quality_frame, from_=1, to=100, width=6,
                    textvariable=self._jpeg_quality).pack(side="left", padx=(4, 0))

        # 按钮
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
            "remove_images": self._remove_images.get(),
            "optimize_images": self._optimize_images.get(),
            "jpeg_quality": self._jpeg_quality.get(),
        }

    def _start(self) -> None:
        input_path = self._input.path
        if not input_path:
            self._log.append("请选择输入文件", "WARNING")
            return

        from ..core.adapter import build_request
        request = build_request(
            task_type="image_extract",
            input_path=input_path,
            output_dir=self._output.path or None,
            options=self._get_options(),
        )

        self._start_btn.configure(state="disabled")
        self._stop_btn.configure(state="normal")
        self._log.append(f"图片分离任务已提交: {input_path}")

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
