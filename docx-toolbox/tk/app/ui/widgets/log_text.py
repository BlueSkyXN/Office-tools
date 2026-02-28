"""共享日志文本组件 — 底部日志框，只读、可复制/保存"""

from __future__ import annotations

import tkinter as tklib
from tkinter import ttk, filedialog
from datetime import datetime

MAX_LINES = 500


class LogText(ttk.Frame):
    """只读日志文本框，带滚动条和工具栏"""

    def __init__(self, master: tklib.Misc, **kwargs):
        super().__init__(master, **kwargs)

        toolbar = ttk.Frame(self)
        toolbar.pack(side="top", fill="x", pady=(0, 2))

        ttk.Label(toolbar, text="日志", style="Title.TLabel").pack(side="left")
        ttk.Button(toolbar, text="保存", command=self._save, width=6).pack(side="right", padx=2)
        ttk.Button(toolbar, text="复制", command=self._copy, width=6).pack(side="right", padx=2)
        ttk.Button(toolbar, text="清空", command=self.clear, width=6).pack(side="right", padx=2)

        container = ttk.Frame(self)
        container.pack(fill="both", expand=True)

        self._text = tklib.Text(container, height=8, wrap="word", state="disabled",
                                font=("Menlo", 11), bg="#FFFFFF", fg="#111827",
                                relief="solid", bd=1, padx=6, pady=4)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=self._text.yview)
        self._text.configure(yscrollcommand=scrollbar.set)

        self._text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 标签颜色
        self._text.tag_configure("INFO", foreground="#111827")
        self._text.tag_configure("WARNING", foreground="#D97706")
        self._text.tag_configure("ERROR", foreground="#DC2626")
        self._text.tag_configure("SUCCESS", foreground="#16A34A")

    def append(self, message: str, level: str = "INFO") -> None:
        """追加一行日志"""
        self._text.configure(state="normal")
        timestamp = datetime.now().strftime("%H:%M:%S")
        line = f"[{timestamp}] {message}\n"
        self._text.insert("end", line, level.upper())
        self._trim()
        self._text.configure(state="disabled")
        self._text.see("end")

    def clear(self) -> None:
        self._text.configure(state="normal")
        self._text.delete("1.0", "end")
        self._text.configure(state="disabled")

    def _trim(self) -> None:
        """保留最近 MAX_LINES 行"""
        line_count = int(self._text.index("end-1c").split(".")[0])
        if line_count > MAX_LINES:
            self._text.delete("1.0", f"{line_count - MAX_LINES}.0")

    def _copy(self) -> None:
        content = self._text.get("1.0", "end-1c")
        self._text.clipboard_clear()
        self._text.clipboard_append(content)

    def _save(self) -> None:
        path = filedialog.asksaveasfilename(
            defaultextension=".log",
            filetypes=[("日志文件", "*.log"), ("文本文件", "*.txt")],
            initialfile=f"docx-toolbox-{datetime.now().strftime('%Y%m%d_%H%M%S')}.log",
        )
        if path:
            with open(path, "w", encoding="utf-8") as f:
                f.write(self._text.get("1.0", "end-1c"))
