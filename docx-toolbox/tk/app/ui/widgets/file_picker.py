"""可复用文件/文件夹选择器"""

from __future__ import annotations

import tkinter as tklib
from tkinter import ttk, filedialog
from pathlib import Path


class FilePicker(ttk.Frame):
    """包含路径输入框和浏览按钮的文件/文件夹选择器"""

    def __init__(self, master: tklib.Misc, *, label: str = "路径",
                 mode: str = "file", filetypes: list[tuple[str, str]] | None = None,
                 **kwargs):
        """
        Args:
            mode: "file" | "folder" | "files"
            filetypes: 文件过滤器，仅 mode="file" 时有效
        """
        super().__init__(master, **kwargs)
        self._mode = mode
        self._filetypes = filetypes or [("Word 文档", "*.docx"), ("所有文件", "*.*")]

        self._var = tklib.StringVar()

        ttk.Label(self, text=label, width=10, anchor="e").grid(
            row=0, column=0, padx=(0, 6), sticky="e")

        self._entry = ttk.Entry(self, textvariable=self._var, width=50)
        self._entry.grid(row=0, column=1, sticky="ew", padx=(0, 6))

        ttk.Button(self, text="浏览…", command=self._browse, width=8).grid(
            row=0, column=2)

        self.columnconfigure(1, weight=1)

    @property
    def path(self) -> str:
        return self._var.get().strip()

    @path.setter
    def path(self, value: str) -> None:
        self._var.set(value)

    def _browse(self) -> None:
        if self._mode == "folder":
            result = filedialog.askdirectory(title=f"选择文件夹")
        elif self._mode == "files":
            result = filedialog.askopenfilenames(
                title="选择文件", filetypes=self._filetypes)
            if result:
                result = ";".join(result)
        else:
            result = filedialog.askopenfilename(
                title="选择文件", filetypes=self._filetypes)
        if result:
            self._var.set(result)
