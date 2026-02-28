"""docx-toolbox Tkinter 版 — 主入口

启动方式: cd docx-toolbox && python3 -m tk.app.main
"""

from __future__ import annotations

import sys
from pathlib import Path

# 确保 docx-toolbox 根目录在 sys.path 中，以便导入顶层 core 包
_DOCX_ROOT = str(Path(__file__).resolve().parent.parent.parent)
if _DOCX_ROOT not in sys.path:
    sys.path.insert(0, _DOCX_ROOT)

import tkinter as tklib
from tkinter import ttk

from tk.app.config.theme import apply_theme
from tk.app.config.settings import Settings
from tk.app.ui.widgets.log_text import LogText
from tk.app.runner.worker import TaskWorker
from tk.app.ui.excel_tab import ExcelTab
from tk.app.ui.image_tab import ImageTab
from tk.app.ui.table_tab import TableTab
from tk.app.ui.batch_tab import BatchTab


APP_TITLE = "docx-toolbox"
MIN_WIDTH = 820
MIN_HEIGHT = 620


def main() -> None:
    root = tklib.Tk()
    root.title(APP_TITLE)
    root.minsize(MIN_WIDTH, MIN_HEIGHT)

    # 主题
    style = ttk.Style(root)
    apply_theme(style)
    root.configure(bg="#F8FAFC")

    # 恢复窗口位置
    settings = Settings()
    geometry = settings.get("window_geometry")
    if geometry:
        root.geometry(geometry)
    else:
        root.geometry(f"{MIN_WIDTH}x{MIN_HEIGHT}")

    # 共享组件
    worker = TaskWorker(root)
    log = LogText(root)

    # Notebook
    notebook = ttk.Notebook(root)

    excel_tab = ExcelTab(notebook, log=log, worker=worker)
    image_tab = ImageTab(notebook, log=log, worker=worker)
    table_tab = TableTab(notebook, log=log, worker=worker)
    batch_tab = BatchTab(notebook, log=log, worker=worker)

    notebook.add(excel_tab, text=" Excel 嵌入处理 ")
    notebook.add(image_tab, text=" 图片分离 ")
    notebook.add(table_tab, text=" 表格提取 ")
    notebook.add(batch_tab, text=" 批处理 ")

    # 布局：Notebook 扩展占满上方，LogText 固定在底部
    notebook.pack(fill="both", expand=True, padx=8, pady=(8, 4))
    log.pack(fill="x", padx=8, pady=(4, 8))

    log.append("docx-toolbox 已启动", "SUCCESS")

    # 关闭时保存窗口位置
    def _on_close() -> None:
        settings.set("window_geometry", root.geometry())
        settings.save()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", _on_close)
    root.mainloop()


if __name__ == "__main__":
    main()
