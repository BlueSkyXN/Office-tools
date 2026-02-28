"""DOCX 工具箱 — Flet GUI 入口

运行方式: cd docx-toolbox && python3 flet/app/main.py
"""

import sys
import os
from pathlib import Path

# 将 docx-toolbox 根目录加入 sys.path，使 core 包可导入
_project_root = str(Path(__file__).resolve().parent.parent.parent)
if _project_root not in sys.path:
    sys.path.insert(0, _project_root)

# 将 flet/  加入 sys.path，使 app 包可导入
_flet_root = str(Path(__file__).resolve().parent.parent)
if _flet_root not in sys.path:
    sys.path.insert(0, _flet_root)

import flet as ft

from app.config.theme import create_theme
from app.config.settings import AppSettings
from app.state.app_state import AppState
from app.ui.app_layout import AppLayout


def main(page: ft.Page) -> None:
    page.title = "DOCX 工具箱"
    page.theme = create_theme()
    page.window.width = 960
    page.window.height = 700
    page.window.min_width = 800
    page.window.min_height = 600
    page.padding = 0

    settings = AppSettings.load()
    state = AppState()

    layout = AppLayout(page, state, settings)
    page.add(layout.build())
    page.update()


if __name__ == "__main__":
    ft.run(main)
