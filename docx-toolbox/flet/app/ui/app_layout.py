"""主布局：NavigationRail + 内容区 + 底部日志面板"""

from __future__ import annotations

import flet as ft

from app.state.app_state import AppState
from app.config.settings import AppSettings
from app.config.theme import COLOR_SUCCESS, COLOR_WARNING, COLOR_ERROR
from app.runner.worker import Worker, BatchWorker
from app.ui.pages.excel_page import ExcelPage
from app.ui.pages.image_page import ImagePage
from app.ui.pages.table_page import TablePage
from app.ui.pages.batch_page import BatchPage
from app.ui.pages.settings_page import SettingsPage


class AppLayout:
    """构建并管理整体布局"""

    def __init__(self, page: ft.Page, state: AppState, settings: AppSettings) -> None:
        self._page = page
        self._state = state
        self._settings = settings

        self._worker = Worker(state, page_update=self._safe_update)
        self._batch_worker = BatchWorker(state, page_update=self._safe_update)

        # 页面实例
        self._pages: list[ft.Control] = [
            ExcelPage(state, self._worker, page),
            ImagePage(state, self._worker, page),
            TablePage(state, self._worker, page),
            BatchPage(state, self._batch_worker, settings, page),
            SettingsPage(settings, page),
        ]

        # 导航栏
        self._nav_rail = ft.NavigationRail(
            selected_index=0,
            label_type=ft.NavigationRailLabelType.ALL,
            min_width=100,
            min_extended_width=200,
            destinations=[
                ft.NavigationRailDestination(icon=ft.Icons.TABLE_CHART, label="Excel处理"),
                ft.NavigationRailDestination(icon=ft.Icons.IMAGE, label="图片分离"),
                ft.NavigationRailDestination(icon=ft.Icons.GRID_ON, label="表格提取"),
                ft.NavigationRailDestination(icon=ft.Icons.BATCH_PREDICTION, label="批量任务"),
                ft.NavigationRailDestination(icon=ft.Icons.SETTINGS, label="设置"),
            ],
            on_change=self._on_nav_change,
        )

        # 内容区
        self._content = ft.Container(
            content=self._pages[0],
            expand=True,
            padding=24,
        )

        # 日志面板
        self._log_view = ft.ListView(spacing=2, auto_scroll=True, height=160)
        state.on_log(self._append_log)

        self._log_panel = ft.Container(
            content=ft.Column([
                ft.Row([
                    ft.Text("日志", size=12, weight=ft.FontWeight.W_600),
                    ft.IconButton(icon=ft.Icons.DELETE_SWEEP, tooltip="清空日志", icon_size=16, on_click=self._clear_logs),
                ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                ft.Container(content=self._log_view, expand=True),
            ], spacing=4),
            padding=ft.padding.symmetric(horizontal=16, vertical=8),
            border=ft.border.only(top=ft.BorderSide(1, "#E5E7EB")),
        )

    def build(self) -> ft.Control:
        """返回完整布局控件"""
        return ft.Column([
            ft.Row([
                self._nav_rail,
                ft.VerticalDivider(width=1),
                self._content,
            ], expand=True),
            self._log_panel,
        ], expand=True)

    def _on_nav_change(self, e: ft.ControlEvent) -> None:
        idx = int(e.data)
        self._state.current_page = idx
        self._content.content = self._pages[idx]
        self._page.update()

    def _append_log(self, line: str) -> None:
        color = None
        if "失败" in line or "异常" in line or "错误" in line:
            color = COLOR_ERROR
        elif "完成" in line or "成功" in line:
            color = COLOR_SUCCESS
        elif "警告" in line:
            color = COLOR_WARNING
        self._log_view.controls.append(ft.Text(line, size=11, color=color))
        # 限制 UI 中的日志条数
        if len(self._log_view.controls) > self._settings.log_max_lines:
            self._log_view.controls = self._log_view.controls[-self._settings.log_max_lines:]

    def _clear_logs(self, _: ft.ControlEvent) -> None:
        self._log_view.controls.clear()
        self._page.update()

    def _safe_update(self) -> None:
        """线程安全的 page.update()"""
        try:
            self._page.update()
        except Exception:
            pass
