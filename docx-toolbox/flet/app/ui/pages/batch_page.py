"""批量任务页面"""

from __future__ import annotations

import os
from pathlib import Path

import flet as ft

from app.state.app_state import AppState
from app.runner.worker import BatchWorker
from app.config.settings import AppSettings


_TASK_TYPES = {
    "Excel 嵌入处理": "excel_allinone",
    "图片分离": "image_extract",
    "表格提取": "table_extract",
}


class BatchPage(ft.Column):
    """批量任务管理页面"""

    def __init__(self, state: AppState, batch_worker: BatchWorker, settings: AppSettings, page: ft.Page) -> None:
        super().__init__(spacing=16, expand=True)
        self._state = state
        self._batch_worker = batch_worker
        self._settings = settings
        self._page = page
        self._files: list[str] = []

        # 文件夹选择
        self._folder_picker = ft.FilePicker(on_result=self._on_folder_picked)
        page.overlay.append(self._folder_picker)
        self._folder_field = ft.TextField(label="输入文件夹", expand=True, read_only=True)
        self._output_field = ft.TextField(label="输出目录（留空则同目录）", expand=True)

        # 任务类型
        self._task_dropdown = ft.Dropdown(
            label="任务类型",
            options=[ft.dropdown.Option(key=v, text=k) for k, v in _TASK_TYPES.items()],
            value="excel_allinone",
            width=220,
        )

        # 文件列表
        self._data_table = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("文件名")),
                ft.DataColumn(ft.Text("状态")),
            ],
            rows=[],
            expand=True,
        )

        # 进度
        self._progress = ft.ProgressBar(value=0, visible=False, expand=True)
        self._progress_text = ft.Text("")

        # 按钮
        self._btn_start = ft.ElevatedButton("开始批量处理", icon=ft.Icons.PLAY_ARROW, on_click=self._on_start)
        self._btn_stop = ft.OutlinedButton("停止", icon=ft.Icons.STOP, on_click=self._on_stop, disabled=True)

        self._build()

    def _build(self) -> None:
        self.controls = [
            ft.Text("批量任务", size=20, weight=ft.FontWeight.BOLD),
            ft.Divider(height=1),
            ft.Row([
                self._folder_field,
                ft.ElevatedButton("选择文件夹", icon=ft.Icons.FOLDER, on_click=lambda _: self._folder_picker.get_directory_path(
                    dialog_title="选择包含 DOCX 文件的文件夹"
                )),
            ]),
            ft.Row([self._task_dropdown, self._output_field]),
            ft.Container(content=self._data_table, expand=True, border=ft.border.all(1, "#E5E7EB"), border_radius=8, padding=8),
            ft.Row([self._progress_text, self._progress]),
            ft.Row([self._btn_start, self._btn_stop]),
        ]

    def _on_folder_picked(self, e: ft.FilePickerResultEvent) -> None:
        if e.path:
            self._folder_field.value = e.path
            self._scan_files(e.path)
            self._page.update()

    def _scan_files(self, folder: str) -> None:
        """扫描文件夹中的 .docx 文件"""
        p = Path(folder)
        self._files = []
        self._data_table.rows.clear()
        if p.is_dir():
            for f in sorted(p.iterdir(), key=lambda x: x.name.lower()):
                if f.suffix.lower() == ".docx" and not f.name.startswith("~$"):
                    self._files.append(str(f))
                    self._data_table.rows.append(ft.DataRow(cells=[
                        ft.DataCell(ft.Text(f.name)),
                        ft.DataCell(ft.Text("待处理")),
                    ]))
        self._progress_text.value = f"共 {len(self._files)} 个文件"

    def _on_start(self, _: ft.ControlEvent) -> None:
        if not self._files:
            self._page.open(ft.SnackBar(ft.Text("请先选择包含 DOCX 文件的文件夹"), open=True))
            self._page.update()
            return
        self._btn_start.disabled = True
        self._btn_stop.disabled = False
        self._progress.visible = True
        self._progress.value = 0
        self._page.update()

        self._batch_worker.run_batch(
            task_type=self._task_dropdown.value,
            file_paths=list(self._files),
            output_dir=self._output_field.value or None,
            options={},
            workers=self._settings.worker_count,
            on_done=self._on_batch_done,
        )

    def _on_stop(self, _: ft.ControlEvent) -> None:
        self._batch_worker.cancel()
        self._state.add_log("[批量] 用户请求取消")

    def _on_batch_done(self, jobs) -> None:
        self._btn_start.disabled = False
        self._btn_stop.disabled = True
        if jobs:
            for i, job in enumerate(jobs):
                if i < len(self._data_table.rows):
                    status_cell = self._data_table.rows[i].cells[1]
                    status_cell.content.value = job.status.value
            self._progress.value = 1.0
        self._page.update()
