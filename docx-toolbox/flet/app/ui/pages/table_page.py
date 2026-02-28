"""表格提取页面"""

from __future__ import annotations

import flet as ft

from app.state.app_state import AppState
from app.runner.worker import Worker


class TablePage(ft.Column):
    """table_extract 任务页面"""

    def __init__(self, state: AppState, worker: Worker, page: ft.Page) -> None:
        super().__init__(spacing=16, expand=True)
        self._state = state
        self._worker = worker
        self._page = page

        # 文件选择
        self._file_picker = ft.FilePicker()
        page.services.append(self._file_picker)
        self._input_field = ft.TextField(label="输入文件 (.docx)", expand=True, read_only=True)
        self._output_field = ft.TextField(label="输出目录（留空则同目录）", expand=True)

        # 选项
        self._opt_include_marked = ft.Checkbox(label="包含已标记表格文件", value=False)

        # 按钮
        self._btn_start = ft.ElevatedButton("开始处理", icon=ft.Icons.PLAY_ARROW, on_click=self._on_start)
        self._btn_reset = ft.OutlinedButton("重置", icon=ft.Icons.REFRESH, on_click=self._on_reset)

        # 进度
        self._progress = ft.ProgressBar(visible=False, expand=True)

        self._build()

    def _build(self) -> None:
        self.controls = [
            ft.Text("表格提取", size=20, weight=ft.FontWeight.BOLD),
            ft.Divider(height=1),
            ft.Row([
                self._input_field,
                ft.ElevatedButton("选择文件", icon=ft.Icons.FOLDER_OPEN, on_click=self._on_pick_file),
            ]),
            self._output_field,
            ft.Text("处理选项", size=14, weight=ft.FontWeight.W_600),
            self._opt_include_marked,
            ft.Divider(height=1),
            ft.Row([self._btn_start, self._btn_reset]),
            self._progress,
        ]

    async def _on_pick_file(self, _: ft.ControlEvent) -> None:
        files = await self._file_picker.pick_files(
            dialog_title="选择 DOCX 文件",
            file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=["docx"],
            allow_multiple=False,
        )
        if files and files[0].path:
            self._input_field.value = files[0].path
            self._page.update()

    def _gather_options(self) -> dict:
        return {
            "include_marked": self._opt_include_marked.value,
        }

    def _set_locked(self, locked: bool) -> None:
        self._opt_include_marked.disabled = locked
        self._btn_start.disabled = locked
        self._progress.visible = locked
        self._page.update()

    def _on_start(self, _: ft.ControlEvent) -> None:
        if not self._input_field.value:
            self._page.open(ft.SnackBar(ft.Text("请先选择文件"), open=True))
            self._page.update()
            return
        self._set_locked(True)
        self._worker.run_single(
            task_key="table",
            task_type="table_extract",
            input_path=self._input_field.value,
            output_dir=self._output_field.value or None,
            options=self._gather_options(),
            on_done=lambda _resp: self._set_locked(False),
        )

    def _on_reset(self, _: ft.ControlEvent) -> None:
        self._input_field.value = ""
        self._output_field.value = ""
        self._opt_include_marked.value = False
        self._set_locked(False)
