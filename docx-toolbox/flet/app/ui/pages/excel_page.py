"""Excel 嵌入处理页面"""

from __future__ import annotations

import flet as ft

from app.state.app_state import AppState
from app.runner.worker import Worker


class ExcelPage(ft.Column):
    """excel_allinone 任务页面"""

    def __init__(self, state: AppState, worker: Worker, page: ft.Page) -> None:
        super().__init__(spacing=16, expand=True)
        self._state = state
        self._worker = worker
        self._page = page

        # 文件选择
        self._file_picker = ft.FilePicker(on_result=self._on_file_picked)
        page.overlay.append(self._file_picker)
        self._input_field = ft.TextField(label="输入文件 (.docx)", expand=True, read_only=True)
        self._output_field = ft.TextField(label="输出目录（留空则同目录）", expand=True)

        # 选项
        self._opt_word_table = ft.Checkbox(label="转换为 Word 原生表格", value=True)
        self._opt_extract_excel = ft.Checkbox(label="提取嵌入 Excel 文件", value=True)
        self._opt_image = ft.Checkbox(label="渲染为图片", value=False)
        self._opt_keep_attachment = ft.Checkbox(label="保留附件入口", value=False)
        self._opt_remove_watermark = ft.Checkbox(label="移除水印", value=False)
        self._opt_a3 = ft.Checkbox(label="设置 A3 横向", value=False)

        # 按钮
        self._btn_start = ft.ElevatedButton("开始处理", icon=ft.Icons.PLAY_ARROW, on_click=self._on_start)
        self._btn_reset = ft.OutlinedButton("重置", icon=ft.Icons.REFRESH, on_click=self._on_reset)

        # 进度
        self._progress = ft.ProgressBar(visible=False, expand=True)

        self._build()

    def _build(self) -> None:
        self.controls = [
            ft.Text("Excel 嵌入处理", size=20, weight=ft.FontWeight.BOLD),
            ft.Divider(height=1),
            ft.Row([
                self._input_field,
                ft.ElevatedButton("选择文件", icon=ft.Icons.FOLDER_OPEN, on_click=lambda _: self._file_picker.pick_files(
                    allowed_extensions=["docx"], dialog_title="选择 DOCX 文件"
                )),
            ]),
            self._output_field,
            ft.Text("处理选项", size=14, weight=ft.FontWeight.W_600),
            ft.Row([self._opt_word_table, self._opt_extract_excel, self._opt_image], wrap=True),
            ft.Row([self._opt_keep_attachment, self._opt_remove_watermark, self._opt_a3], wrap=True),
            ft.Divider(height=1),
            ft.Row([self._btn_start, self._btn_reset]),
            self._progress,
        ]

    def _on_file_picked(self, e: ft.FilePickerResultEvent) -> None:
        if e.files:
            self._input_field.value = e.files[0].path
            self._page.update()

    def _gather_options(self) -> dict:
        return {
            "word_table": self._opt_word_table.value,
            "extract_excel": self._opt_extract_excel.value,
            "image": self._opt_image.value,
            "keep_attachment": self._opt_keep_attachment.value,
            "remove_watermark": self._opt_remove_watermark.value,
            "a3": self._opt_a3.value,
        }

    def _set_locked(self, locked: bool) -> None:
        for ctrl in [self._opt_word_table, self._opt_extract_excel, self._opt_image,
                     self._opt_keep_attachment, self._opt_remove_watermark, self._opt_a3,
                     self._btn_start]:
            ctrl.disabled = locked
        self._progress.visible = locked
        self._page.update()

    def _on_start(self, _: ft.ControlEvent) -> None:
        if not self._input_field.value:
            self._page.open(ft.SnackBar(ft.Text("请先选择文件"), open=True))
            self._page.update()
            return
        self._set_locked(True)
        self._worker.run_single(
            task_key="excel",
            task_type="excel_allinone",
            input_path=self._input_field.value,
            output_dir=self._output_field.value or None,
            options=self._gather_options(),
            on_done=lambda _resp: self._set_locked(False),
        )

    def _on_reset(self, _: ft.ControlEvent) -> None:
        self._input_field.value = ""
        self._output_field.value = ""
        self._opt_word_table.value = True
        self._opt_extract_excel.value = True
        self._opt_image.value = False
        self._opt_keep_attachment.value = False
        self._opt_remove_watermark.value = False
        self._opt_a3.value = False
        self._set_locked(False)
