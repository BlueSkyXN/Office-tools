"""设置页面"""

from __future__ import annotations

import flet as ft

from app.config.settings import AppSettings


class SettingsPage(ft.Column):
    """应用设置页面"""

    def __init__(self, settings: AppSettings, page: ft.Page) -> None:
        super().__init__(spacing=16, expand=True)
        self._settings = settings
        self._page = page

        # 输出目录
        self._dir_picker = ft.FilePicker(on_result=self._on_dir_picked)
        page.overlay.append(self._dir_picker)
        self._output_dir_field = ft.TextField(
            label="默认输出目录",
            value=settings.default_output_dir,
            expand=True,
        )

        # Worker 数量
        self._worker_slider = ft.Slider(
            min=1, max=8, value=settings.worker_count,
            divisions=7, label="并发数: {value}",
        )
        self._worker_label = ft.Text(f"并发 Worker 数: {settings.worker_count}")
        self._worker_slider.on_change = self._on_worker_change

        # 保存按钮
        self._btn_save = ft.ElevatedButton("保存设置", icon=ft.Icons.SAVE, on_click=self._on_save)

        self._build()

    def _build(self) -> None:
        self.controls = [
            ft.Text("设置", size=20, weight=ft.FontWeight.BOLD),
            ft.Divider(height=1),
            ft.Text("默认输出目录", size=14, weight=ft.FontWeight.W_600),
            ft.Row([
                self._output_dir_field,
                ft.ElevatedButton("选择", icon=ft.Icons.FOLDER, on_click=lambda _: self._dir_picker.get_directory_path(
                    dialog_title="选择默认输出目录"
                )),
            ]),
            ft.Divider(height=1),
            ft.Text("并发设置", size=14, weight=ft.FontWeight.W_600),
            self._worker_label,
            self._worker_slider,
            ft.Divider(height=1),
            self._btn_save,
        ]

    def _on_dir_picked(self, e: ft.FilePickerResultEvent) -> None:
        if e.path:
            self._output_dir_field.value = e.path
            self._page.update()

    def _on_worker_change(self, _: ft.ControlEvent) -> None:
        self._worker_label.value = f"并发 Worker 数: {int(self._worker_slider.value)}"
        self._page.update()

    def _on_save(self, _: ft.ControlEvent) -> None:
        self._settings.default_output_dir = self._output_dir_field.value or ""
        self._settings.worker_count = int(self._worker_slider.value)
        self._settings.save()
        self._page.open(ft.SnackBar(ft.Text("设置已保存"), open=True))
        self._page.update()
