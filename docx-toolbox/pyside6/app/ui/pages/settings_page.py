"""设置页"""

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QGroupBox, QFormLayout, QLabel, QSpinBox,
)
from PySide6.QtCore import Signal

from pyside6.app.ui.widgets import FilePicker
from pyside6.app.config.settings import AppSettings
from pyside6.app.config.theme import COLORS


class SettingsPage(QWidget):
    log_message = Signal(str)
    status_message = Signal(str)

    def __init__(self, settings: AppSettings, parent=None):
        super().__init__(parent)
        self._settings = settings
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 12)

        title = QLabel("设置")
        title.setStyleSheet("font-size: 18px; font-weight: 700; margin-bottom: 4px;")
        layout.addWidget(title)

        # General
        gen_group = QGroupBox("通用")
        gl = QFormLayout(gen_group)

        self._output_picker = FilePicker("默认输出目录", mode="dir")
        self._output_picker.set_path(self._settings.default_output_dir)
        self._output_picker.path_changed.connect(self._save_output_dir)
        gl.addRow("默认输出目录:", self._output_picker)

        self._spin_workers = QSpinBox()
        self._spin_workers.setRange(1, 8)
        self._spin_workers.setValue(self._settings.worker_count)
        self._spin_workers.valueChanged.connect(self._save_workers)
        gl.addRow("并发线程数:", self._spin_workers)
        layout.addWidget(gen_group)

        # Theme preview
        theme_group = QGroupBox("主题预览")
        tl = QVBoxLayout(theme_group)
        for name, color in COLORS.items():
            row = QLabel(f"  {name}: {color}")
            row.setStyleSheet(f"background-color: {color}; padding: 4px 8px; border-radius: 4px; color: {'#fff' if name in ('text', 'log_bg') else '#111'};")
            tl.addWidget(row)
        layout.addWidget(theme_group)

        layout.addStretch()

    def _save_output_dir(self, path: str):
        self._settings.default_output_dir = path

    def _save_workers(self, val: int):
        self._settings.worker_count = val
