"""表格提取页 — table_extract 参数配置"""

from __future__ import annotations

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QGroupBox, QCheckBox, QLabel, QFormLayout,
)
from PyQt6.QtCore import pyqtSignal

from pyqt6.app.ui.widgets.file_picker import FilePicker
from pyqt6.app.ui.widgets.task_controls import TaskControls
from pyqt6.app.models.task_model import TaskItem


class TablePage(QWidget):
    """DOCX 表格提取"""

    sig_run_task = pyqtSignal(object)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 16, 20, 12)

        title = QLabel("DOCX 表格提取")
        title.setProperty("heading", True)
        layout.addWidget(title)

        self._input = FilePicker("选择 DOCX 文件或目录", mode="file")
        layout.addWidget(self._input)
        self._output = FilePicker("输出目录（可选）", mode="dir")
        layout.addWidget(self._output)

        group = QGroupBox("处理选项")
        form = QFormLayout(group)

        self._cb_include_marked = QCheckBox("包含已标记表格文件")
        form.addRow(self._cb_include_marked)

        layout.addWidget(group)

        self._controls = TaskControls()
        self._controls.sig_start.connect(self._on_start)
        self._controls.sig_reset.connect(self._on_reset)
        layout.addWidget(self._controls)

        layout.addStretch()

    def _gather_options(self) -> dict:
        return {
            "include_marked": self._cb_include_marked.isChecked(),
        }

    def _on_start(self):
        item = TaskItem(
            task_type="table_extract",
            input_path=self._input.path(),
            output_dir=self._output.path(),
            options=self._gather_options(),
        )
        self._controls.set_running(True)
        self.sig_run_task.emit(item)

    def _on_reset(self):
        self._controls.reset()

    @property
    def controls(self) -> TaskControls:
        return self._controls
