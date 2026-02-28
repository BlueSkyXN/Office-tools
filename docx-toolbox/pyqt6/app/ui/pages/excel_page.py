"""Excel 处理页 — excel_allinone 参数配置"""

from __future__ import annotations

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QGroupBox, QCheckBox, QLabel, QFormLayout,
)
from PyQt6.QtCore import pyqtSignal

from pyqt6.app.ui.widgets.file_picker import FilePicker
from pyqt6.app.ui.widgets.task_controls import TaskControls
from pyqt6.app.models.task_model import TaskItem


class ExcelPage(QWidget):
    """Excel 嵌入对象处理"""

    sig_run_task = pyqtSignal(object)  # TaskItem

    def __init__(self, parent=None):
        super().__init__(parent)
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 16, 20, 12)

        title = QLabel("Excel 嵌入对象处理")
        title.setProperty("heading", True)
        layout.addWidget(title)

        # 文件选择
        self._input = FilePicker("选择 DOCX 文件或目录", mode="file")
        layout.addWidget(self._input)
        self._output = FilePicker("输出目录（可选）", mode="dir")
        layout.addWidget(self._output)

        # 参数组
        group = QGroupBox("处理选项")
        form = QFormLayout(group)

        self._cb_word_table = QCheckBox("转换为 Word 原生表格")
        self._cb_word_table.setChecked(True)
        form.addRow(self._cb_word_table)

        self._cb_extract_excel = QCheckBox("提取嵌入 Excel 文件")
        self._cb_extract_excel.setChecked(True)
        form.addRow(self._cb_extract_excel)

        self._cb_image = QCheckBox("渲染为图片")
        self._cb_image.setChecked(True)
        form.addRow(self._cb_image)

        self._cb_keep_attachment = QCheckBox("保留附件入口")
        form.addRow(self._cb_keep_attachment)

        self._cb_remove_watermark = QCheckBox("移除水印")
        form.addRow(self._cb_remove_watermark)

        self._cb_a3 = QCheckBox("A3 横向")
        form.addRow(self._cb_a3)

        layout.addWidget(group)

        # 控制按钮
        self._controls = TaskControls()
        self._controls.sig_start.connect(self._on_start)
        self._controls.sig_reset.connect(self._on_reset)
        layout.addWidget(self._controls)

        layout.addStretch()

    def _gather_options(self) -> dict:
        return {
            "word_table": self._cb_word_table.isChecked(),
            "extract_excel": self._cb_extract_excel.isChecked(),
            "image": self._cb_image.isChecked(),
            "keep_attachment": self._cb_keep_attachment.isChecked(),
            "remove_watermark": self._cb_remove_watermark.isChecked(),
            "a3": self._cb_a3.isChecked(),
        }

    def _on_start(self):
        item = TaskItem(
            task_type="excel_allinone",
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
