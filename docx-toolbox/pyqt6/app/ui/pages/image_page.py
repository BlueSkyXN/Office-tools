"""图片分离页 — image_extract 参数配置"""

from __future__ import annotations

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QGroupBox, QCheckBox, QLabel,
    QFormLayout, QSpinBox, QSlider, QHBoxLayout,
)
from PyQt6.QtCore import Qt, pyqtSignal

from pyqt6.app.ui.widgets.file_picker import FilePicker
from pyqt6.app.ui.widgets.task_controls import TaskControls
from pyqt6.app.models.task_model import TaskItem


class ImagePage(QWidget):
    """DOCX 图片分离"""

    sig_run_task = pyqtSignal(object)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 16, 20, 12)

        title = QLabel("DOCX 图片分离")
        title.setProperty("heading", True)
        layout.addWidget(title)

        self._input = FilePicker("选择 DOCX 文件或目录", mode="file")
        layout.addWidget(self._input)
        self._output = FilePicker("输出目录（可选）", mode="dir")
        layout.addWidget(self._output)

        group = QGroupBox("处理选项")
        form = QFormLayout(group)

        self._cb_remove = QCheckBox("删除原图（仅保留标记）")
        form.addRow(self._cb_remove)

        self._cb_optimize = QCheckBox("启用图片优化")
        self._cb_optimize.setChecked(True)
        form.addRow(self._cb_optimize)

        # JPEG 质量
        quality_row = QHBoxLayout()
        self._slider_quality = QSlider(Qt.Orientation.Horizontal)
        self._slider_quality.setRange(1, 100)
        self._slider_quality.setValue(85)
        self._spin_quality = QSpinBox()
        self._spin_quality.setRange(1, 100)
        self._spin_quality.setValue(85)
        self._slider_quality.valueChanged.connect(self._spin_quality.setValue)
        self._spin_quality.valueChanged.connect(self._slider_quality.setValue)
        quality_row.addWidget(self._slider_quality, 1)
        quality_row.addWidget(self._spin_quality)
        form.addRow("JPEG 质量:", quality_row)

        layout.addWidget(group)

        self._controls = TaskControls()
        self._controls.sig_start.connect(self._on_start)
        self._controls.sig_reset.connect(self._on_reset)
        layout.addWidget(self._controls)

        layout.addStretch()

    def _gather_options(self) -> dict:
        return {
            "remove_images": self._cb_remove.isChecked(),
            "optimize_images": self._cb_optimize.isChecked(),
            "jpeg_quality": self._spin_quality.value(),
        }

    def _on_start(self):
        item = TaskItem(
            task_type="image_extract",
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
