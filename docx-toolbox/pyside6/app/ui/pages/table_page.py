"""表格提取页 — task_type: table_extract"""

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QGroupBox, QFormLayout, QCheckBox, QLabel,
)
from PySide6.QtCore import Signal

from pyside6.app.ui.widgets import FilePicker, TaskControls
from pyside6.app.core.adapter import build_task_request
from pyside6.app.runner.worker import TaskWorker


class TablePage(QWidget):
    log_message = Signal(str)
    status_message = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._worker: TaskWorker | None = None
        self._setup_ui()
        self._connect_signals()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 12)

        title = QLabel("DOCX 表格提取")
        title.setStyleSheet("font-size: 18px; font-weight: 700; margin-bottom: 4px;")
        layout.addWidget(title)

        # Input
        input_group = QGroupBox("输入")
        ig_layout = QFormLayout(input_group)
        self._input_picker = FilePicker("选择 .docx 文件或目录")
        ig_layout.addRow("输入路径:", self._input_picker)
        self._output_picker = FilePicker("输出目录（可选）", mode="dir")
        ig_layout.addRow("输出目录:", self._output_picker)
        layout.addWidget(input_group)

        # Options (CORE-INTERFACE §2.1)
        opts_group = QGroupBox("选项")
        og_layout = QVBoxLayout(opts_group)
        self._cb_include_marked = QCheckBox("包含已标记表格文件 (include_marked)")
        og_layout.addWidget(self._cb_include_marked)
        layout.addWidget(opts_group)

        # Controls
        self._controls = TaskControls()
        layout.addWidget(self._controls)
        layout.addStretch()

    def _connect_signals(self):
        self._controls.start_clicked.connect(self._on_start)
        self._controls.stop_clicked.connect(self._on_stop)
        self._controls.reset_clicked.connect(self._on_reset)

    def _gather_options(self) -> dict:
        return {
            "include_marked": self._cb_include_marked.isChecked(),
        }

    def _on_start(self):
        input_path = self._input_picker.path()
        if not input_path:
            self.log_message.emit("错误: 请选择输入路径")
            return

        request = build_task_request(
            task_type="table_extract",
            input_path=input_path,
            output_dir=self._output_picker.path(),
            options=self._gather_options(),
        )
        self._worker = TaskWorker(request)
        self._worker.log_message.connect(self.log_message.emit)
        self._worker.finished.connect(self._on_finished)
        self._worker.error.connect(self._on_error)
        self._controls.set_running(True)
        self.status_message.emit("运行中: 表格提取…")
        self._worker.start()

    def _on_stop(self):
        if self._worker:
            self._worker.cancel()
            self._worker.quit()
            self._worker.wait(3000)
            self._controls.set_running(False)
            self.status_message.emit("已停止")
            self.log_message.emit("任务已停止")

    def _on_reset(self):
        self._input_picker.set_path("")
        self._output_picker.set_path("")
        self._cb_include_marked.setChecked(False)
        self.status_message.emit("就绪")

    def _on_finished(self, response):
        self._controls.set_running(False)
        if response.ok:
            self.status_message.emit("完成 ✓")
        else:
            msg = response.error.get("message", "未知错误") if response.error else "未知错误"
            self.status_message.emit(f"失败: {msg}")

    def _on_error(self, msg: str):
        self._controls.set_running(False)
        self.status_message.emit(f"错误: {msg}")
