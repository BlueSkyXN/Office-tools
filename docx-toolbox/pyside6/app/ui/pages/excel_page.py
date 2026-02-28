"""Excel 嵌入对象处理页 — task_type: excel_allinone"""

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QGroupBox, QFormLayout, QCheckBox, QLabel,
)
from PySide6.QtCore import Signal

from pyside6.app.ui.widgets import FilePicker, TaskControls
from pyside6.app.core.adapter import build_task_request, execute_task
from pyside6.app.runner.worker import TaskWorker


class ExcelPage(QWidget):
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

        title = QLabel("Excel 嵌入对象处理")
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
        self._cb_word_table = QCheckBox("转换为 Word 原生表格 (word_table)")
        self._cb_word_table.setChecked(True)
        self._cb_extract_excel = QCheckBox("提取嵌入 Excel 文件 (extract_excel)")
        self._cb_image = QCheckBox("渲染为图片 (image)")
        self._cb_keep_attachment = QCheckBox("保留附件入口 (keep_attachment)")
        self._cb_remove_watermark = QCheckBox("移除水印 (remove_watermark)")
        self._cb_a3 = QCheckBox("设置 A3 横向 (a3)")
        for cb in (self._cb_word_table, self._cb_extract_excel, self._cb_image,
                   self._cb_keep_attachment, self._cb_remove_watermark, self._cb_a3):
            og_layout.addWidget(cb)
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
            "word_table": self._cb_word_table.isChecked(),
            "extract_excel": self._cb_extract_excel.isChecked(),
            "image": self._cb_image.isChecked(),
            "keep_attachment": self._cb_keep_attachment.isChecked(),
            "remove_watermark": self._cb_remove_watermark.isChecked(),
            "a3": self._cb_a3.isChecked(),
        }

    def _on_start(self):
        input_path = self._input_picker.path()
        if not input_path:
            self.log_message.emit("错误: 请选择输入路径")
            return

        request = build_task_request(
            task_type="excel_allinone",
            input_path=input_path,
            output_dir=self._output_picker.path(),
            options=self._gather_options(),
        )
        self._worker = TaskWorker(request)
        self._worker.log_message.connect(self.log_message.emit)
        self._worker.finished.connect(self._on_finished)
        self._worker.error.connect(self._on_error)
        self._controls.set_running(True)
        self.status_message.emit("运行中: Excel 处理…")
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
        self._cb_word_table.setChecked(True)
        for cb in (self._cb_extract_excel, self._cb_image,
                   self._cb_keep_attachment, self._cb_remove_watermark, self._cb_a3):
            cb.setChecked(False)
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
