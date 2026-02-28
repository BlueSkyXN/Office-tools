"""批量任务页"""

from pathlib import Path

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QGroupBox, QFormLayout, QComboBox, QLabel,
    QTableWidget, QTableWidgetItem, QProgressBar, QHBoxLayout,
)
from PySide6.QtCore import Signal, Qt

from pyside6.app.ui.widgets import FilePicker, TaskControls
from pyside6.app.core.adapter import build_task_request
from pyside6.app.runner.worker import BatchWorker

_TASK_TYPES = {
    "Excel 嵌入对象处理": "excel_allinone",
    "DOCX 图片分离": "image_extract",
    "DOCX 表格提取": "table_extract",
}


class BatchPage(QWidget):
    log_message = Signal(str)
    status_message = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._worker: BatchWorker | None = None
        self._setup_ui()
        self._connect_signals()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 12)

        title = QLabel("批量任务")
        title.setStyleSheet("font-size: 18px; font-weight: 700; margin-bottom: 4px;")
        layout.addWidget(title)

        # Input
        input_group = QGroupBox("批量设置")
        ig_layout = QFormLayout(input_group)
        self._folder_picker = FilePicker("选择包含 .docx 文件的目录", mode="dir")
        ig_layout.addRow("输入目录:", self._folder_picker)
        self._combo_type = QComboBox()
        self._combo_type.addItems(_TASK_TYPES.keys())
        ig_layout.addRow("任务类型:", self._combo_type)
        layout.addWidget(input_group)

        # Queue table
        queue_group = QGroupBox("任务队列")
        qg_layout = QVBoxLayout(queue_group)
        self._table = QTableWidget(0, 3)
        self._table.setHorizontalHeaderLabels(["文件名", "任务类型", "状态"])
        self._table.horizontalHeader().setStretchLastSection(True)
        self._table.setEditTriggers(QTableWidget.NoEditTriggers)
        self._table.setSelectionBehavior(QTableWidget.SelectRows)
        qg_layout.addWidget(self._table)

        self._progress = QProgressBar()
        self._progress.setValue(0)
        qg_layout.addWidget(self._progress)
        layout.addWidget(queue_group, 1)

        # Controls
        ctrl_row = QHBoxLayout()
        self._btn_scan = TaskControls.__new__(TaskControls)  # We'll make our own
        # Actually use a simple button + TaskControls
        from PySide6.QtWidgets import QPushButton
        self._btn_scan_files = QPushButton("扫描文件")
        self._btn_scan_files.clicked.connect(self._on_scan)
        ctrl_row.addWidget(self._btn_scan_files)
        ctrl_row.addStretch()
        layout.addLayout(ctrl_row)

        self._controls = TaskControls()
        layout.addWidget(self._controls)

    def _connect_signals(self):
        self._controls.start_clicked.connect(self._on_start)
        self._controls.stop_clicked.connect(self._on_stop)
        self._controls.reset_clicked.connect(self._on_reset)

    def _on_scan(self):
        folder = self._folder_picker.path()
        if not folder or not Path(folder).is_dir():
            self.log_message.emit("错误: 请选择有效的目录")
            return

        docx_files = sorted(
            p for p in Path(folder).iterdir()
            if p.suffix.lower() == ".docx" and not p.name.startswith("~$")
        )
        self._table.setRowCount(0)
        task_label = self._combo_type.currentText()
        for f in docx_files:
            row = self._table.rowCount()
            self._table.insertRow(row)
            self._table.setItem(row, 0, QTableWidgetItem(f.name))
            self._table.setItem(row, 1, QTableWidgetItem(task_label))
            item = QTableWidgetItem("pending")
            item.setTextAlignment(Qt.AlignCenter)
            self._table.setItem(row, 2, item)

        self.log_message.emit(f"扫描到 {len(docx_files)} 个 .docx 文件")
        self._progress.setMaximum(len(docx_files))
        self._progress.setValue(0)

    def _on_start(self):
        if self._table.rowCount() == 0:
            self.log_message.emit("错误: 队列为空，请先扫描文件")
            return

        folder = self._folder_picker.path()
        task_type = _TASK_TYPES[self._combo_type.currentText()]
        requests = []
        for row in range(self._table.rowCount()):
            fname = self._table.item(row, 0).text()
            fpath = str(Path(folder) / fname)
            requests.append(build_task_request(task_type=task_type, input_path=fpath))

        self._worker = BatchWorker(requests)
        self._worker.job_updated.connect(self._on_job_updated)
        self._worker.progress.connect(self._on_progress)
        self._worker.all_finished.connect(self._on_all_finished)
        self._worker.log_message.connect(self.log_message.emit)
        self._controls.set_running(True)
        self.status_message.emit("运行中: 批量处理…")
        self._worker.start()

    def _on_stop(self):
        if self._worker:
            self._worker.cancel()
            self._worker.quit()
            self._worker.wait(3000)
            self._controls.set_running(False)
            self.status_message.emit("已停止")
            self.log_message.emit("批量任务已停止")

    def _on_reset(self):
        self._table.setRowCount(0)
        self._progress.setValue(0)
        self.status_message.emit("就绪")

    def _on_job_updated(self, idx: int, status: str):
        if 0 <= idx < self._table.rowCount():
            item = self._table.item(idx, 2)
            if item:
                item.setText(status)

    def _on_progress(self, current: int, total: int):
        self._progress.setMaximum(total)
        self._progress.setValue(current)

    def _on_all_finished(self, jobs):
        self._controls.set_running(False)
        success = sum(1 for j in jobs if j.status.value == "success")
        failed = sum(1 for j in jobs if j.status.value == "failed")
        self.status_message.emit(f"完成: 成功 {success}, 失败 {failed}")
