"""批量任务页 — 多文件队列处理"""

from __future__ import annotations

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QGroupBox, QComboBox, QTableWidget, QTableWidgetItem,
    QHeaderView, QAbstractItemView,
)
from PyQt6.QtCore import pyqtSignal

from pyqt6.app.ui.widgets.file_picker import FilePicker
from pyqt6.app.ui.widgets.task_controls import TaskControls
from pyqt6.app.models.task_model import TaskItem, TaskStatus


_TASK_TYPES = [
    ("excel_allinone", "Excel 处理"),
    ("image_extract", "图片分离"),
    ("table_extract", "表格提取"),
]


class BatchPage(QWidget):
    """批量任务队列"""

    sig_run_task = pyqtSignal(object)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._queue: list[TaskItem] = []
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 16, 20, 12)

        title = QLabel("批量任务")
        title.setProperty("heading", True)
        layout.addWidget(title)

        # 添加任务区
        add_group = QGroupBox("添加任务")
        add_layout = QVBoxLayout(add_group)

        type_row = QHBoxLayout()
        type_row.addWidget(QLabel("任务类型:"))
        self._combo_type = QComboBox()
        for val, label in _TASK_TYPES:
            self._combo_type.addItem(label, val)
        type_row.addWidget(self._combo_type, 1)
        add_layout.addLayout(type_row)

        self._input = FilePicker("选择目录", mode="dir")
        add_layout.addWidget(self._input)
        self._output = FilePicker("输出目录（可选）", mode="dir")
        add_layout.addWidget(self._output)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_add = QPushButton("添加到队列")
        btn_add.clicked.connect(self._add_to_queue)
        btn_row.addWidget(btn_add)
        add_layout.addLayout(btn_row)

        layout.addWidget(add_group)

        # 队列表格
        self._table = QTableWidget(0, 4)
        self._table.setHorizontalHeaderLabels(["任务ID", "类型", "路径", "状态"])
        self._table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self._table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        layout.addWidget(self._table)

        # 控制
        ctrl_row = QHBoxLayout()
        ctrl_row.addStretch()

        btn_clear = QPushButton("清空队列")
        btn_clear.setProperty("secondary", True)
        btn_clear.clicked.connect(self._clear_queue)
        ctrl_row.addWidget(btn_clear)

        btn_retry = QPushButton("重试失败")
        btn_retry.setProperty("secondary", True)
        btn_retry.clicked.connect(self._retry_failed)
        ctrl_row.addWidget(btn_retry)

        btn_run = QPushButton("全部执行")
        btn_run.clicked.connect(self._run_all)
        ctrl_row.addWidget(btn_run)

        layout.addLayout(ctrl_row)

    def _add_to_queue(self):
        task_type = self._combo_type.currentData()
        input_path = self._input.path()
        if not input_path:
            return
        item = TaskItem(
            task_type=task_type,
            input_path=input_path,
            output_dir=self._output.path(),
        )
        self._queue.append(item)
        self._refresh_table()

    def _refresh_table(self):
        self._table.setRowCount(len(self._queue))
        for i, item in enumerate(self._queue):
            self._table.setItem(i, 0, QTableWidgetItem(item.task_id))
            self._table.setItem(i, 1, QTableWidgetItem(item.task_type))
            self._table.setItem(i, 2, QTableWidgetItem(item.input_path))
            self._table.setItem(i, 3, QTableWidgetItem(item.status.value))

    def _clear_queue(self):
        self._queue.clear()
        self._refresh_table()

    def _retry_failed(self):
        for item in self._queue:
            if item.status == TaskStatus.FAILED:
                item.status = TaskStatus.PENDING
        self._run_all()

    def _run_all(self):
        for item in self._queue:
            if item.status == TaskStatus.PENDING:
                self.sig_run_task.emit(item)

    def update_task_status(self, task_id: str, status: TaskStatus):
        for item in self._queue:
            if item.task_id == task_id:
                item.status = status
                break
        self._refresh_table()
