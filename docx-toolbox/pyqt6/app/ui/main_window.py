"""主窗口 — 左侧导航 + 右侧页面栈 + 底部日志"""

from __future__ import annotations

from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QHBoxLayout, QVBoxLayout,
    QListWidget, QStackedWidget, QSplitter,
)
from PyQt6.QtCore import Qt, pyqtSlot

from pyqt6.app.ui.pages.excel_page import ExcelPage
from pyqt6.app.ui.pages.image_page import ImagePage
from pyqt6.app.ui.pages.table_page import TablePage
from pyqt6.app.ui.pages.batch_page import BatchPage
from pyqt6.app.ui.pages.settings_page import SettingsPage
from pyqt6.app.ui.widgets.log_panel import LogPanel
from pyqt6.app.models.task_model import TaskItem, TaskStatus
from pyqt6.app.runner.worker import TaskWorker


class MainWindow(QMainWindow):
    """docx-toolbox 主窗口"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("docx-toolbox (PyQt6)")
        self.resize(960, 640)
        self._workers: list[TaskWorker] = []
        self._setup_ui()
        self._connect_signals()

    # ------------------------------------------------------------------ UI
    def _setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        splitter = QSplitter(Qt.Orientation.Vertical)

        # 上半区：导航 + 内容
        upper = QWidget()
        upper_layout = QHBoxLayout(upper)
        upper_layout.setContentsMargins(0, 0, 0, 0)
        upper_layout.setSpacing(0)

        # 左侧导航
        self._nav = QListWidget()
        self._nav.setFixedWidth(160)
        for label in ["Excel 处理", "图片分离", "表格提取", "批量任务", "设置"]:
            self._nav.addItem(label)
        self._nav.setCurrentRow(0)
        upper_layout.addWidget(self._nav)

        # 右侧页面栈
        self._stack = QStackedWidget()
        self._excel_page = ExcelPage()
        self._image_page = ImagePage()
        self._table_page = TablePage()
        self._batch_page = BatchPage()
        self._settings_page = SettingsPage()

        self._stack.addWidget(self._excel_page)
        self._stack.addWidget(self._image_page)
        self._stack.addWidget(self._table_page)
        self._stack.addWidget(self._batch_page)
        self._stack.addWidget(self._settings_page)
        upper_layout.addWidget(self._stack, 1)

        splitter.addWidget(upper)

        # 底部日志
        self._log = LogPanel()
        splitter.addWidget(self._log)

        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 1)

        root.addWidget(splitter)

    # -------------------------------------------------------------- 信号连接
    def _connect_signals(self):
        self._nav.currentRowChanged.connect(self._stack.setCurrentIndex)

        # 各页面任务信号
        self._excel_page.sig_run_task.connect(self._run_task)
        self._image_page.sig_run_task.connect(self._run_task)
        self._table_page.sig_run_task.connect(self._run_task)
        self._batch_page.sig_run_task.connect(self._run_task)

        # 停止信号
        self._excel_page.controls.sig_stop.connect(self._stop_current)
        self._image_page.controls.sig_stop.connect(self._stop_current)
        self._table_page.controls.sig_stop.connect(self._stop_current)

    # -------------------------------------------------------------- 任务执行
    @pyqtSlot(object)
    def _run_task(self, item: TaskItem):
        worker = TaskWorker(item)
        worker.task_log.connect(self._log.append_log)
        worker.task_started.connect(self._on_task_started)
        worker.task_finished.connect(self._on_task_finished)
        worker.finished.connect(lambda: self._cleanup_worker(worker))
        self._workers.append(worker)
        worker.start()

    def _on_task_started(self, task_id: str):
        self._log.append_system(f"任务 {task_id} 已启动")
        self._batch_page.update_task_status(task_id, TaskStatus.RUNNING)

    @pyqtSlot(str, bool, str)
    def _on_task_finished(self, task_id: str, ok: bool, msg: str):
        status = TaskStatus.SUCCESS if ok else TaskStatus.FAILED
        self._log.append_system(f"任务 {task_id} {'成功' if ok else '失败'}: {msg}")
        self._batch_page.update_task_status(task_id, status)

        # 恢复当前页面控制按钮
        page = self._stack.currentWidget()
        if hasattr(page, "controls"):
            page.controls.set_running(False)
            page.controls.set_progress(100 if ok else 0)

    def _stop_current(self):
        for w in self._workers:
            if w.isRunning():
                w.cancel()

    def _cleanup_worker(self, worker: TaskWorker):
        if worker in self._workers:
            self._workers.remove(worker)
        worker.deleteLater()
