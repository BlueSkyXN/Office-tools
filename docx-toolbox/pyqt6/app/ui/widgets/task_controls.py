"""任务控制按钮组 — 开始 / 停止 / 重置"""

from __future__ import annotations

from PyQt6.QtWidgets import QWidget, QHBoxLayout, QPushButton, QProgressBar
from PyQt6.QtCore import pyqtSignal


class TaskControls(QWidget):
    """开始 / 停止 / 重置 按钮组 + 进度条"""

    sig_start = pyqtSignal()
    sig_stop = pyqtSignal()
    sig_reset = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._setup_ui()

    def _setup_ui(self):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 6, 0, 0)

        self._progress = QProgressBar()
        self._progress.setValue(0)
        self._progress.setTextVisible(True)
        layout.addWidget(self._progress, 1)

        self._btn_start = QPushButton("开始")
        self._btn_start.clicked.connect(self.sig_start.emit)
        layout.addWidget(self._btn_start)

        self._btn_stop = QPushButton("停止")
        self._btn_stop.setProperty("secondary", True)
        self._btn_stop.setEnabled(False)
        self._btn_stop.clicked.connect(self.sig_stop.emit)
        layout.addWidget(self._btn_stop)

        self._btn_reset = QPushButton("重置")
        self._btn_reset.setProperty("secondary", True)
        self._btn_reset.clicked.connect(self.sig_reset.emit)
        layout.addWidget(self._btn_reset)

    def set_running(self, running: bool):
        self._btn_start.setEnabled(not running)
        self._btn_stop.setEnabled(running)

    def set_progress(self, value: int):
        self._progress.setValue(value)

    def reset(self):
        self._progress.setValue(0)
        self.set_running(False)
