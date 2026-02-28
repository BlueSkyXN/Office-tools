"""日志面板 — 底部实时日志区"""

from __future__ import annotations

from datetime import datetime

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPlainTextEdit, QPushButton, QLabel,
)
from PyQt6.QtCore import pyqtSlot
from PyQt6.QtGui import QTextCharFormat, QColor


_MAX_LINES = 500


class LogPanel(QWidget):
    """底部日志面板，支持按任务过滤"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._setup_ui()
        self._lines: list[str] = []

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 4, 0, 0)

        # 标题栏
        header = QHBoxLayout()
        title = QLabel("日志")
        title.setProperty("heading", True)
        header.addWidget(title)
        header.addStretch()

        self._btn_clear = QPushButton("清空")
        self._btn_clear.setProperty("secondary", True)
        self._btn_clear.setFixedWidth(60)
        self._btn_clear.clicked.connect(self._clear)
        header.addWidget(self._btn_clear)
        layout.addLayout(header)

        # 日志文本区
        self._text = QPlainTextEdit()
        self._text.setReadOnly(True)
        self._text.setMaximumBlockCount(_MAX_LINES)
        layout.addWidget(self._text)

    @pyqtSlot(str, str)
    def append_log(self, task_id: str, message: str):
        ts = datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] [{task_id}] {message}"
        self._text.appendPlainText(line)

    def append_system(self, message: str):
        ts = datetime.now().strftime("%H:%M:%S")
        self._text.appendPlainText(f"[{ts}] [SYS] {message}")

    def _clear(self):
        self._text.clear()
