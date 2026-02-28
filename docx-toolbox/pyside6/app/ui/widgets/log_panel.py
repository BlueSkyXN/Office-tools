"""Bottom log panel widget"""

from PySide6.QtWidgets import QWidget, QVBoxLayout, QPlainTextEdit, QHBoxLayout, QPushButton
from PySide6.QtCore import Slot


class LogPanel(QWidget):
    """Displays real-time log output (max 500 lines per CORE-INTERFACE §5)."""

    MAX_LINES = 500

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("log_panel")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Toolbar
        toolbar = QHBoxLayout()
        toolbar.setContentsMargins(8, 4, 8, 4)
        self._lbl = QPushButton("日志")
        self._lbl.setFlat(True)
        self._lbl.setEnabled(False)
        toolbar.addWidget(self._lbl)
        toolbar.addStretch()
        self._btn_clear = QPushButton("清除")
        self._btn_clear.setFixedWidth(60)
        self._btn_clear.clicked.connect(self.clear)
        toolbar.addWidget(self._btn_clear)
        layout.addLayout(toolbar)

        # Text area
        self._text = QPlainTextEdit()
        self._text.setObjectName("log_text")
        self._text.setReadOnly(True)
        self._text.setMaximumBlockCount(self.MAX_LINES)
        layout.addWidget(self._text)

    @Slot(str)
    def append(self, message: str):
        self._text.appendPlainText(message)

    def clear(self):
        self._text.clear()
