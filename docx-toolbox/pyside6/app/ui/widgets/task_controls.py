"""Start / Stop / Reset button group"""

from PySide6.QtWidgets import QWidget, QHBoxLayout, QPushButton
from PySide6.QtCore import Signal


class TaskControls(QWidget):
    """Unified Start/Stop/Reset button row."""

    start_clicked = Signal()
    stop_clicked = Signal()
    reset_clicked = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self.btn_start = QPushButton("开始")
        self.btn_start.setObjectName("btn_start")
        self.btn_start.clicked.connect(self.start_clicked.emit)

        self.btn_stop = QPushButton("停止")
        self.btn_stop.setObjectName("btn_stop")
        self.btn_stop.setEnabled(False)
        self.btn_stop.clicked.connect(self.stop_clicked.emit)

        self.btn_reset = QPushButton("重置")
        self.btn_reset.clicked.connect(self.reset_clicked.emit)

        layout.addStretch()
        layout.addWidget(self.btn_start)
        layout.addWidget(self.btn_stop)
        layout.addWidget(self.btn_reset)

    def set_running(self, running: bool):
        self.btn_start.setEnabled(not running)
        self.btn_stop.setEnabled(running)
        self.btn_reset.setEnabled(not running)
