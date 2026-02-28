"""文件/目录选择控件"""

from __future__ import annotations

from PyQt6.QtWidgets import (
    QWidget, QHBoxLayout, QLineEdit, QPushButton, QFileDialog,
)
from PyQt6.QtCore import pyqtSignal


class FilePicker(QWidget):
    """带浏览按钮的路径选择器"""

    path_changed = pyqtSignal(str)

    def __init__(self, label: str = "选择文件", mode: str = "file", parent=None):
        super().__init__(parent)
        self._mode = mode  # "file" | "dir"
        self._setup_ui(label)

    def _setup_ui(self, label: str):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self._edit = QLineEdit()
        self._edit.setPlaceholderText(label)
        self._edit.textChanged.connect(self.path_changed.emit)
        layout.addWidget(self._edit, 1)

        btn = QPushButton("浏览...")
        btn.setProperty("secondary", True)
        btn.setFixedWidth(72)
        btn.clicked.connect(self._browse)
        layout.addWidget(btn)

    def _browse(self):
        if self._mode == "dir":
            path = QFileDialog.getExistingDirectory(self, "选择目录")
        else:
            path, _ = QFileDialog.getOpenFileName(
                self, "选择文件", "", "Word 文档 (*.docx);;所有文件 (*)"
            )
        if path:
            self._edit.setText(path)

    def path(self) -> str:
        return self._edit.text().strip()

    def set_path(self, p: str):
        self._edit.setText(p)
