"""Reusable file / folder picker widget"""

from PySide6.QtWidgets import QWidget, QHBoxLayout, QLineEdit, QPushButton, QFileDialog
from PySide6.QtCore import Signal


class FilePicker(QWidget):
    """A QLineEdit + Browse button for file or folder selection."""

    path_changed = Signal(str)

    def __init__(
        self,
        label: str = "选择路径",
        mode: str = "file_or_dir",  # "file" | "dir" | "file_or_dir"
        file_filter: str = "DOCX 文件 (*.docx);;所有文件 (*)",
        parent=None,
    ):
        super().__init__(parent)
        self._mode = mode
        self._file_filter = file_filter

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self._edit = QLineEdit()
        self._edit.setPlaceholderText(label)
        self._edit.textChanged.connect(self.path_changed.emit)
        layout.addWidget(self._edit, 1)

        self._btn = QPushButton("浏览…")
        self._btn.setFixedWidth(72)
        self._btn.clicked.connect(self._browse)
        layout.addWidget(self._btn)

    def path(self) -> str:
        return self._edit.text().strip()

    def set_path(self, p: str):
        self._edit.setText(p)

    def _browse(self):
        if self._mode == "dir":
            p = QFileDialog.getExistingDirectory(self, "选择目录")
        elif self._mode == "file":
            p, _ = QFileDialog.getOpenFileName(self, "选择文件", "", self._file_filter)
        else:
            # file_or_dir: let user pick file first, fallback to dir
            p, _ = QFileDialog.getOpenFileName(self, "选择文件或目录", "", self._file_filter)
            if not p:
                p = QFileDialog.getExistingDirectory(self, "选择目录")
        if p:
            self._edit.setText(p)
