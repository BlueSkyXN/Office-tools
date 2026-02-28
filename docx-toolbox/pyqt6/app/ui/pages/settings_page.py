"""设置页 — 全局配置与持久化"""

from __future__ import annotations

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QGroupBox, QFormLayout, QLabel,
    QSpinBox, QPushButton, QHBoxLayout,
)
from PyQt6.QtCore import pyqtSignal

from pyqt6.app.config.settings import load_settings, save_settings


class SettingsPage(QWidget):
    """设置页面"""

    sig_settings_changed = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._settings = load_settings()
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 16, 20, 12)

        title = QLabel("设置")
        title.setProperty("heading", True)
        layout.addWidget(title)

        # 运行时参数
        runtime_group = QGroupBox("运行时")
        form = QFormLayout(runtime_group)

        self._spin_workers = QSpinBox()
        self._spin_workers.setRange(1, 8)
        self._spin_workers.setValue(self._settings.get("workers", 1))
        form.addRow("并发数:", self._spin_workers)

        layout.addWidget(runtime_group)

        # 关于
        about_group = QGroupBox("关于")
        about_layout = QVBoxLayout(about_group)
        about_layout.addWidget(QLabel("docx-toolbox (PyQt6)"))
        about_layout.addWidget(QLabel("内部评估版本 — 仅供对比测试"))
        layout.addWidget(about_group)

        # 保存按钮
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_save = QPushButton("保存设置")
        btn_save.clicked.connect(self._save)
        btn_row.addWidget(btn_save)
        layout.addLayout(btn_row)

        layout.addStretch()

    def _save(self):
        self._settings["workers"] = self._spin_workers.value()
        save_settings(self._settings)
        self.sig_settings_changed.emit()

    @property
    def workers(self) -> int:
        return self._spin_workers.value()
