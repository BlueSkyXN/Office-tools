"""Main window: left nav + right content (QStackedWidget) + bottom log panel"""

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QHBoxLayout, QVBoxLayout,
    QPushButton, QStackedWidget, QSplitter, QStatusBar,
)
from PySide6.QtCore import Qt

from pyside6.app.ui.pages import ExcelPage, ImagePage, TablePage, BatchPage, SettingsPage
from pyside6.app.ui.widgets import LogPanel
from pyside6.app.config.settings import AppSettings


_NAV_ITEMS = [
    ("Excel处理", "excel"),
    ("图片分离", "image"),
    ("表格提取", "table"),
    ("批量任务", "batch"),
    ("设置",     "settings"),
]


class MainWindow(QMainWindow):
    def __init__(self, settings: AppSettings, parent=None):
        super().__init__(parent)
        self._settings = settings
        self.setWindowTitle("DOCX 工具箱")
        self.setMinimumSize(960, 640)
        self._nav_buttons: list[QPushButton] = []
        self._setup_ui()
        self._restore_geometry()

    def _setup_ui(self):
        # Central widget
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # Splitter: top (nav + content) | bottom (log)
        splitter = QSplitter(Qt.Vertical)

        # Top section
        top = QWidget()
        top_layout = QHBoxLayout(top)
        top_layout.setContentsMargins(0, 0, 0, 0)
        top_layout.setSpacing(0)

        # Left nav panel
        nav = QWidget()
        nav.setObjectName("nav_panel")
        nav.setFixedWidth(160)
        nav_layout = QVBoxLayout(nav)
        nav_layout.setContentsMargins(0, 12, 0, 12)
        nav_layout.setSpacing(2)

        for label, _key in _NAV_ITEMS:
            btn = QPushButton(label)
            btn.setCheckable(True)
            btn.setCursor(Qt.PointingHandCursor)
            btn.clicked.connect(lambda checked, k=_key: self._switch_page(k))
            nav_layout.addWidget(btn)
            self._nav_buttons.append(btn)
        nav_layout.addStretch()
        top_layout.addWidget(nav)

        # Right stacked widget
        self._stack = QStackedWidget()
        self._pages: dict[str, QWidget] = {}

        self._pages["excel"] = ExcelPage()
        self._pages["image"] = ImagePage()
        self._pages["table"] = TablePage()
        self._pages["batch"] = BatchPage()
        self._pages["settings"] = SettingsPage(self._settings)

        for key in ("excel", "image", "table", "batch", "settings"):
            self._stack.addWidget(self._pages[key])

        top_layout.addWidget(self._stack, 1)
        splitter.addWidget(top)

        # Bottom log panel
        self._log_panel = LogPanel()
        splitter.addWidget(self._log_panel)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 1)

        root.addWidget(splitter)

        # Status bar
        self._status_bar = QStatusBar()
        self.setStatusBar(self._status_bar)
        self._status_bar.showMessage("就绪")

        # Wire log / status signals from pages
        for page in self._pages.values():
            if hasattr(page, "log_message"):
                page.log_message.connect(self._log_panel.append)
            if hasattr(page, "status_message"):
                page.status_message.connect(self._status_bar.showMessage)

        # Default selection
        self._switch_page("excel")

    def _switch_page(self, key: str):
        keys = [k for _, k in _NAV_ITEMS]
        idx = keys.index(key) if key in keys else 0
        self._stack.setCurrentIndex(idx)
        for i, btn in enumerate(self._nav_buttons):
            btn.setChecked(i == idx)

    def _restore_geometry(self):
        geo = self._settings.load_geometry()
        if geo:
            self.restoreGeometry(geo)

    def closeEvent(self, event):
        self._settings.save_geometry(self.saveGeometry())
        super().closeEvent(event)
