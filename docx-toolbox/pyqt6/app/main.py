"""入口点 — python3 -m pyqt6.app.main"""

from __future__ import annotations

import sys
from pathlib import Path

# 将 docx-toolbox 根目录加入 sys.path
_TOOLBOX_ROOT = str(Path(__file__).resolve().parent.parent.parent)
if _TOOLBOX_ROOT not in sys.path:
    sys.path.insert(0, _TOOLBOX_ROOT)

from PyQt6.QtWidgets import QApplication  # noqa: E402

from pyqt6.app.ui.main_window import MainWindow  # noqa: E402
from pyqt6.app.config.theme import build_stylesheet  # noqa: E402


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("docx-toolbox")
    app.setStyleSheet(build_stylesheet())

    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
