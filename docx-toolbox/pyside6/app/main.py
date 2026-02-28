"""Entry point — python3 -m pyside6.app.main"""

import sys
from pathlib import Path

# Ensure docx-toolbox root is importable
_ROOT = str(Path(__file__).resolve().parent.parent.parent)
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from PySide6.QtWidgets import QApplication  # noqa: E402
from PySide6.QtCore import Qt  # noqa: E402

from pyside6.app.config import AppSettings, STYLESHEET  # noqa: E402
from pyside6.app.ui import MainWindow  # noqa: E402


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("DOCX 工具箱")
    app.setOrganizationName("docx-toolbox")
    app.setStyleSheet(STYLESHEET)

    settings = AppSettings()
    window = MainWindow(settings)
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
