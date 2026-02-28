"""QSettings-based config persistence"""

from PySide6.QtCore import QSettings


class AppSettings:
    """Wraps QSettings for app preferences."""

    _ORG = "docx-toolbox"
    _APP = "pyside6"

    def __init__(self):
        self._qs = QSettings(self._ORG, self._APP)

    # ---- default output dir ----
    @property
    def default_output_dir(self) -> str:
        return self._qs.value("default_output_dir", "", type=str)

    @default_output_dir.setter
    def default_output_dir(self, v: str):
        self._qs.setValue("default_output_dir", v)

    # ---- worker count ----
    @property
    def worker_count(self) -> int:
        return self._qs.value("worker_count", 1, type=int)

    @worker_count.setter
    def worker_count(self, v: int):
        self._qs.setValue("worker_count", max(1, v))

    # ---- last input path ----
    @property
    def last_input_path(self) -> str:
        return self._qs.value("last_input_path", "", type=str)

    @last_input_path.setter
    def last_input_path(self, v: str):
        self._qs.setValue("last_input_path", v)

    # ---- window geometry ----
    def save_geometry(self, geo: bytes):
        self._qs.setValue("window_geometry", geo)

    def load_geometry(self) -> bytes | None:
        return self._qs.value("window_geometry")
