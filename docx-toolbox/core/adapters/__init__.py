"""适配器基类与公共工具"""

from __future__ import annotations

import os
from abc import ABC, abstractmethod
from pathlib import Path
import threading

from core.api import TaskRequest, TaskSummary
from core.errors import (
    CancelledError,
    InvalidInputError,
    PermissionDeniedError,
    UnsupportedFormatError,
)


class BaseAdapter(ABC):
    """所有任务适配器的基类"""

    @abstractmethod
    def execute(self, request: TaskRequest) -> TaskSummary:
        ...

    # ---- 公共校验 ----

    @staticmethod
    def validate_input_path(path_str: str, must_be_file: bool = False) -> Path:
        p = Path(path_str)
        if not p.exists():
            raise InvalidInputError(f"路径不存在: {path_str}")
        if must_be_file and p.is_dir():
            raise InvalidInputError(f"预期文件，但输入是目录: {path_str}")
        return p

    @staticmethod
    def validate_docx(path: Path) -> None:
        if path.suffix.lower() != ".docx":
            raise UnsupportedFormatError(f"不是 .docx 文件: {path.name}")

    @staticmethod
    def ensure_output_dir(output_dir: str | None, fallback: Path) -> Path:
        if output_dir:
            d = Path(output_dir)
        else:
            d = fallback
        d.mkdir(parents=True, exist_ok=True)
        if not os.access(d, os.W_OK):
            raise PermissionDeniedError(f"输出目录无写权限: {d}")
        return d

    @staticmethod
    def collect_docx_files(directory: Path) -> list[Path]:
        """收集目录下的 .docx 文件（不递归，跳过临时文件）"""
        files = []
        for item in sorted(directory.iterdir(), key=lambda p: p.name.lower()):
            if item.is_dir():
                continue
            if item.suffix.lower() != ".docx":
                continue
            if item.name.startswith("~$"):
                continue
            files.append(item)
        return files

    @staticmethod
    def ensure_not_cancelled(
        cancel_event: threading.Event | None,
        *,
        detail: str = "",
    ) -> None:
        """协作式取消检查。"""
        if cancel_event is not None and cancel_event.is_set():
            raise CancelledError(detail=detail)
