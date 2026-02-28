"""表格提取适配器 — 封装 references/DOCX表格提取.py"""

from __future__ import annotations

import importlib.util
import os
from pathlib import Path

from core.adapters import BaseAdapter
from core.api import TaskRequest, TaskSummary
from core.errors import CancelledError, ProcessFailedError
from core.logging_utils import get_logger

logger = get_logger()

# ---------------------------------------------------------------------------
# 动态加载参考脚本
# ---------------------------------------------------------------------------
_REF_DIR = str(Path(__file__).resolve().parent.parent.parent / "references")
_spec = importlib.util.spec_from_file_location(
    "docx_table_extract",
    os.path.join(_REF_DIR, "DOCX表格提取.py"),
)
_ref = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_ref)

TABLE_MARK_SUFFIX: str = _ref.TABLE_MARK_SUFFIX  # "_已标记表格"


class TableExtractAdapter(BaseAdapter):
    """从 DOCX 文档提取表格并生成标记文档 + 导出文件"""

    # ------------------------------------------------------------------ #
    #  public API
    # ------------------------------------------------------------------ #

    def execute(self, request: TaskRequest) -> TaskSummary:
        include_marked: bool = request.options.get("include_marked", False)
        cancel_event = request.runtime.cancel_event

        input_path = self.validate_input_path(request.input_path)
        output_dir = self.ensure_output_dir(request.output_dir, input_path if input_path.is_dir() else input_path.parent)

        if input_path.is_file():
            self.validate_docx(input_path)
            self.ensure_not_cancelled(cancel_event, detail="cancelled before single file task")
            return self._process_single(input_path, output_dir)

        # 目录模式
        return self._process_directory(
            input_path,
            output_dir,
            include_marked=include_marked,
            cancel_event=cancel_event,
        )

    # ------------------------------------------------------------------ #
    #  单文件处理
    # ------------------------------------------------------------------ #

    def _process_single(self, file_path: Path, output_dir: Path) -> TaskSummary:
        summary = TaskSummary()
        self._run_one(file_path, output_dir, summary)
        return summary

    # ------------------------------------------------------------------ #
    #  目录处理
    # ------------------------------------------------------------------ #

    def _process_directory(
        self,
        directory: Path,
        output_dir: Path,
        *,
        include_marked: bool = False,
        cancel_event=None,
    ) -> TaskSummary:
        summary = TaskSummary()

        docx_files = self._collect_files(directory, include_marked=include_marked)
        if not docx_files:
            logger.info("目录中未找到可处理的 .docx 文件: %s", directory)
            return summary

        logger.info("待处理 DOCX: %d 个", len(docx_files))
        for file_path in docx_files:
            self.ensure_not_cancelled(
                cancel_event,
                detail=f"cancelled before processing {file_path.name}",
            )
            self._run_one(file_path, output_dir, summary)

        return summary

    # ------------------------------------------------------------------ #
    #  辅助：收集可处理文件（复用参考脚本的过滤逻辑）
    # ------------------------------------------------------------------ #

    @staticmethod
    def _collect_files(directory: Path, *, include_marked: bool = False) -> list[Path]:
        files: list[Path] = []
        for item in sorted(directory.iterdir(), key=lambda p: p.name.lower()):
            if item.is_dir():
                continue
            if item.suffix.lower() != ".docx":
                continue
            if item.name.startswith("~$"):
                continue
            if not include_marked and TABLE_MARK_SUFFIX in item.stem:
                logger.debug("跳过已标记表格文件: %s", item.name)
                continue
            files.append(item)
        return files

    # ------------------------------------------------------------------ #
    #  辅助：单文件调用参考脚本并记录结果
    # ------------------------------------------------------------------ #

    def _run_one(self, file_path: Path, output_dir: Path, summary: TaskSummary) -> None:
        logger.info("处理文件: %s", file_path.name)
        try:
            result = _ref.process_docx(str(file_path))
        except CancelledError:
            raise
        except Exception as exc:
            logger.error("处理失败 %s: %s", file_path.name, exc)
            summary.failed += 1
            return

        if result is None:
            # 文档中没有表格
            logger.info("跳过（无表格）: %s", file_path.name)
            summary.skipped += 1
            return

        if result is not True:
            summary.failed += 1
            return

        # 成功 — 收集输出文件路径
        summary.processed += 1
        stem = file_path.stem
        parent = file_path.parent

        candidates = [
            parent / f"{stem}{TABLE_MARK_SUFFIX}.docx",
            parent / f"{stem}_表格提取.txt",
            parent / f"{stem}_表格提取.xlsx",
            parent / f"{stem}_表格提取.pdf",
        ]
        for p in candidates:
            if p.exists():
                summary.outputs.append(str(p))
