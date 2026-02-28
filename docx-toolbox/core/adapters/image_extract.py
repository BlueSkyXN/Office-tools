"""图片分离适配器 — 将 DOCX 中的图片提取为 PDF 并标记原文档"""

from __future__ import annotations

import importlib.util
import os
from pathlib import Path

from core.adapters import BaseAdapter
from core.api import TaskRequest, TaskSummary
from core.errors import CancelledError, InvalidInputError, ProcessFailedError
from core.logging_utils import get_logger

logger = get_logger()

# ---------------------------------------------------------------------------
# 动态加载参考脚本（文件名含中文，无法 import）
# ---------------------------------------------------------------------------
_REF_DIR = str(Path(__file__).resolve().parent.parent.parent / "references")
_spec = importlib.util.spec_from_file_location(
    "docx_image_extract",
    os.path.join(_REF_DIR, "DOCX图片分离.py"),
)
_ref = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_ref)


class ImageExtractAdapter(BaseAdapter):
    """DOCX 图片分离：提取图片生成 PDF，原文标记位置"""

    def execute(self, request: TaskRequest) -> TaskSummary:
        opts = request.options
        cancel_event = request.runtime.cancel_event
        remove_images: bool = opts.get("remove_images", False)
        optimize_images: bool = opts.get("optimize_images", True)
        jpeg_quality: int = opts.get("jpeg_quality", 85)

        input_path = self.validate_input_path(request.input_path)

        if input_path.is_file():
            self.validate_docx(input_path)
            files = [input_path]
        else:
            files = self.collect_docx_files(input_path)
            if not files:
                raise InvalidInputError(f"目录中未找到 .docx 文件: {input_path}")

        fallback = input_path if input_path.is_dir() else input_path.parent
        output_dir = self.ensure_output_dir(request.output_dir, fallback)

        summary = TaskSummary()

        for file_path in files:
            self.ensure_not_cancelled(
                cancel_event,
                detail=f"cancelled before processing {file_path.name}",
            )
            # 跳过已处理的标记文件
            if "_已标记图片" in file_path.stem:
                logger.info("跳过已标记文件: %s", file_path.name)
                summary.skipped += 1
                continue

            try:
                ok: bool = _ref.process_docx_file(
                    str(file_path),
                    remove_images=remove_images,
                    output_dir=str(output_dir),
                    optimize_images=optimize_images,
                    jpeg_quality=jpeg_quality,
                )
                self.ensure_not_cancelled(
                    cancel_event,
                    detail=f"cancelled after processing {file_path.name}",
                )
            except CancelledError:
                raise
            except Exception as exc:
                logger.error("处理失败 %s: %s", file_path.name, exc)
                summary.failed += 1
                continue

            if ok:
                summary.processed += 1
                stem = file_path.stem
                marked = output_dir / f"{stem}_已标记图片.docx"
                pdf = output_dir / f"{stem}_附图.pdf"
                if marked.exists():
                    summary.outputs.append(str(marked))
                if pdf.exists():
                    summary.outputs.append(str(pdf))
            else:
                summary.failed += 1

        if summary.processed == 0 and summary.failed > 0:
            raise ProcessFailedError(
                f"全部 {summary.failed} 个文件处理失败",
                detail=f"input={request.input_path}",
            )

        return summary
