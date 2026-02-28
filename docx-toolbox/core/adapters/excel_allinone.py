"""ExcelAllinone 适配器 — 将 references/docx-allinone.py 封装为 core 统一接口"""

from __future__ import annotations

import argparse
import importlib.util
import io
import os
import sys
from contextlib import redirect_stdout, redirect_stderr
from pathlib import Path

from core.adapters import BaseAdapter
from core.api import TaskRequest, TaskSummary
from core.errors import CancelledError, InvalidInputError, ProcessFailedError
from core.logging_utils import get_logger

logger = get_logger()

# ---------------------------------------------------------------------------
# 懒加载 reference 模块（文件名含连字符，无法直接 import）
# ---------------------------------------------------------------------------
_ref_mod = None
_REF_DIR = str(Path(__file__).resolve().parent.parent.parent / "references")


def _load_ref_module():
    global _ref_mod
    if _ref_mod is not None:
        return _ref_mod
    ref_path = os.path.join(_REF_DIR, "docx-allinone.py")
    if not os.path.isfile(ref_path):
        raise ProcessFailedError(
            f"参考脚本不存在: {ref_path}",
            detail="请确认 references/docx-allinone.py 文件存在",
        )
    spec = importlib.util.spec_from_file_location("docx_allinone_ref", ref_path)
    _ref_mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(_ref_mod)
    return _ref_mod


# ---------------------------------------------------------------------------
# 输出文件路径推断（与 reference process_document 逻辑保持一致）
# ---------------------------------------------------------------------------

def _predict_output_path(input_path: str, args: argparse.Namespace) -> str:
    """根据 reference 命名规则推算输出文件路径"""
    base, ext = os.path.splitext(input_path)
    suffix_parts = ["-AIO"]
    if args.keep_attachment:
        suffix_parts.append("WithAttachments")
    if getattr(args, "remove_watermark", False):
        suffix_parts.append("NoWM")
    if getattr(args, "a3", False):
        suffix_parts.append("A3")
    suffix = "-" + "-".join(suffix_parts[1:]) if len(suffix_parts) > 1 else suffix_parts[0]
    return f"{base}{suffix}{ext}"


# ---------------------------------------------------------------------------
# 适配器
# ---------------------------------------------------------------------------

# reference 脚本用于过滤已处理文件的标签
_OUTPUT_FILE_TAGS = ["-WithAttachments", "-NoWM", "-A3", "-AIO"]


class ExcelAllinoneAdapter(BaseAdapter):
    """将 references/docx-allinone.py 的 process_document 包装为 core 适配器"""

    def execute(self, request: TaskRequest) -> TaskSummary:
        path = self.validate_input_path(request.input_path)
        cancel_event = request.runtime.cancel_event
        # output_dir 仅用于日志，实际输出路径由 reference 脚本决定（同目录）
        self.ensure_output_dir(
            request.output_dir, path.parent if path.is_file() else path
        )

        opts = request.options

        # 构建与 reference argparse.Namespace 兼容的参数对象
        args = argparse.Namespace(
            word_table=opts.get("word_table", False),
            extract_excel=opts.get("extract_excel", False),
            image=opts.get("image", False),
            keep_attachment=opts.get("keep_attachment", False),
            remove_watermark=opts.get("remove_watermark", False),
            a3=opts.get("a3", False),
            workers=request.runtime.workers,
            input_path=str(path),
        )

        # 未指定任何模式时默认启用 word_table
        if not any([
            args.word_table, args.extract_excel, args.image,
            args.remove_watermark, args.a3,
        ]):
            args.word_table = True

        # 收集待处理文件
        if path.is_file():
            self.validate_docx(path)
            files = [path]
        else:
            files = self.collect_docx_files(path)
            # 过滤已处理文件（文件名含输出标签）
            files = [
                f for f in files
                if not any(tag in f.stem for tag in _OUTPUT_FILE_TAGS)
            ]

        if not files:
            logger.info("未找到待处理的 .docx 文件")
            return TaskSummary(processed=0, failed=0, skipped=0, outputs=[])

        # 懒加载 reference 模块
        ref = _load_ref_module()

        processed = 0
        failed = 0
        skipped = 0
        outputs: list[str] = []

        for file_path in files:
            file_name = file_path.name
            self.ensure_not_cancelled(
                cancel_event,
                detail=f"cancelled before processing {file_name}",
            )
            logger.info("开始处理: %s", file_name)
            try:
                # reference 的 process_document 使用 print() 输出日志，
                # 此处捕获 stdout/stderr 转写到 logger
                buf = io.StringIO()
                with redirect_stdout(buf), redirect_stderr(buf):
                    try:
                        ref.process_document(str(file_path), args)
                    except SystemExit as exc:
                        # process_document 在某些分支会 sys.exit()
                        if exc.code not in (None, 0):
                            raise ProcessFailedError(
                                f"处理跳过: {file_name}",
                                detail=buf.getvalue().strip(),
                            )
                        # exit(0) 视为成功

                self.ensure_not_cancelled(
                    cancel_event,
                    detail=f"cancelled after processing {file_name}",
                )

                captured = buf.getvalue()
                if captured:
                    for line in captured.splitlines():
                        if line.strip():
                            logger.debug("[ref] %s", line.rstrip())

                out_path = _predict_output_path(str(file_path), args)
                if os.path.isfile(out_path):
                    outputs.append(out_path)
                    processed += 1
                    logger.info("处理完成: %s -> %s", file_name, out_path)
                else:
                    # 输出文件未生成，视为跳过
                    skipped += 1
                    logger.warning("输出文件未生成: %s", out_path)

            except CancelledError:
                raise
            except ProcessFailedError as exc:
                failed += 1
                logger.error("处理失败: %s — %s", file_name, exc)
            except Exception as exc:
                failed += 1
                logger.error("处理失败: %s — %s", file_name, exc)

        return TaskSummary(
            processed=processed,
            failed=failed,
            skipped=skipped,
            outputs=outputs,
        )
