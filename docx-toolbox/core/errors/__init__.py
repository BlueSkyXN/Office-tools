"""错误码枚举与异常类 — 遵循 CORE-INTERFACE.md §4"""

from enum import Enum


class ErrorCode(str, Enum):
    INVALID_INPUT = "E_INVALID_INPUT"
    UNSUPPORTED_FORMAT = "E_UNSUPPORTED_FORMAT"
    PERMISSION_DENIED = "E_PERMISSION_DENIED"
    PROCESS_FAILED = "E_PROCESS_FAILED"
    CANCELLED = "E_CANCELLED"
    INTERNAL = "E_INTERNAL"


class TaskError(Exception):
    """所有 core 任务异常的基类"""

    def __init__(self, code: ErrorCode, message: str, detail: str = ""):
        self.code = code
        self.message = message
        self.detail = detail
        super().__init__(message)

    def to_dict(self) -> dict:
        return {
            "code": self.code.value,
            "message": self.message,
            "detail": self.detail,
        }


class InvalidInputError(TaskError):
    def __init__(self, message: str, detail: str = ""):
        super().__init__(ErrorCode.INVALID_INPUT, message, detail)


class UnsupportedFormatError(TaskError):
    def __init__(self, message: str, detail: str = ""):
        super().__init__(ErrorCode.UNSUPPORTED_FORMAT, message, detail)


class PermissionDeniedError(TaskError):
    def __init__(self, message: str, detail: str = ""):
        super().__init__(ErrorCode.PERMISSION_DENIED, message, detail)


class ProcessFailedError(TaskError):
    def __init__(self, message: str, detail: str = ""):
        super().__init__(ErrorCode.PROCESS_FAILED, message, detail)


class CancelledError(TaskError):
    def __init__(self, message: str = "任务被用户取消", detail: str = ""):
        super().__init__(ErrorCode.CANCELLED, message, detail)
