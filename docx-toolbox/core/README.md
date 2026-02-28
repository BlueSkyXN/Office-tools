# core

共享内核包，提供统一的任务调度、适配器、执行器、错误码与日志。

## 模块结构

```
core/
  __init__.py          # 导出 run_task, TaskRequest, TaskResponse
  api.py               # 调度入口与请求/响应模型
  adapters/
    __init__.py        # BaseAdapter 基类与公共校验
    excel_allinone.py  # Excel 嵌入对象处理适配器
    image_extract.py   # 图片分离适配器
    table_extract.py   # 表格提取适配器
  runner/
    __init__.py        # TaskRunner（串行/并行/取消/重试）
  errors/
    __init__.py        # ErrorCode 枚举与 TaskError 异常类
  logging_utils/
    __init__.py        # 统一日志格式与文件落盘
```

## 使用示例

```python
from core.api import TaskRequest, run_task

request = TaskRequest(
    task_type="excel_allinone",
    input_path="/path/to/document.docx",
    options={"word_table": True, "remove_watermark": True},
)
response = run_task(request)
print(response.to_dict())
```

## 接口规范

遵循 `../CORE-INTERFACE.md`，所有任务统一返回 `ok/status/summary/error` 结构。
