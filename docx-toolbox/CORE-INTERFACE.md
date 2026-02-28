# CORE-INTERFACE

定义 `docx-toolbox` 所有 GUI 子项目共享的内核接口约定。

## 1. 任务类型

- `excel_allinone`：对应 `docx-allinone.py`
- `image_extract`：对应 `DOCX图片分离.py`
- `table_extract`：对应 `DOCX表格提取.py`

## 2. 通用请求模型

```json
{
  "task_id": "uuid-string",
  "task_type": "excel_allinone|image_extract|table_extract",
  "input_path": "string",
  "output_dir": "string|null",
  "options": {},
  "runtime": {
    "workers": 1,
    "dry_run": false
  }
}
```

## 2.1 任务级 options 定义

- `excel_allinone`：
  - `word_table` (bool)：是否转换为 Word 原生表格
  - `extract_excel` (bool)：是否提取嵌入 Excel 文件
  - `image` (bool)：是否渲染为图片
  - `keep_attachment` (bool)：是否保留附件入口
  - `remove_watermark` (bool)：是否移除水印
  - `a3` (bool)：是否设置 A3 横向
- `image_extract`：
  - `remove_images` (bool)：是否删除原图并仅保留标记
  - `optimize_images` (bool)：是否启用图片优化
  - `jpeg_quality` (int, 1-100)：JPEG 质量
- `table_extract`：
  - `include_marked` (bool)：是否包含已标记表格文件（兼容参数，默认 false）

说明：
- `options` 仅包含任务语义参数，不包含运行时参数。
- 并发数、dry-run 等执行参数统一放在 `runtime`。

## 2.2 core 与子项目适配层边界

- 顶层 `docx-toolbox/core/` 是唯一共享内核实现层。
- 各子项目中的 `app/core/` 仅作为适配层：
  - 将 GUI 参数映射为 `CORE-INTERFACE` 请求模型。
  - 调用顶层 `core` 的稳定接口。
  - 不重复实现核心处理逻辑。

## 3. 通用响应模型

成功：

```json
{
  "ok": true,
  "task_id": "uuid-string",
  "status": "success",
  "summary": {
    "processed": 0,
    "failed": 0,
    "skipped": 0,
    "outputs": []
  }
}
```

失败：

```json
{
  "ok": false,
  "task_id": "uuid-string",
  "status": "failed",
  "error": {
    "code": "E_INVALID_INPUT",
    "message": "...",
    "detail": "..."
  }
}
```

## 4. 错误码枚举

- `E_INVALID_INPUT`：输入路径无效或类型错误
- `E_UNSUPPORTED_FORMAT`：文件格式不支持
- `E_PERMISSION_DENIED`：无文件权限
- `E_PROCESS_FAILED`：处理过程异常
- `E_CANCELLED`：任务被用户取消
- `E_INTERNAL`：未分类内部错误

## 5. 日志规范

- UI 仅展示最近 500 行日志。
- 完整日志写入文件：`logs/<date>/<task_id>.log`。
- 日志格式：

```text
2026-02-28T10:00:00+08:00 | INFO  | excel_allinone | task=<id> | message
```

## 6. 测试规范（最低要求）

- `core` 单元测试：
  - 每个任务类型至少 3 个有效样例 + 2 个异常样例。
- `runner` 集成测试：
  - 批处理继续执行、取消、重试、并发稳定性。
- `ui` 冒烟测试：
  - 打开页面、提交任务、展示结果、错误提示。
