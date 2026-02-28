# Changelog

## [0.1.2] - 2026-02-28

### Changed

- 项目 GUI 路线收敛为 `pyside6`，移除 `flet` 与 `pywebview` 代码与打包流程
- GitHub Actions 仅保留 `/.github/workflows/package-docx-pyside6.yml`
- 打包脚本 `scripts/package_app.py` 仅支持 `--app pyside6`

## [0.1.1] - 2026-02-28

### Added

- GitHub Actions 实际打包工作流：`/.github/workflows/package-docx-all.yml`
- 新增统一打包脚本：`scripts/package_app.py`

### Changed

- `TaskRunner` 并行调度改为“限量提交 + 取消后停止提交”，并将 `E_CANCELLED` 映射为 `JobStatus.CANCELLED`
- `pywebview` 后台 worker 将 `cancel_event` 传递到 core，取消响应不再仅停留在回调层
- `excel_allinone` 适配器改为“单文件失败记失败并继续”，不再中断整批任务
- `TaskService` 去除锁外写入，统一在锁内更新 `_tasks` 和 `TaskRecord` 状态
- 根 `.gitignore` 补齐 PyInstaller `*.spec` 忽略规则
- 移除 `pyqt6` 与 `tk` 两条 GUI 路线（该版本仍保留 `pyside6/flet/pywebview`）

### Fixed

- `tests/test_core.py` 的取消用例补充关键断言，新增并行取消回归测试
- 新增 `excel_allinone` 批处理容错回归测试（`ProcessFailedError` 不再中断批次）

## [0.1.0] - 2026-02-28

### Added

- **共享内核 (`core/`)**
  - 统一调度入口 `run_task()` 与请求/响应模型
  - 三类任务适配器：`excel_allinone`、`image_extract`、`table_extract`
  - 串行/并行任务执行器 `TaskRunner`（支持取消与重试）
  - 错误码枚举（6 种标准错误码）
  - 统一日志模块（500 行 UI 上限 + 文件落盘）
  - 17 个单元测试全部通过

- **PySide6 子项目 (`pyside6/`)** — 生产默认
  - 左侧导航 + 右侧工作区 + 底部日志面板
  - Excel/图片/表格/批处理/设置 5 个功能页面
  - QThread 非阻塞任务执行
  - OpenAI 浅色风格 QSS 主题

- **PyQt6 子项目 (`pyqt6/`)** — 内部评估
  - 与 PySide6 功能对齐，使用 PyQt6 API

- **Tkinter 子项目 (`tk/`)**
  - Notebook 标签页布局
  - ttk.Style 浅色主题

- **Flet 子项目 (`flet/`)**
  - NavigationRail + 页面切换 + 事件总线状态管理
  - Flet Theme 统一 token

- **pywebview 子项目 (`pywebview/`)**
  - React 18 + Vite + TypeScript 前端（已构建）
  - Python pywebview 后端 + JS Bridge API
  - CSS 设计 token 体系

- **项目基础设施**
  - `pyproject.toml` 统一依赖管理
  - `CORE-INTERFACE.md` 跨项目接口规范
  - `GITHUB-ACTIONS-PACKAGING.md` CI 打包设计
  - `PYQT6-LICENSE-POLICY.md` 许可管理策略
