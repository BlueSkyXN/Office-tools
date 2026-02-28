# docx-toolbox

统一的文档工具箱工程，包含共享内核 + 5 个 GUI 框架子项目 + 参考脚本。

## 功能

- **Excel 嵌入对象处理**：将 DOCX 中嵌入的 Excel 表格转换为 Word 原生表格/图片/独立文件
- **图片分离**：提取 DOCX 中的图片并生成带目录的 PDF 附图集
- **表格提取**：提取 DOCX 表格并导出为 TXT/XLSX/PDF 格式

## 快速开始

```bash
# 安装基础依赖
cd docx-toolbox
pip install -e .

# 安装特定 GUI 框架依赖
pip install -e ".[pyside6]"   # 生产默认
pip install -e ".[flet]"      # Flet 方案
pip install -e ".[pywebview]" # React + pywebview 方案

# 启动 GUI
python3 -m pyside6.app.main   # PySide6（生产默认）
python3 -m pyqt6.app.main     # PyQt6（内部评估）
python3 -m tk.app.main         # Tkinter
python3 flet/app/main.py       # Flet
python3 pywebview/backend/app.py  # pywebview

# 运行测试
python3 -m pytest tests/ -v
```

## 目录结构

```
docx-toolbox/
  core/                    # 共享内核（适配器、执行器、错误码、日志）
    adapters/              #   三类任务适配器
    runner/                #   串行/并行任务执行器
    errors/                #   统一错误码与异常类
    logging_utils/         #   统一日志模块
    api.py                 #   调度入口（run_task）
  references/              # 参考脚本（docx-allinone / 图片分离 / 表格提取）
  pyside6/                 # PySide6 子项目（生产默认）
  pyqt6/                   # PyQt6 子项目（内部评估）
  tk/                      # Tkinter 子项目（轻量方案）
  flet/                    # Flet 子项目（现代 Python UI）
  pywebview/               # React + pywebview 子项目（Web 技术栈）
    frontend/              #   React + Vite + TypeScript
    backend/               #   Python pywebview 后端
  tests/                   # 共享内核测试
  pyproject.toml           # 项目依赖与构建配置
  CORE-INTERFACE.md        # 跨项目共享接口规范
  DESIGN.md                # 主设计文档
```

## 架构

```
┌─────────────────────────────────────────────┐
│              GUI 子项目                      │
│  pyside6 │ pyqt6 │ tk │ flet │ pywebview   │
│   app/core/adapter.py (参数适配层)           │
├─────────────────────────────────────────────┤
│              core/ 共享内核                   │
│  api.run_task() → adapters → references     │
│  runner.TaskRunner (串行/并行/取消/重试)     │
│  errors (E_INVALID_INPUT / E_PROCESS_FAILED)│
│  logging_utils (统一日志格式与文件落盘)      │
├─────────────────────────────────────────────┤
│           references/ 参考脚本               │
│  docx-allinone.py │ 图片分离.py │ 表格提取.py│
└─────────────────────────────────────────────┘
```

## 统一约束

- 所有子项目通过 `core/` 共享内核调用业务能力，禁止直接调用参考脚本
- 接口规范：[`CORE-INTERFACE.md`](./CORE-INTERFACE.md)
- PyQt6 许可：[`docs/PYQT6-LICENSE-POLICY.md`](./docs/PYQT6-LICENSE-POLICY.md)

## 当前决策

- 生产默认路线：`pyside6`
- `pyqt6` 路线：仅内部评估与对比，不做外部分发

## 依赖管理

- 根级 `pyproject.toml` 维护统一依赖基线
- `core/` 作为共享包，各子项目通过路径导入
- GUI 框架依赖通过 optional-dependencies 按需安装

## GitHub Actions 打包

- 双平台打包：macOS arm64 + Windows x64
- 策略详见：[`docs/GITHUB-ACTIONS-PACKAGING.md`](./docs/GITHUB-ACTIONS-PACKAGING.md)
- 已落地 workflow：[`../.github/workflows/package-docx-toolbox.yml`](../.github/workflows/package-docx-toolbox.yml)

## 本地打包（PyInstaller）

```bash
cd docx-toolbox

# 例：打包 pyside6
python3 scripts/package_app.py --app pyside6 --version dev-local

# 例：打包 pywebview（需先构建前端）
cd pywebview/frontend && npm ci && npm run build && cd ../..
python3 scripts/package_app.py --app pywebview --version dev-local

# 默认 auto：所有平台统一输出单文件（onefile），macOS 不使用 --windowed 以规避 PyInstaller 7.0 弃用
python3 scripts/package_app.py --app pyside6 --bundle-mode onefile --version dev-local
```
