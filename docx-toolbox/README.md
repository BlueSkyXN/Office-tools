# docx-toolbox

统一的文档工具箱工程，包含共享内核 + `pyside6` GUI 子项目 + 参考脚本。

## 功能

- **Excel 嵌入对象处理**：将 DOCX 中嵌入的 Excel 表格转换为 Word 原生表格/图片/独立文件
- **图片分离**：提取 DOCX 中的图片并生成带目录的 PDF 附图集
- **表格提取**：提取 DOCX 表格并导出为 TXT/XLSX/PDF 格式

## 快速开始

```bash
cd docx-toolbox

# 安装基础依赖
pip install -e .

# 安装 GUI 依赖（PySide6）
pip install -e ".[pyside6]"

# 启动 GUI
python3 -m pyside6.app.main

# 运行测试
python3 -m pytest tests/ -v
```

## 目录结构

```text
docx-toolbox/
  core/                    # 共享内核（适配器、执行器、错误码、日志）
    adapters/              #   三类任务适配器
    runner/                #   串行/并行任务执行器
    errors/                #   统一错误码与异常类
    logging_utils/         #   统一日志模块
    api.py                 #   调度入口（run_task）
  references/              # 参考脚本（docx-allinone / 图片分离 / 表格提取）
  pyside6/                 # GUI 子项目
  tests/                   # 共享内核测试
  pyproject.toml           # 项目依赖与构建配置
  CORE-INTERFACE.md        # 跨项目共享接口规范
```

## 架构

```text
┌─────────────────────────────────────────────┐
│               pyside6 GUI                    │
│            app/core/adapter.py               │
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

- 所有 GUI 调用通过 `core/` 共享内核，禁止直接调用参考脚本
- 接口规范：[`CORE-INTERFACE.md`](./CORE-INTERFACE.md)

## 当前决策

- 唯一 GUI 路线：`pyside6`

## GitHub Actions 打包

- 双平台打包：macOS arm64 + Windows x64
- 策略详见：[`docs/GITHUB-ACTIONS-PACKAGING.md`](./docs/GITHUB-ACTIONS-PACKAGING.md)
- 已落地 workflow：[`../.github/workflows/package-docx-pyside6.yml`](../.github/workflows/package-docx-pyside6.yml)

## 本地打包（PyInstaller）

```bash
cd docx-toolbox

# 默认 auto：macOS 输出 onedir + .app（双击无终端），其他平台输出 onefile
python3 scripts/package_app.py --app pyside6 --version dev-local

# 若你明确需要 macOS 单文件可执行（会以终端进程形式运行）
python3 scripts/package_app.py --app pyside6 --bundle-mode onefile --version dev-local
```
