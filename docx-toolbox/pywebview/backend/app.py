"""pywebview 窗口创建与启动入口"""

import sys
from pathlib import Path

# 将 docx-toolbox 根目录与 pywebview 根目录加入 sys.path
# 这样在从 `docx-toolbox` 根目录直接运行 `python3 pywebview/backend/app.py` 时，
# 既能导入 core，也能导入 backend 包。
backend_root = Path(__file__).resolve().parent.parent
project_root = backend_root.parent
sys.path.insert(0, str(project_root))
sys.path.insert(0, str(backend_root))

import webview

from backend.api import ApiBridge


def main():
    api = ApiBridge()
    dev_mode = "--dev" in sys.argv

    if dev_mode:
        url = "http://localhost:5173"
    else:
        url = str(Path(__file__).parent.parent / "frontend" / "dist" / "index.html")

    window = webview.create_window(
        "DOCX 工具箱",
        url,
        js_api=api,
        width=1200,
        height=800,
        min_size=(900, 600),
    )
    webview.start(debug=dev_mode)


if __name__ == "__main__":
    main()
