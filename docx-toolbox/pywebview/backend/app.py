"""pywebview 窗口创建与启动入口"""

import sys
from pathlib import Path

# 将 docx-toolbox 根目录加入 sys.path，以便导入 core
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

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
