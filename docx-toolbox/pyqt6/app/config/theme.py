"""主题 token 定义 — 语义驱动样式，禁止页面内硬编码颜色"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class ThemeTokens:
    primary: str = "#3B82F6"
    text_primary: str = "#111827"
    text_secondary: str = "#6B7280"
    bg_canvas: str = "#F8FAFC"
    bg_card: str = "#FFFFFF"
    border: str = "#E5E7EB"
    success: str = "#16A34A"
    warning: str = "#D97706"
    error: str = "#DC2626"
    nav_bg: str = "#F1F5F9"
    nav_hover: str = "#E2E8F0"
    nav_active: str = "#DBEAFE"


LIGHT = ThemeTokens()


def build_stylesheet(t: ThemeTokens | None = None) -> str:
    """根据 token 生成完整 QSS"""
    if t is None:
        t = LIGHT

    return f"""
    /* ---- 全局 ---- */
    QWidget {{
        font-family: "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
        font-size: 13px;
        color: {t.text_primary};
        background-color: {t.bg_canvas};
    }}

    /* ---- 主窗口 ---- */
    QMainWindow {{
        background-color: {t.bg_canvas};
    }}

    /* ---- 导航列表 ---- */
    QListWidget {{
        background-color: {t.nav_bg};
        border: none;
        border-right: 1px solid {t.border};
        outline: none;
        padding: 6px 0;
    }}
    QListWidget::item {{
        padding: 10px 18px;
        border-radius: 6px;
        margin: 2px 6px;
        color: {t.text_primary};
    }}
    QListWidget::item:hover {{
        background-color: {t.nav_hover};
    }}
    QListWidget::item:selected {{
        background-color: {t.nav_active};
        color: {t.primary};
        font-weight: 600;
    }}

    /* ---- 卡片容器 ---- */
    QGroupBox {{
        background-color: {t.bg_card};
        border: 1px solid {t.border};
        border-radius: 8px;
        margin-top: 12px;
        padding: 16px 12px 12px 12px;
        font-weight: 600;
    }}
    QGroupBox::title {{
        subcontrol-origin: margin;
        left: 14px;
        padding: 0 6px;
        color: {t.text_primary};
    }}

    /* ---- 按钮 ---- */
    QPushButton {{
        background-color: {t.primary};
        color: #FFFFFF;
        border: none;
        border-radius: 6px;
        padding: 7px 20px;
        font-weight: 500;
    }}
    QPushButton:hover {{
        background-color: #2563EB;
    }}
    QPushButton:pressed {{
        background-color: #1D4ED8;
    }}
    QPushButton:disabled {{
        background-color: {t.border};
        color: {t.text_secondary};
    }}
    QPushButton[secondary="true"] {{
        background-color: {t.bg_card};
        color: {t.text_primary};
        border: 1px solid {t.border};
    }}
    QPushButton[secondary="true"]:hover {{
        background-color: {t.nav_hover};
    }}

    /* ---- 输入框 ---- */
    QLineEdit {{
        border: 1px solid {t.border};
        border-radius: 6px;
        padding: 6px 10px;
        background-color: {t.bg_card};
    }}
    QLineEdit:focus {{
        border-color: {t.primary};
    }}

    /* ---- 复选框 ---- */
    QCheckBox {{
        spacing: 6px;
    }}
    QCheckBox::indicator {{
        width: 16px;
        height: 16px;
        border: 1px solid {t.border};
        border-radius: 3px;
        background-color: {t.bg_card};
    }}
    QCheckBox::indicator:checked {{
        background-color: {t.primary};
        border-color: {t.primary};
    }}

    /* ---- 滑块 ---- */
    QSlider::groove:horizontal {{
        height: 4px;
        background: {t.border};
        border-radius: 2px;
    }}
    QSlider::handle:horizontal {{
        background: {t.primary};
        width: 14px;
        height: 14px;
        margin: -5px 0;
        border-radius: 7px;
    }}

    /* ---- SpinBox ---- */
    QSpinBox {{
        border: 1px solid {t.border};
        border-radius: 6px;
        padding: 4px 8px;
        background-color: {t.bg_card};
    }}

    /* ---- 日志区 ---- */
    QTextEdit[readOnly="true"], QPlainTextEdit[readOnly="true"] {{
        background-color: {t.bg_card};
        border: 1px solid {t.border};
        border-radius: 6px;
        font-family: "JetBrains Mono", "Menlo", "Consolas", monospace;
        font-size: 12px;
        color: {t.text_primary};
        padding: 8px;
    }}

    /* ---- 标签 ---- */
    QLabel {{
        background-color: transparent;
    }}
    QLabel[heading="true"] {{
        font-size: 16px;
        font-weight: 700;
        color: {t.text_primary};
    }}

    /* ---- 分割器 ---- */
    QSplitter::handle {{
        background-color: {t.border};
    }}
    QSplitter::handle:horizontal {{
        width: 1px;
    }}
    QSplitter::handle:vertical {{
        height: 1px;
    }}

    /* ---- 进度条 ---- */
    QProgressBar {{
        border: 1px solid {t.border};
        border-radius: 4px;
        text-align: center;
        background-color: {t.nav_bg};
        height: 18px;
    }}
    QProgressBar::chunk {{
        background-color: {t.primary};
        border-radius: 3px;
    }}
    """
