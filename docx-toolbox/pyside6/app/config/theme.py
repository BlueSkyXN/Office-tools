"""QSS theme with semantic tokens — OpenAI light style (DESIGN.md §5)"""

# Semantic color tokens
COLORS = {
    "primary":    "#3B82F6",
    "primary_hover": "#2563EB",
    "background": "#F8FAFC",
    "card":       "#FFFFFF",
    "text":       "#111827",
    "text_secondary": "#6B7280",
    "border":     "#E5E7EB",
    "success":    "#16A34A",
    "warning":    "#D97706",
    "error":      "#DC2626",
    "nav_bg":     "#F1F5F9",
    "nav_active": "#E0ECFF",
    "input_bg":   "#FFFFFF",
    "log_bg":     "#1E293B",
    "log_text":   "#E2E8F0",
}

STYLESHEET = f"""
/* ---- Global ---- */
QWidget {{
    font-family: "Segoe UI", "SF Pro Text", "PingFang SC", "Microsoft YaHei", sans-serif;
    font-size: 13px;
    color: {COLORS['text']};
    background-color: {COLORS['background']};
}}

/* ---- Main Window ---- */
QMainWindow {{
    background-color: {COLORS['background']};
}}

/* ---- Left Navigation ---- */
#nav_panel {{
    background-color: {COLORS['nav_bg']};
    border-right: 1px solid {COLORS['border']};
}}
#nav_panel QPushButton {{
    text-align: left;
    padding: 10px 18px;
    border: none;
    border-radius: 6px;
    margin: 2px 6px;
    font-size: 14px;
    color: {COLORS['text']};
    background-color: transparent;
}}
#nav_panel QPushButton:hover {{
    background-color: {COLORS['border']};
}}
#nav_panel QPushButton:checked {{
    background-color: {COLORS['nav_active']};
    color: {COLORS['primary']};
    font-weight: 600;
}}

/* ---- Cards / Group boxes ---- */
QGroupBox {{
    background-color: {COLORS['card']};
    border: 1px solid {COLORS['border']};
    border-radius: 8px;
    margin-top: 12px;
    padding: 16px 12px 12px 12px;
    font-weight: 600;
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    left: 14px;
    padding: 0 4px;
    color: {COLORS['text']};
}}

/* ---- Inputs ---- */
QLineEdit, QSpinBox, QComboBox {{
    padding: 6px 10px;
    border: 1px solid {COLORS['border']};
    border-radius: 6px;
    background-color: {COLORS['input_bg']};
    selection-background-color: {COLORS['primary']};
}}
QLineEdit:focus, QSpinBox:focus, QComboBox:focus {{
    border-color: {COLORS['primary']};
}}

/* ---- Buttons ---- */
QPushButton {{
    padding: 7px 18px;
    border: 1px solid {COLORS['border']};
    border-radius: 6px;
    background-color: {COLORS['card']};
}}
QPushButton:hover {{
    background-color: {COLORS['background']};
}}
QPushButton#btn_start {{
    background-color: {COLORS['primary']};
    color: white;
    border: none;
    font-weight: 600;
}}
QPushButton#btn_start:hover {{
    background-color: {COLORS['primary_hover']};
}}
QPushButton#btn_start:disabled {{
    background-color: {COLORS['border']};
    color: {COLORS['text_secondary']};
}}
QPushButton#btn_stop {{
    background-color: {COLORS['error']};
    color: white;
    border: none;
}}
QPushButton#btn_stop:disabled {{
    background-color: {COLORS['border']};
    color: {COLORS['text_secondary']};
}}

/* ---- Check boxes ---- */
QCheckBox {{
    spacing: 6px;
}}
QCheckBox::indicator {{
    width: 16px;
    height: 16px;
    border: 1px solid {COLORS['border']};
    border-radius: 3px;
    background-color: {COLORS['input_bg']};
}}
QCheckBox::indicator:checked {{
    background-color: {COLORS['primary']};
    border-color: {COLORS['primary']};
}}

/* ---- Progress bar ---- */
QProgressBar {{
    border: 1px solid {COLORS['border']};
    border-radius: 6px;
    text-align: center;
    height: 18px;
    background-color: {COLORS['background']};
}}
QProgressBar::chunk {{
    background-color: {COLORS['primary']};
    border-radius: 5px;
}}

/* ---- Log Panel ---- */
#log_panel {{
    background-color: {COLORS['log_bg']};
    border-top: 1px solid {COLORS['border']};
}}
#log_text {{
    background-color: {COLORS['log_bg']};
    color: {COLORS['log_text']};
    border: none;
    font-family: "Menlo", "Consolas", "Courier New", monospace;
    font-size: 12px;
    padding: 6px;
}}

/* ---- Table widget ---- */
QTableWidget {{
    border: 1px solid {COLORS['border']};
    border-radius: 6px;
    gridline-color: {COLORS['border']};
    background-color: {COLORS['card']};
}}
QTableWidget::item {{
    padding: 4px 8px;
}}
QHeaderView::section {{
    background-color: {COLORS['nav_bg']};
    border: none;
    border-bottom: 1px solid {COLORS['border']};
    padding: 6px 8px;
    font-weight: 600;
}}

/* ---- Status bar ---- */
QStatusBar {{
    background-color: {COLORS['nav_bg']};
    border-top: 1px solid {COLORS['border']};
    color: {COLORS['text_secondary']};
    font-size: 12px;
}}

/* ---- Scroll bars ---- */
QScrollBar:vertical {{
    width: 8px;
    background: transparent;
}}
QScrollBar::handle:vertical {{
    background: {COLORS['border']};
    border-radius: 4px;
    min-height: 30px;
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0px;
}}

/* ---- Splitter ---- */
QSplitter::handle {{
    background-color: {COLORS['border']};
}}
QSplitter::handle:vertical {{
    height: 1px;
}}
"""
