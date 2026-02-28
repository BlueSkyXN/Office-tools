"""Flet 主题配置 — 遵循 DESIGN.md §5 OpenAI 风格浅色"""

import flet as ft


# 语义色彩 token
COLOR_PRIMARY = "#3B82F6"
COLOR_SURFACE = "#FFFFFF"
COLOR_BACKGROUND = "#F8FAFC"
COLOR_OUTLINE = "#E5E7EB"
COLOR_ON_SURFACE = "#111827"
COLOR_SUCCESS = "#16A34A"
COLOR_WARNING = "#D97706"
COLOR_ERROR = "#DC2626"


def create_theme() -> ft.Theme:
    """构建全局 Flet Theme"""
    return ft.Theme(
        color_scheme=ft.ColorScheme(
            primary=COLOR_PRIMARY,
            surface=COLOR_SURFACE,
            surface_container_lowest=COLOR_BACKGROUND,
            outline=COLOR_OUTLINE,
            on_surface=COLOR_ON_SURFACE,
            error=COLOR_ERROR,
        ),
        visual_density=ft.VisualDensity.COMFORTABLE,
    )
