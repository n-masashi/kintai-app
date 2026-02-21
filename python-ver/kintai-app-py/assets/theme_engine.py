"""QSS テーマ文字列生成・適用"""
from PyQt5.QtWidgets import QApplication


# ════════════════════════════════════════════════
# カラートークン辞書
# ════════════════════════════════════════════════
THEME_COLORS = {
    "light": {
        # --- Surface / Background ---
        "bg_primary":       "#FFFFFF",
        "bg_secondary":     "#F8FAFC",
        "bg_tertiary":      "#F1F5F9",
        "bg_elevated":      "#FFFFFF",

        # --- Text ---
        "text_primary":     "#1E293B",
        "text_secondary":   "#475569",
        "text_tertiary":    "#94A3B8",
        "text_on_accent":   "#FFFFFF",

        # --- Border ---
        "border_default":   "#E2E8F0",
        "border_strong":    "#CBD5E1",
        "border_focus":     "#3B82F6",

        # --- Accent (Primary Blue) ---
        "accent":           "#3B82F6",
        "accent_hover":     "#2563EB",
        "accent_pressed":   "#1D4ED8",
        "accent_subtle":    "#DBEAFE",
        "accent_text":      "#1E40AF",

        # --- Semantic: Success ---
        "success":          "#10B981",
        "success_hover":    "#059669",
        "success_pressed":  "#047857",

        # --- Semantic: Danger ---
        "danger":           "#EF4444",
        "danger_hover":     "#DC2626",
        "danger_pressed":   "#B91C1C",

        # --- Semantic: Warning ---
        "warning":          "#F59E0B",
        "warning_hover":    "#D97706",

        # --- Semantic: Secondary ---
        "secondary":        "#64748B",
        "secondary_hover":  "#475569",

        # --- Component-specific ---
        "tab_active_bg":    "#FFFFFF",
        "tab_inactive_bg":  "#F1F5F9",
        "tab_active_text":  "#3B82F6",
        "tab_border":       "#E2E8F0",
        "input_bg":         "#FFFFFF",
        "input_border":     "#CBD5E1",
        "input_focus":      "#3B82F6",
        "table_header_bg":  "#F1F5F9",
        "table_alt_row":    "#F8FAFC",
        "table_selected":   "#DBEAFE",
        "table_grid":       "#E2E8F0",
        "scrollbar_bg":     "#F1F5F9",
        "scrollbar_handle": "#CBD5E1",
        "group_border":     "#E2E8F0",
        "selection_bg":     "#DBEAFE",
        "selection_text":   "#1E40AF",
        "disabled_bg":      "#E2E8F0",
        "disabled_text":    "#94A3B8",

        # --- Button overrides (clock in / clock out) ---
        "btn_clock_in":         "#10B981",
        "btn_clock_in_hover":   "#059669",
        "btn_clock_in_pressed": "#047857",
        "btn_clock_out":        "#EF4444",
        "btn_clock_out_hover":  "#DC2626",
        "btn_clock_out_pressed":"#B91C1C",


        # --- Calendar-specific ---
        "cal_normal_bg":       "#FAFAFA",
        "cal_normal_text":     "#374151",
        "cal_holiday_bg":      "#FEF2F2",
        "cal_holiday_text":    "#EF4444",
        "cal_saturday_bg":     "#EFF6FF",
        "cal_saturday_text":   "#3B82F6",
        "cal_today_bg":        "#DBEAFE",
        "cal_today_text":      "#1E40AF",
        "cal_today_border":    "#3B82F6",
        "cal_selected_bg":     "#3B82F6",
        "cal_selected_text":   "#FFFFFF",
        "cal_selected_border": "#1D4ED8",
        "cal_hover_normal":    "#E2E8F0",
        "cal_hover_holiday":   "#FEE2E2",
        "cal_hover_saturday":  "#DBEAFE",
        "cal_hover_today":     "#BFDBFE",
        "cal_header_sunday":   "#EF4444",
        "cal_header_saturday": "#3B82F6",
    },
    "dark": {
        # --- Surface / Background ---
        "bg_primary":       "#0F172A",
        "bg_secondary":     "#1E293B",
        "bg_tertiary":      "#334155",
        "bg_elevated":      "#1E293B",

        # --- Text ---
        "text_primary":     "#E2E8F0",
        "text_secondary":   "#94A3B8",
        "text_tertiary":    "#64748B",
        "text_on_accent":   "#FFFFFF",

        # --- Border ---
        "border_default":   "#334155",
        "border_strong":    "#475569",
        "border_focus":     "#3B82F6",

        # --- Accent (Primary Blue) ---
        "accent":           "#3B82F6",
        "accent_hover":     "#60A5FA",
        "accent_pressed":   "#2563EB",
        "accent_subtle":    "#1E3A5F",
        "accent_text":      "#93C5FD",

        # --- Semantic: Success ---
        "success":          "#059669",
        "success_hover":    "#10B981",
        "success_pressed":  "#047857",

        # --- Semantic: Danger ---
        "danger":           "#DC2626",
        "danger_hover":     "#EF4444",
        "danger_pressed":   "#B91C1C",

        # --- Semantic: Warning ---
        "warning":          "#D97706",
        "warning_hover":    "#F59E0B",

        # --- Semantic: Secondary ---
        "secondary":        "#475569",
        "secondary_hover":  "#64748B",

        # --- Component-specific ---
        "tab_active_bg":    "#1E293B",
        "tab_inactive_bg":  "#0F172A",
        "tab_active_text":  "#60A5FA",
        "tab_border":       "#334155",
        "input_bg":         "#0F172A",
        "input_border":     "#334155",
        "input_focus":      "#3B82F6",
        "table_header_bg":  "#1E293B",
        "table_alt_row":    "#162032",
        "table_selected":   "#1E3A5F",
        "table_grid":       "#334155",
        "scrollbar_bg":     "#0F172A",
        "scrollbar_handle": "#475569",
        "group_border":     "#334155",
        "selection_bg":     "#1565C0",
        "selection_text":   "#FFFFFF",
        "disabled_bg":      "#1E293B",
        "disabled_text":    "#475569",

        # --- Button overrides (clock in / clock out) ---
        "btn_clock_in":         "#059669",
        "btn_clock_in_hover":   "#10B981",
        "btn_clock_in_pressed": "#047857",
        "btn_clock_out":        "#DC2626",
        "btn_clock_out_hover":  "#EF4444",
        "btn_clock_out_pressed":"#B91C1C",


        # --- Calendar-specific ---
        "cal_normal_bg":       "#1E293B",
        "cal_normal_text":     "#CBD5E1",
        "cal_holiday_bg":      "#2D1515",
        "cal_holiday_text":    "#F87171",
        "cal_saturday_bg":     "#0F1F3B",
        "cal_saturday_text":   "#60A5FA",
        "cal_today_bg":        "#1E3A8A",
        "cal_today_text":      "#93C5FD",
        "cal_today_border":    "#3B82F6",
        "cal_selected_bg":     "#3B82F6",
        "cal_selected_text":   "#FFFFFF",
        "cal_selected_border": "#2563EB",
        "cal_hover_normal":    "#334155",
        "cal_hover_holiday":   "#3D1C1C",
        "cal_hover_saturday":  "#1E3A5F",
        "cal_hover_today":     "#1D4ED8",
        "cal_header_sunday":   "#F87171",
        "cal_header_saturday": "#60A5FA",
    },
    "sepia": {
        # --- Surface / Background ---
        "bg_primary":       "#FDF6E3",
        "bg_secondary":     "#F5EDDA",
        "bg_tertiary":      "#EDE0C4",
        "bg_elevated":      "#FDF6E3",

        # --- Text ---
        "text_primary":     "#3D2B1F",
        "text_secondary":   "#7D5A4F",
        "text_tertiary":    "#A8826F",
        "text_on_accent":   "#FFFFFF",

        # --- Border ---
        "border_default":   "#D4BC8A",
        "border_strong":    "#C4A870",
        "border_focus":     "#B8860B",

        # --- Accent (Amber / Dark Goldenrod) ---
        "accent":           "#B8860B",
        "accent_hover":     "#DAA520",
        "accent_pressed":   "#8B6508",
        "accent_subtle":    "#FEF9C3",
        "accent_text":      "#7A5200",

        # --- Semantic: Success ---
        "success":          "#2E7D32",
        "success_hover":    "#388E3C",
        "success_pressed":  "#1B5E20",

        # --- Semantic: Danger ---
        "danger":           "#C62828",
        "danger_hover":     "#D32F2F",
        "danger_pressed":   "#B71C1C",

        # --- Semantic: Warning ---
        "warning":          "#F57F17",
        "warning_hover":    "#FF8F00",

        # --- Semantic: Secondary ---
        "secondary":        "#8D6E63",
        "secondary_hover":  "#6D4C41",

        # --- Component-specific ---
        "tab_active_bg":    "#FDF6E3",
        "tab_inactive_bg":  "#EDE0C4",
        "tab_active_text":  "#8B6508",
        "tab_border":       "#D4BC8A",
        "input_bg":         "#FFFBF0",
        "input_border":     "#D4BC8A",
        "input_focus":      "#B8860B",
        "table_header_bg":  "#EDE0C4",
        "table_alt_row":    "#F5EDDA",
        "table_selected":   "#FEF08A",
        "table_grid":       "#D4BC8A",
        "scrollbar_bg":     "#EDE0C4",
        "scrollbar_handle": "#C4A870",
        "group_border":     "#D4BC8A",
        "selection_bg":     "#FEF08A",
        "selection_text":   "#3D2B1F",
        "disabled_bg":      "#E8DFC8",
        "disabled_text":    "#A8926A",

        # --- Button overrides (clock in / clock out) ---
        "btn_clock_in":         "#5F7A3D",
        "btn_clock_in_hover":   "#4A6030",
        "btn_clock_in_pressed": "#3A4D25",
        "btn_clock_out":        "#A94040",
        "btn_clock_out_hover":  "#8B3030",
        "btn_clock_out_pressed":"#6F2424",


        # --- Calendar-specific ---
        "cal_normal_bg":       "#FDF6E3",
        "cal_normal_text":     "#5C3D2E",
        "cal_holiday_bg":      "#FEF2F2",
        "cal_holiday_text":    "#C62828",
        "cal_saturday_bg":     "#EFF6FF",
        "cal_saturday_text":   "#3B82F6",
        "cal_today_bg":        "#FEF9C3",
        "cal_today_text":      "#7A5200",
        "cal_today_border":    "#B8860B",
        "cal_selected_bg":     "#B8860B",
        "cal_selected_text":   "#FFFFFF",
        "cal_selected_border": "#8B6508",
        "cal_hover_normal":    "#EDE0C4",
        "cal_hover_holiday":   "#FEE2E2",
        "cal_hover_saturday":  "#DBEAFE",
        "cal_hover_today":     "#FEF08A",
        "cal_header_sunday":   "#C62828",
        "cal_header_saturday": "#3B82F6",
    },
    "green": {
        # --- Surface / Background ---
        "bg_primary":       "#0D1A0D",
        "bg_secondary":     "#132213",
        "bg_tertiary":      "#1E3A1E",
        "bg_elevated":      "#162B16",

        # --- Text ---
        "text_primary":     "#D4EAD4",
        "text_secondary":   "#7BA87B",
        "text_tertiary":    "#4D6E4D",
        "text_on_accent":   "#FFFFFF",

        # --- Border ---
        "border_default":   "#1E3A1E",
        "border_strong":    "#2D5A2D",
        "border_focus":     "#10B981",

        # --- Accent (Emerald) ---
        "accent":           "#10B981",
        "accent_hover":     "#34D399",
        "accent_pressed":   "#059669",
        "accent_subtle":    "#052E16",
        "accent_text":      "#6EE7B7",

        # --- Semantic: Success ---
        "success":          "#059669",
        "success_hover":    "#10B981",
        "success_pressed":  "#047857",

        # --- Semantic: Danger ---
        "danger":           "#DC2626",
        "danger_hover":     "#EF4444",
        "danger_pressed":   "#B91C1C",

        # --- Semantic: Warning ---
        "warning":          "#D97706",
        "warning_hover":    "#F59E0B",

        # --- Semantic: Secondary ---
        "secondary":        "#2D5A2D",
        "secondary_hover":  "#3D7A3D",

        # --- Component-specific ---
        "tab_active_bg":    "#162B16",
        "tab_inactive_bg":  "#0D1A0D",
        "tab_active_text":  "#34D399",
        "tab_border":       "#1E3A1E",
        "input_bg":         "#0D1A0D",
        "input_border":     "#1E3A1E",
        "input_focus":      "#10B981",
        "table_header_bg":  "#162B16",
        "table_alt_row":    "#0F1F0F",
        "table_selected":   "#052E16",
        "table_grid":       "#1E3A1E",
        "scrollbar_bg":     "#0D1A0D",
        "scrollbar_handle": "#2D5A2D",
        "group_border":     "#1E3A1E",
        "selection_bg":     "#065F46",
        "selection_text":   "#FFFFFF",
        "disabled_bg":      "#162B16",
        "disabled_text":    "#2D5A2D",

        # --- Button overrides (clock in / clock out) ---
        "btn_clock_in":         "#0891B2",
        "btn_clock_in_hover":   "#0E7490",
        "btn_clock_in_pressed": "#155E75",
        "btn_clock_out":        "#DC2626",
        "btn_clock_out_hover":  "#EF4444",
        "btn_clock_out_pressed":"#B91C1C",


        # --- Calendar-specific ---
        "cal_normal_bg":       "#162B16",
        "cal_normal_text":     "#9DBD9D",
        "cal_holiday_bg":      "#2D1515",
        "cal_holiday_text":    "#F87171",
        "cal_saturday_bg":     "#0D1A2D",
        "cal_saturday_text":   "#60A5FA",
        "cal_today_bg":        "#052E16",
        "cal_today_text":      "#6EE7B7",
        "cal_today_border":    "#10B981",
        "cal_selected_bg":     "#10B981",
        "cal_selected_text":   "#FFFFFF",
        "cal_selected_border": "#059669",
        "cal_hover_normal":    "#1E3A1E",
        "cal_hover_holiday":   "#3D1C1C",
        "cal_hover_saturday":  "#1E3A5F",
        "cal_hover_today":     "#065F46",
        "cal_header_sunday":   "#F87171",
        "cal_header_saturday": "#60A5FA",
    },
    "high_contrast": {
        # --- Surface / Background ---
        "bg_primary":       "#000000",
        "bg_secondary":     "#000000",
        "bg_tertiary":      "#1A1A1A",
        "bg_elevated":      "#0C0C0C",

        # --- Text ---
        "text_primary":     "#FFFFFF",
        "text_secondary":   "#FFFF00",
        "text_tertiary":    "#C0C0C0",
        "text_on_accent":   "#000000",

        # --- Border ---
        "border_default":   "#FFFFFF",
        "border_strong":    "#FFFF00",
        "border_focus":     "#1AEBFF",

        # --- Accent (Cyan) ---
        "accent":           "#1AEBFF",
        "accent_hover":     "#00D4E8",
        "accent_pressed":   "#00B8CC",
        "accent_subtle":    "#001A1F",
        "accent_text":      "#1AEBFF",

        # --- Semantic: Success ---
        "success":          "#00FF00",
        "success_hover":    "#00CC00",
        "success_pressed":  "#009900",

        # --- Semantic: Danger ---
        "danger":           "#FF4040",
        "danger_hover":     "#FF0000",
        "danger_pressed":   "#CC0000",

        # --- Semantic: Warning ---
        "warning":          "#FFFF00",
        "warning_hover":    "#E6E600",

        # --- Semantic: Secondary ---
        "secondary":        "#C0C0C0",
        "secondary_hover":  "#E0E0E0",

        # --- Component-specific ---
        "tab_active_bg":    "#000000",
        "tab_inactive_bg":  "#000000",
        "tab_active_text":  "#FFFF00",
        "tab_border":       "#FFFFFF",
        "input_bg":         "#000000",
        "input_border":     "#FFFFFF",
        "input_focus":      "#1AEBFF",
        "table_header_bg":  "#000000",
        "table_alt_row":    "#0D0D0D",
        "table_selected":   "#004080",
        "table_grid":       "#FFFFFF",
        "scrollbar_bg":     "#000000",
        "scrollbar_handle": "#FFFFFF",
        "group_border":     "#FFFFFF",
        "selection_bg":     "#0078D7",
        "selection_text":   "#FFFFFF",
        "disabled_bg":      "#1A1A1A",
        "disabled_text":    "#808080",

        # --- Button overrides (clock in / clock out) ---
        "btn_clock_in":         "#00FF00",
        "btn_clock_in_hover":   "#00CC00",
        "btn_clock_in_pressed": "#009900",
        "btn_clock_out":        "#FF4040",
        "btn_clock_out_hover":  "#FF0000",
        "btn_clock_out_pressed":"#CC0000",

        # --- Calendar-specific ---
        "cal_normal_bg":       "#000000",
        "cal_normal_text":     "#FFFFFF",
        "cal_holiday_bg":      "#1A0000",
        "cal_holiday_text":    "#FF6060",
        "cal_saturday_bg":     "#00001A",
        "cal_saturday_text":   "#6699FF",
        "cal_today_bg":        "#003300",
        "cal_today_text":      "#00FF00",
        "cal_today_border":    "#00FF00",
        "cal_selected_bg":     "#1AEBFF",
        "cal_selected_text":   "#000000",
        "cal_selected_border": "#00D4E8",
        "cal_hover_normal":    "#1A1A1A",
        "cal_hover_holiday":   "#2D0000",
        "cal_hover_saturday":  "#00002D",
        "cal_hover_today":     "#005500",
        "cal_header_sunday":   "#FF6060",
        "cal_header_saturday": "#6699FF",
    },
}


def get_theme_colors(theme: str) -> dict:
    """外部モジュール（calendar_widget 等）向けにカラー辞書を返す"""
    return THEME_COLORS.get(theme, THEME_COLORS["light"])


def get_stylesheet(theme: str) -> str:
    """light / dark テーマの QSS 文字列を返す"""
    c = get_theme_colors(theme)
    return _build_stylesheet(c)


def apply_theme(app: QApplication, theme: str) -> None:
    """アプリケーション全体にテーマを適用する"""
    app.setStyleSheet(get_stylesheet(theme))


def _build_stylesheet(c: dict) -> str:
    return f"""
    /* ==================== */
    /*   Global / Window    */
    /* ==================== */
    QMainWindow, QWidget {{
        background-color: {c['bg_secondary']};
        color: {c['text_primary']};
        font-size: 13px;
    }}

    /* ==================== */
    /*       Tab Bar        */
    /* ==================== */
    QTabWidget::pane {{
        border: 1px solid {c['tab_border']};
        background-color: {c['bg_elevated']};
        border-radius: 0 0 8px 8px;
    }}
    QTabBar::tab {{
        background-color: {c['tab_inactive_bg']};
        color: {c['text_secondary']};
        padding: 8px 20px;
        border: 1px solid {c['tab_border']};
        border-bottom: none;
        border-radius: 6px 6px 0 0;
        margin-right: 2px;
        font-size: 13px;
    }}
    QTabBar::tab:selected {{
        background-color: {c['tab_active_bg']};
        color: {c['tab_active_text']};
        font-weight: bold;
        border-bottom: 2px solid {c['accent']};
    }}
    QTabBar::tab:hover:!selected {{
        background-color: {c['bg_tertiary']};
    }}

    /* ==================== */
    /*    Push Buttons      */
    /* ==================== */
    QPushButton {{
        background-color: {c['accent']};
        color: {c['text_on_accent']};
        border: none;
        padding: 8px 20px;
        border-radius: 8px;
        font-size: 13px;
        font-weight: bold;
    }}
    QPushButton:hover {{
        background-color: {c['accent_hover']};
    }}
    QPushButton:pressed {{
        background-color: {c['accent_pressed']};
    }}
    QPushButton:disabled {{
        background-color: {c['disabled_bg']};
        color: {c['disabled_text']};
    }}

    /* -- Clock In -- */
    QPushButton#clock_in_btn {{
        background-color: {c['btn_clock_in']};
        font-size: 15px;
        padding: 10px 24px;
        border-radius: 10px;
    }}
    QPushButton#clock_in_btn:hover {{
        background-color: {c['btn_clock_in_hover']};
    }}
    QPushButton#clock_in_btn:pressed {{
        background-color: {c['btn_clock_in_pressed']};
    }}
    QPushButton#clock_in_btn:disabled {{
        background-color: {c['bg_secondary']};
        color: {c['text_tertiary']};
        border: 1px solid {c['border_default']};
        font-weight: normal;
    }}

    /* -- Clock Out -- */
    QPushButton#clock_out_btn {{
        background-color: {c['btn_clock_out']};
        font-size: 15px;
        padding: 10px 24px;
        border-radius: 10px;
    }}
    QPushButton#clock_out_btn:hover {{
        background-color: {c['btn_clock_out_hover']};
    }}
    QPushButton#clock_out_btn:pressed {{
        background-color: {c['btn_clock_out_pressed']};
    }}
    QPushButton#clock_out_btn:disabled {{
        background-color: {c['bg_secondary']};
        color: {c['text_tertiary']};
        border: 1px solid {c['border_default']};
        font-weight: normal;
    }}

    /* -- Info (Primary) -- */
    QPushButton#info_btn {{
        background-color: {c['accent']};
    }}
    QPushButton#info_btn:hover {{
        background-color: {c['accent_hover']};
    }}

    /* -- Secondary -- */
    QPushButton#secondary_btn {{
        background-color: {c['secondary']};
    }}
    QPushButton#secondary_btn:hover {{
        background-color: {c['secondary_hover']};
    }}

    /* -- Warning -- */
    QPushButton#warning_btn {{
        background-color: {c['warning']};
    }}
    QPushButton#warning_btn:hover {{
        background-color: {c['warning_hover']};
    }}
    QPushButton#warning_btn:disabled {{
        background-color: {c['bg_secondary']};
        color: {c['text_tertiary']};
        border: 1px solid {c['border_default']};
        font-weight: normal;
    }}

    /* -- Navigation (Calendar arrows) -- */
    QPushButton#nav_btn {{
        background-color: transparent;
        color: {c['text_secondary']};
        border: 1px solid {c['border_default']};
        border-radius: 6px;
        padding: 0px;
        font-size: 14px;
        min-width: 0px;
    }}
    QPushButton#nav_btn:hover {{
        background-color: {c['bg_tertiary']};
        border-color: {c['border_strong']};
    }}

    /* ==================== */
    /*    Text Inputs       */
    /* ==================== */
    QLineEdit, QTextEdit, QTimeEdit, QDateEdit {{
        background-color: {c['input_bg']};
        color: {c['text_primary']};
        border: 1px solid {c['input_border']};
        border-radius: 6px;
        padding: 6px 10px;
        font-size: 13px;
    }}
    QLineEdit:focus, QTextEdit:focus, QTimeEdit:focus, QDateEdit:focus {{
        border: 2px solid {c['input_focus']};
        padding: 5px 9px;
    }}

    /* ==================== */
    /*      ComboBox        */
    /* ==================== */
    QComboBox {{
        background-color: {c['input_bg']};
        color: {c['text_primary']};
        border: 1px solid {c['input_border']};
        border-radius: 6px;
        padding: 6px 28px 6px 10px;
        font-size: 13px;
    }}
    QComboBox:hover {{
        border: 1px solid {c['border_strong']};
    }}
    QComboBox QAbstractItemView {{
        background-color: {c['bg_elevated']};
        color: {c['text_primary']};
        selection-background-color: {c['selection_bg']};
        selection-color: {c['selection_text']};
        border: 1px solid {c['border_default']};
        border-radius: 6px;
        outline: 0px;
        padding: 4px;
    }}

    /* ==================== */
    /*      GroupBox        */
    /* ==================== */
    QGroupBox {{
        font-weight: bold;
        font-size: 13px;
        color: {c['text_primary']};
        border: 1px solid {c['group_border']};
        border-radius: 8px;
        margin-top: 12px;
        padding: 16px 12px 12px 12px;
    }}
    QGroupBox::title {{
        subcontrol-origin: margin;
        left: 12px;
        padding: 0 6px;
        color: {c['accent_text']};
    }}

    /* ==================== */
    /*       Labels         */
    /* ==================== */
    QLabel {{
        color: {c['text_primary']};
    }}
    QLabel#month_label {{
        color: {c['text_primary']};
        font-size: 18px;
        font-weight: bold;
        padding: 2px 12px;
    }}
    QLabel#date_badge {{
        background-color: {c['accent_subtle']};
        color: {c['accent_text']};
        border: 1px solid {c['border_focus']};
        border-radius: 10px;
        padding: 6px 16px;
        font-size: 14px;
        font-weight: bold;
    }}
    QRadioButton, QCheckBox {{
        color: {c['text_primary']};
        spacing: 8px;
        font-size: 13px;
    }}

    /* ==================== */
    /*     Table Widget     */
    /* ==================== */
    QTableWidget {{
        background-color: {c['bg_elevated']};
        alternate-background-color: {c['table_alt_row']};
        color: {c['text_primary']};
        border: 1px solid {c['border_default']};
        border-radius: 8px;
        gridline-color: {c['table_grid']};
        font-size: 13px;
    }}
    QTableWidget::item:selected {{
        background-color: {c['table_selected']};
        color: {c['text_primary']};
    }}
    QHeaderView::section {{
        background-color: {c['table_header_bg']};
        color: {c['text_primary']};
        border: none;
        border-bottom: 2px solid {c['accent']};
        padding: 6px 8px;
        font-weight: bold;
        font-size: 12px;
    }}

    /* ==================== */
    /*     List Widget      */
    /* ==================== */
    QListWidget {{
        background-color: {c['bg_elevated']};
        alternate-background-color: {c['table_alt_row']};
        color: {c['text_primary']};
        border: 1px solid {c['border_default']};
        border-radius: 8px;
        padding: 4px;
        font-size: 13px;
    }}
    QListWidget::item {{
        padding: 6px 10px;
        border-radius: 4px;
    }}
    QListWidget::item:hover {{
        background-color: {c['bg_tertiary']};
    }}
    QListWidget::item:selected {{
        background-color: {c['selection_bg']};
        color: {c['selection_text']};
    }}

    /* ==================== */
    /*     ScrollBar        */
    /* ==================== */
    QScrollBar:vertical {{
        background: {c['scrollbar_bg']};
        width: 8px;
        border-radius: 4px;
    }}
    QScrollBar::handle:vertical {{
        background: {c['scrollbar_handle']};
        border-radius: 4px;
        min-height: 30px;
    }}
    QScrollBar::handle:vertical:hover {{
        background: {c['border_strong']};
    }}
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
        height: 0px;
    }}
    QScrollBar:horizontal {{
        background: {c['scrollbar_bg']};
        height: 8px;
        border-radius: 4px;
    }}
    QScrollBar::handle:horizontal {{
        background: {c['scrollbar_handle']};
        border-radius: 4px;
        min-width: 30px;
    }}

    /* ==================== */
    /*     Dialog           */
    /* ==================== */
    QDialog {{
        background-color: {c['bg_secondary']};
        color: {c['text_primary']};
    }}
    QDialogButtonBox QPushButton {{
        min-width: 80px;
        padding: 8px 16px;
    }}

    /* ==================== */
    /*     ScrollArea       */
    /* ==================== */
    QScrollArea {{
        border: none;
        background-color: transparent;
    }}
    """
