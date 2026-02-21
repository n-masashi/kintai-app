import os
import sys
from pathlib import Path

from PyQt5.QtWidgets import QApplication, QMainWindow, QTabWidget, QWidget
from PyQt5.QtGui import QFont, QFontDatabase
from PyQt5.QtCore import Qt

from assets.app_logger import setup_logging, get_logger
from assets.config import Config
from assets.theme_engine import apply_theme
from assets.tabs.attendance_tab import AttendanceTab
from assets.tabs.settings_tab import SettingsTab
from assets.tabs.shift_type_tab import ShiftTypeTab
VER = "勤怠打刻 v3.0.0"

def _apply_titlebar_theme(hwnd: int, theme: str) -> None:
    if sys.platform != "win32":
        return
    try:
        import ctypes
        from assets.theme_engine import get_theme_colors

        c = get_theme_colors(theme)

        # bg_secondary の輝度でテキスト色を自動判定（暗い背景→白テキスト）
        DWMWA_USE_IMMERSIVE_DARK_MODE = 20
        hex_bg = c["bg_secondary"].lstrip("#")
        r0, g0, b0 = int(hex_bg[0:2], 16), int(hex_bg[2:4], 16), int(hex_bg[4:6], 16)
        is_dark = (0.2126 * r0 + 0.7152 * g0 + 0.0722 * b0) < 128
        dark_val = ctypes.c_int(1 if is_dark else 0)
        ctypes.windll.dwmapi.DwmSetWindowAttribute(
            hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE,
            ctypes.byref(dark_val), ctypes.sizeof(dark_val),
        )

        # カスタム背景色
        DWMWA_CAPTION_COLOR = 35
        hex_col = c["bg_secondary"].lstrip("#")
        r, g, b = int(hex_col[0:2], 16), int(hex_col[2:4], 16), int(hex_col[4:6], 16)
        colorref = ctypes.c_int((b << 16) | (g << 8) | r)
        ctypes.windll.dwmapi.DwmSetWindowAttribute(
            hwnd, DWMWA_CAPTION_COLOR,
            ctypes.byref(colorref), ctypes.sizeof(colorref),
        )
    except Exception:
        pass


class MainWindow(QMainWindow):
    """メインウィンドウ"""

    def __init__(self, config: Config, app: QApplication):
        super().__init__()
        self.config = config
        self.app = app

        self.setWindowTitle(VER)

        # タブウィジェット
        self.tab_widget = QTabWidget(self)
        self.setCentralWidget(self.tab_widget)

        # 打刻タブ
        self.attendance_tab = AttendanceTab(config=self.config, parent=self)
        self.tab_widget.addTab(self.attendance_tab, "打刻")

        # 設定タブ
        self.settings_tab = SettingsTab(config=self.config, main_window=self, parent=self)
        self.tab_widget.addTab(self.settings_tab, "設定")

        # 出勤形態タブ
        self.shift_type_tab = ShiftTypeTab(
            config=self.config,
            attendance_tab_ref=self.attendance_tab,
            parent=self
        )
        self.tab_widget.addTab(self.shift_type_tab, "出勤形態")

    def apply_theme(self) -> None:
        # 現在のテーマを適用
        apply_theme(self.app, self.config.theme)
        if hasattr(self.attendance_tab, "calendar") and self.attendance_tab.calendar:
            self.attendance_tab.calendar.set_theme(self.config.theme)
        _apply_titlebar_theme(int(self.winId()), self.config.theme)


def _setup_font(app: QApplication) -> None:
    """日本語フォントを設定する"""
    candidates = ["Noto Sans CJK JP", "Noto Sans JP", "IPAGothic", "Yu Gothic", "Meiryo", "MS Gothic"]
    db = QFontDatabase()
    available = db.families()
    for name in candidates:
        if name in available:
            app.setFont(QFont(name, 10))
            return
    # システムデフォルト
    font = app.font()
    font.setPointSize(10)
    app.setFont(font)


def main() -> None:
    setup_logging()
    log = get_logger("kintai.main")
    log.info("=== アプリ起動 ===")

    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)
    app.setApplicationName(VER)

    _setup_font(app)

    # 設定読込
    settings_path = Path("configs/settings.json")
    config = Config.load(str(settings_path))

    # テスト日付オーバーライド（settings.json の test_date が優先、環境変数で上書きもできる）
    if config.test_date and not os.environ.get("KINTAI_TEST_DATE"):
        os.environ["KINTAI_TEST_DATE"] = config.test_date

    # テーマ適用
    apply_theme(app, config.theme)

    # メインウィンドウ
    window = MainWindow(config=config, app=app)
    window.show()
    # 起動時にカレンダーへもテーマを反映
    window.apply_theme()
    window.resize(880, 480)

    sys.exit(app.exec_())

 
if __name__ == "__main__":
    main()
