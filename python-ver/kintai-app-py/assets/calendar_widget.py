"""カスタムカレンダーウィジェット"""
import calendar
from datetime import date, timedelta
from typing import Optional, Set

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QPushButton, QSizePolicy, QFrame
)
from PyQt5.QtCore import Qt, pyqtSignal, QDate
from PyQt5.QtGui import QFont, QColor, QPalette

from assets.theme_engine import get_theme_colors
from assets.timesheet_helpers import get_today

try:
    from assets.timesheet_helpers import get_holidays
except ImportError:
    def get_holidays(year: int, month: int) -> Set[date]:
        return set()


class CalendarCell(QLabel):
    """カレンダーの各日セル"""

    clicked = pyqtSignal(date)
    double_clicked = pyqtSignal(date)

    # クラスレベルのテーマ・カラー（CalendarWidget.set_theme() で一括更新）
    _theme: str = "light"
    _colors: dict = {}

    @classmethod
    def update_theme_colors(cls, theme: str) -> None:
        cls._theme = theme
        cls._colors = get_theme_colors(theme)

    def __init__(self, cell_date: Optional[date] = None, parent=None):
        super().__init__(parent)
        self._date: Optional[date] = cell_date
        self._is_selected = False
        self._is_today = False
        self._is_holiday = False
        self._is_weekend = False
        self._is_saturday = False
        self._is_hovered = False

        self.setAlignment(Qt.AlignCenter)
        self.setMinimumSize(40, 43)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.setCursor(Qt.PointingHandCursor if cell_date else Qt.ArrowCursor)

        if cell_date:
            self.setText(str(cell_date.day))
        else:
            self.setText("")

        self._update_style()

    def set_date(self, cell_date: Optional[date]) -> None:
        self._date = cell_date
        if cell_date:
            self.setText(str(cell_date.day))
            self.setCursor(Qt.PointingHandCursor)
        else:
            self.setText("")
            self.setCursor(Qt.ArrowCursor)

    def get_date(self) -> Optional[date]:
        return self._date

    def setSelected(self, selected: bool) -> None:
        self._is_selected = selected
        self._update_style()

    def setToday(self, is_today: bool) -> None:
        self._is_today = is_today
        self._update_style()

    def setHoliday(self, is_holiday: bool) -> None:
        self._is_holiday = is_holiday
        self._update_style()

    def setWeekend(self, is_weekend: bool) -> None:
        self._is_weekend = is_weekend
        self._update_style()

    def setSaturday(self, is_saturday: bool) -> None:
        self._is_saturday = is_saturday
        self._update_style()

    def _update_style(self) -> None:
        """状態に応じてスタイルを更新する（テーマ辞書参照）"""
        c = CalendarCell._colors
        if not c:
            c = get_theme_colors(CalendarCell._theme)

        border_radius = "6px"
        font_weight = "normal"

        # ── ベース背景色（優先度: selected > today > holiday/saturday > normal）──
        if self._is_selected:
            base_bg    = c['cal_selected_bg']
            border     = f"2px solid {c['cal_selected_border']}"
            text_color = c['cal_selected_text']
            font_weight = "bold"
        elif self._is_today:
            base_bg    = c['cal_today_bg']
            border     = f"2px solid {c['cal_today_border']}"
            text_color = c['cal_today_text']
            font_weight = "bold"
        elif self._is_holiday or self._is_weekend:
            base_bg    = c['cal_holiday_bg']
            border     = "1px solid transparent"
            text_color = c['cal_holiday_text']
        elif self._is_saturday:
            base_bg    = c['cal_saturday_bg']
            border     = "1px solid transparent"
            text_color = c['cal_saturday_text']
        else:
            base_bg    = c['cal_normal_bg']
            border     = "1px solid transparent"
            text_color = c['cal_normal_text']

        # ホバー時（選択中は変えない）
        if self._is_hovered and not self._is_selected and self._date:
            if self._is_today:
                base_bg = c['cal_hover_today']
            elif self._is_holiday or self._is_weekend:
                base_bg = c['cal_hover_holiday']
            elif self._is_saturday:
                base_bg = c['cal_hover_saturday']
            else:
                base_bg = c['cal_hover_normal']

        if not self._date:
            base_bg = "transparent"
            border = "1px solid transparent"
            text_color = "transparent"

        self.setStyleSheet(f"""
            QLabel {{
                background-color: {base_bg};
                color: {text_color};
                border: {border};
                border-radius: {border_radius};
                font-weight: {font_weight};
                font-size: 13px;
                padding: 2px;
            }}
        """)

    def mousePressEvent(self, event) -> None:
        if self._date and event.button() == Qt.LeftButton:
            self.clicked.emit(self._date)
        super().mousePressEvent(event)

    def mouseDoubleClickEvent(self, event) -> None:
        if self._date and event.button() == Qt.LeftButton:
            self.double_clicked.emit(self._date)
        super().mouseDoubleClickEvent(event)

    def enterEvent(self, event) -> None:
        self._is_hovered = True
        self._update_style()
        super().enterEvent(event)

    def leaveEvent(self, event) -> None:
        self._is_hovered = False
        self._update_style()
        super().leaveEvent(event)


class CalendarWidget(QWidget):
    """日本語カスタムカレンダーウィジェット"""

    date_selected = pyqtSignal(date)
    date_double_clicked = pyqtSignal(date)

    DAY_HEADERS = ["日", "月", "火", "水", "木", "金", "土"]

    def __init__(self, parent=None):
        super().__init__(parent)
        today = get_today()
        self._year = today.year
        self._month = today.month
        self._selected_date: Optional[date] = None
        self._cells: list = []
        self._day_header_labels: list = []

        self._init_ui()
        self._build_grid()

    def _init_ui(self) -> None:
        """UI を初期化する"""
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(4)
        main_layout.setContentsMargins(4, 4, 4, 4)

        # ── ヘッダー行（月ナビゲーション）──
        header_layout = QHBoxLayout()
        header_layout.setSpacing(4)

        self.prev_btn = QPushButton("◀")
        self.prev_btn.setFixedSize(30, 30)
        self.prev_btn.clicked.connect(self.prev_month)
        self.prev_btn.setObjectName("nav_btn")
        header_layout.addWidget(self.prev_btn)

        header_layout.addStretch()

        self.month_label = QLabel()
        self.month_label.setAlignment(Qt.AlignCenter)
        self.month_label.setObjectName("month_label")
        header_layout.addWidget(self.month_label)

        header_layout.addStretch()

        self.next_btn = QPushButton("▶")
        self.next_btn.setFixedSize(30, 30)
        self.next_btn.clicked.connect(self.next_month)
        self.next_btn.setObjectName("nav_btn")
        header_layout.addWidget(self.next_btn)

        main_layout.addLayout(header_layout)

        # ── 曜日ヘッダー行 ──
        day_header_layout = QGridLayout()
        day_header_layout.setSpacing(2)
        self._day_header_labels = []
        for col, day_name in enumerate(self.DAY_HEADERS):
            lbl = QLabel(day_name)
            lbl.setAlignment(Qt.AlignCenter)
            lbl.setMinimumWidth(40)
            font = lbl.font()
            font.setBold(True)
            lbl.setFont(font)
            self._day_header_labels.append(lbl)
            day_header_layout.addWidget(lbl, 0, col)
        self._update_day_header_colors(CalendarCell._theme)
        main_layout.addLayout(day_header_layout)

        # ── カレンダーグリッド ──
        self.grid_layout = QGridLayout()
        self.grid_layout.setSpacing(2)
        main_layout.addLayout(self.grid_layout)

        main_layout.addStretch()

    def _build_grid(self) -> None:
        """現在の月のカレンダーグリッドを構築する"""
        # 既存セルを削除
        for cell in self._cells:
            self.grid_layout.removeWidget(cell)
            cell.deleteLater()
        self._cells.clear()

        # 月ラベル更新
        self.month_label.setText(f"{self._year}年{self._month:02d}月")

        # 祝日取得
        holidays: Set[date] = get_holidays(self._year, self._month)

        today = get_today()

        # calendar.Calendar(firstweekday=6) で日曜始まりの週リストを取得
        # week[0]=日, week[1]=月, ..., week[6]=土  (0 はその月の日付なし)
        cal = calendar.Calendar(firstweekday=6).monthdayscalendar(self._year, self._month)
        # 月によって4〜6週になるけど、常に6行分確保して選択日ラベル位置を固定する
        while len(cal) < 6:
            cal.append([0, 0, 0, 0, 0, 0, 0])

        for row_idx, week in enumerate(cal):
            for col_idx, day_num in enumerate(week):
                if day_num == 0:
                    cell = CalendarCell(None)
                else:
                    cell_date = date(self._year, self._month, day_num)
                    cell = CalendarCell(cell_date)

                    # col_idx 0=日, 6=土
                    is_sunday = col_idx == 0
                    is_saturday = col_idx == 6
                    is_holiday = cell_date in holidays

                    cell.setToday(cell_date == today)
                    cell.setSelected(cell_date == self._selected_date)
                    cell.setHoliday(is_holiday or is_sunday)
                    cell.setSaturday(is_saturday)

                    cell.clicked.connect(self._on_cell_clicked)
                    cell.double_clicked.connect(self._on_cell_double_clicked)

                self.grid_layout.addWidget(cell, row_idx, col_idx)
                self._cells.append(cell)

    def _on_cell_clicked(self, d: date) -> None:
        """セルクリック時の処理"""
        old_selected = self._selected_date
        self._selected_date = d
        # 旧選択セルの状態を更新
        for cell in self._cells:
            cell_date = cell.get_date()
            if cell_date == old_selected or cell_date == d:
                cell.setSelected(cell_date == d)
        self.date_selected.emit(d)

    def _on_cell_double_clicked(self, d: date) -> None:
        """セルダブルクリック時の処理"""
        self._selected_date = d
        for cell in self._cells:
            cell.setSelected(cell.get_date() == d)
        self.date_double_clicked.emit(d)

    def set_month(self, year: int, month: int) -> None:
        """表示月を変更してグリッドを再構築する"""
        self._year = year
        self._month = month
        self._build_grid()

    def get_selected_date(self) -> Optional[date]:
        """選択中の日付を返す"""
        return self._selected_date

    def prev_month(self) -> None:
        """前の月に移動する"""
        if self._month == 1:
            self._year -= 1
            self._month = 12
        else:
            self._month -= 1
        self._build_grid()

    def next_month(self) -> None:
        """次の月に移動する"""
        if self._month == 12:
            self._year += 1
            self._month = 1
        else:
            self._month += 1
        self._build_grid()

    def _update_day_header_colors(self, theme: str) -> None:
        """曜日ヘッダーの文字色をテーマに合わせて更新する"""
        c = get_theme_colors(theme)
        for col, lbl in enumerate(self._day_header_labels):
            if col == 0:  # 日曜
                lbl.setStyleSheet(f"color: {c['cal_header_sunday']}; font-weight: bold;")
            elif col == 6:  # 土曜
                lbl.setStyleSheet(f"color: {c['cal_header_saturday']}; font-weight: bold;")
            else:
                lbl.setStyleSheet(f"color: {c['text_primary']}; font-weight: bold;")

    def set_theme(self, theme: str) -> None:
        """テーマを切り替えてグリッドを再描画する"""
        CalendarCell.update_theme_colors(theme)
        self._update_day_header_colors(theme)
        self._build_grid()
