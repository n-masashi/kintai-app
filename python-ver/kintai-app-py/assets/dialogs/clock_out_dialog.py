"""退勤情報入力ダイアログ"""
from datetime import date, timedelta
from typing import List, Dict

try:
    from assets.timesheet_helpers import get_today
except ImportError:
    def get_today():
        return date.today()

from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout, QLabel,
    QDateEdit, QComboBox, QLineEdit, QDialogButtonBox,
    QRadioButton, QButtonGroup, QGroupBox
)
from PyQt5.QtCore import Qt, QDate


class ClockOutDialog(QDialog):
    """退勤時に次回出勤日・シフト・メンション先・コメントを入力させるダイアログ"""

    def __init__(
        self,
        shift_types: List[str] = None,
        managers: List[Dict[str, str]] = None,
        is_night: bool = False,
        parent=None
    ):
        super().__init__(parent)
        self.setWindowTitle("退勤情報入力")
        self.setMinimumWidth(400)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        shift_types = shift_types or []
        managers = managers or []

        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)

        title_label = QLabel("退勤情報を入力してください。")
        layout.addWidget(title_label)

        form = QFormLayout()
        form.setSpacing(8)

        # 次回出勤日
        tomorrow = get_today() + timedelta(days=1)
        self.next_workday_edit = QDateEdit()
        self.next_workday_edit.setCalendarPopup(True)
        self.next_workday_edit.setDate(QDate(tomorrow.year, tomorrow.month, tomorrow.day))
        self.next_workday_edit.setDisplayFormat("yyyy/MM/dd")
        form.addRow("次回出勤日：", self.next_workday_edit)

        # 次回シフト
        self.next_shift_combo = QComboBox()
        self.next_shift_combo.addItems(shift_types)
        form.addRow("次回シフト：", self.next_shift_combo)

        # メンション先
        self.mention_combo = QComboBox()
        self.mention_combo.addItem("（なし）")
        self.mention_combo.addItem("@All管理職")
        for m in managers:
            name = m.get("name", "")
            if name:
                self.mention_combo.addItem(name)
        form.addRow("メンション先：", self.mention_combo)

        # 次回勤務形態
        next_work_mode_row = QHBoxLayout()
        next_work_mode_row.setSpacing(12)
        self._next_work_mode_group = QButtonGroup(self)
        self.next_remote_radio = QRadioButton("リモート")
        self.next_office_radio = QRadioButton("出社")
        self.next_remote_radio.setChecked(True)
        self._next_work_mode_group.addButton(self.next_remote_radio)
        self._next_work_mode_group.addButton(self.next_office_radio)
        next_work_mode_row.addWidget(self.next_remote_radio)
        next_work_mode_row.addWidget(self.next_office_radio)
        next_work_mode_row.addStretch()
        form.addRow("次回勤務形態：", next_work_mode_row)

        # コメント
        self.comment_edit = QLineEdit()
        self.comment_edit.setPlaceholderText("コメント（任意）")
        form.addRow("コメント：", self.comment_edit)

        layout.addLayout(form)

        # 日跨ぎ選択
        if is_night:
            cross_day_box = QGroupBox("退勤タイプ（深夜シフト）")
            normal_label = "通常退勤（翌日退勤）"
            cross_label = "日を跨ぐ退勤（翌々日退勤）"
        else:
            cross_day_box = QGroupBox("退勤タイプ")
            normal_label = "通常退勤"
            cross_label = "日を跨ぐ退勤(翌日退勤）"
        cross_day_layout = QHBoxLayout(cross_day_box)
        cross_day_layout.setSpacing(16)
        self._cross_day_group = QButtonGroup(self)
        self.normal_radio = QRadioButton(normal_label)
        self.cross_day_radio = QRadioButton(cross_label)
        self.normal_radio.setChecked(True)
        self._cross_day_group.addButton(self.normal_radio)
        self._cross_day_group.addButton(self.cross_day_radio)
        cross_day_layout.addWidget(self.normal_radio)
        cross_day_layout.addWidget(self.cross_day_radio)
        layout.addWidget(cross_day_box)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_next_workday(self) -> date:
        """次回出勤日を返す"""
        q = self.next_workday_edit.date()
        return date(q.year(), q.month(), q.day())

    def get_next_shift(self) -> str:
        """次回シフトを返す"""
        return self.next_shift_combo.currentText()

    def get_mention(self) -> str:
        """メンション先の名前を返す（なし の場合は空文字）"""
        text = self.mention_combo.currentText()
        return "" if text == "（なし）" else text

    def get_comment(self) -> str:
        """コメントを返す。メンション指定あり＆コメント空の場合は「-」を補完する"""
        comment = self.comment_edit.text().strip()
        if not comment and self.get_mention():
            return "-"
        return comment

    def get_next_work_mode(self) -> str:
        """次回勤務形態を返す"""
        return "リモート" if self.next_remote_radio.isChecked() else "出社"

    def get_is_cross_day(self) -> bool:
        """日跨ぎ退勤かどうかを返す"""
        return self.cross_day_radio.isChecked()
