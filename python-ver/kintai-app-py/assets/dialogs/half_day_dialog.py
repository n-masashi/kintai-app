"""0.5日有給 時刻・備考入力ダイアログ"""
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QFormLayout, QLabel,
    QTimeEdit, QLineEdit, QDialogButtonBox, QWidget
)
from PyQt5.QtCore import Qt, QTime


class HalfDayDialog(QDialog):
    """0.5日有給の始業・終業時刻と備考を入力させるダイアログ"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("0.5日有給 時刻入力")
        self.setMinimumWidth(350)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)

        title_label = QLabel("0.5日有給の時刻と備考を入力してください。")
        layout.addWidget(title_label)

        form = QFormLayout()
        form.setSpacing(8)

        self.start_time = QTimeEdit()
        self.start_time.setTime(QTime(9, 0))
        self.start_time.setDisplayFormat("HH:mm")
        form.addRow("始業時刻：", self.start_time)

        self.end_time = QTimeEdit()
        self.end_time.setTime(QTime(13, 0))
        self.end_time.setDisplayFormat("HH:mm")
        form.addRow("終業時刻：", self.end_time)

        self.remark_edit = QLineEdit()
        self.remark_edit.setPlaceholderText("備考（任意）")
        form.addRow("備考：", self.remark_edit)

        layout.addLayout(form)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_start_time(self) -> QTime:
        """始業時刻を返す"""
        return self.start_time.time()

    def get_end_time(self) -> QTime:
        """終業時刻を返す"""
        return self.end_time.time()

    def get_remark(self) -> str:
        """備考を返す"""
        return self.remark_edit.text().strip()
