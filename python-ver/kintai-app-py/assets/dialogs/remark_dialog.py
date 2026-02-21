"""汎用備考入力ダイアログ（振休・1日有給など）"""
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QLineEdit,
    QDialogButtonBox
)
from PyQt5.QtCore import Qt


class RemarkDialog(QDialog):
    """備考を入力させる汎用ダイアログ"""

    def __init__(self, title: str = "備考入力", parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setMinimumWidth(350)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)

        label = QLabel("備考を入力してください：")
        layout.addWidget(label)

        self.remark_edit = QLineEdit()
        self.remark_edit.setPlaceholderText("備考（任意）")
        layout.addWidget(self.remark_edit)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_remark(self) -> str:
        """入力された備考を返す"""
        return self.remark_edit.text().strip()
