"""遅刻理由入力ダイアログ"""
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QLineEdit,
    QDialogButtonBox
)
from PyQt5.QtCore import Qt


class LateReasonDialog(QDialog):
    """遅刻理由を入力させるダイアログ"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("遅刻理由入力")
        self.setMinimumWidth(350)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)

        label = QLabel("遅刻理由を入力してください：")
        layout.addWidget(label)

        self.reason_edit = QLineEdit()
        self.reason_edit.setPlaceholderText("例: 電車遅延、体調不良 など")
        layout.addWidget(self.reason_edit)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_reason(self) -> str:
        """入力された遅刻理由を返す"""
        return self.reason_edit.text().strip()
