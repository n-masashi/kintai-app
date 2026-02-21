"""出勤形態タブ"""
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QListWidget, QLineEdit, QPushButton, QMessageBox
)
from PyQt5.QtCore import Qt


class ShiftTypeTab(QWidget):
    """出勤形態を管理するタブ"""

    def __init__(self, config, attendance_tab_ref=None, parent=None):
        super().__init__(parent)
        self.config = config
        self.attendance_tab_ref = attendance_tab_ref
        self._init_ui()
        self._load_shifts()

    def _init_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setSpacing(8)
        layout.setContentsMargins(12, 12, 12, 12)

        desc_label = QLabel("出勤形態を管理します。追加・削除した変更は即座に設定ファイルに保存されます。\n⚠️ 新しい出勤形態を追加した場合、必ずtimesheet_constants.py, timesheet_actions.pyに処理内容を追記すること")
        desc_label.setWordWrap(True)
        layout.addWidget(desc_label)

        self.shift_list = QListWidget()
        self.shift_list.setAlternatingRowColors(True)
        layout.addWidget(self.shift_list)

        # 入力行
        input_row = QHBoxLayout()
        self.new_shift_edit = QLineEdit()
        self.new_shift_edit.setPlaceholderText("新しい出勤形態を入力…")
        self.new_shift_edit.returnPressed.connect(self.add_shift)
        input_row.addWidget(self.new_shift_edit)

        add_btn = QPushButton("追加")
        add_btn.clicked.connect(self.add_shift)
        input_row.addWidget(add_btn)

        del_btn = QPushButton("削除")
        del_btn.clicked.connect(self.remove_shift)
        input_row.addWidget(del_btn)

        layout.addLayout(input_row)

    def _load_shifts(self) -> None:
        """設定から出勤形態リストを読み込む"""
        self.shift_list.clear()
        if self.config:
            for shift in (self.config.shift_types or []):
                self.shift_list.addItem(shift)

    def add_shift(self) -> None:
        """新しい出勤形態を追加する"""
        name = self.new_shift_edit.text().strip()
        if not name:
            return
        if self.config and name in self.config.shift_types:
            QMessageBox.warning(self, "重複", f"'{name}' はすでに存在します。")
            return
        self.shift_list.addItem(name)
        if self.config:
            self.config.shift_types.append(name)
        self.new_shift_edit.clear()
        self._sync()

    def remove_shift(self) -> None:
        """選択中の出勤形態を削除する"""
        current = self.shift_list.currentItem()
        if not current:
            QMessageBox.information(self, "選択なし", "削除する出勤形態を選択してください。")
            return
        name = current.text()
        row = self.shift_list.row(current)
        self.shift_list.takeItem(row)
        if self.config and name in self.config.shift_types:
            self.config.shift_types.remove(name)
        self._sync()

    def _sync(self) -> None:
        """設定ファイルに保存し、打刻タブのコンボボックスを更新する"""
        if self.config:
            self.config.save("configs/settings.json")
        if self.attendance_tab_ref:
            shift_types = self.config.shift_types if self.config else []
            self.attendance_tab_ref.update_shift_types(shift_types)
