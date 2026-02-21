"""設定タブ"""
from pathlib import Path

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFormLayout, QGroupBox,
    QPushButton, QLineEdit, QComboBox, QLabel,
    QTableWidget, QTableWidgetItem, QHeaderView, QScrollArea,
    QFileDialog, QMessageBox, QFrame
)

# テーマ名 ↔ 内部キーのマッピング
_THEME_LABELS = ["ライト", "ダーク", "グリーン", "セピア", "ハイコントラスト"]
_THEME_TO_KEY  = {"ライト": "light", "ダーク": "dark", "グリーン": "green", "セピア": "sepia", "ハイコントラスト": "high_contrast"}
_KEY_TO_THEME  = {v: k for k, v in _THEME_TO_KEY.items()}
from PyQt5.QtCore import Qt


class SettingsTab(QWidget):
    """設定タブ"""

    def __init__(self, config, main_window=None, parent=None):
        super().__init__(parent)
        self.config = config
        self.main_window = main_window
        self._init_ui()
        self.load_from_config()

    def _init_ui(self) -> None:
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)

        # スクロールエリア
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        outer.addWidget(scroll)

        container = QWidget()
        scroll.setWidget(container)

        main_layout = QVBoxLayout(container)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(12, 12, 12, 12)

        # ── ユーザー情報 ──
        user_group = QGroupBox("ユーザー情報")
        user_form = QFormLayout(user_group)
        user_form.setSpacing(6)
        self.ad_name_edit = QLineEdit()
        self.display_name_edit = QLineEdit()
        self.teams_id_edit = QLineEdit()
        self.teams_id_edit.setPlaceholderText("TeamsPrincipalID（例: yamada@example.com）")
        self.shift_display_name_edit = QLineEdit()
        self.shift_display_name_edit.setPlaceholderText("例: 山田（シフト表に記載の名前）")
        self.timesheet_display_name_edit = QLineEdit()
        self.timesheet_display_name_edit.setPlaceholderText("例: 山田 （タイムシートファイル名「YYYYMM山田.xlsx」に含まれる名前）")
        user_form.addRow("ADユーザ名：", self.ad_name_edit)
        user_form.addRow("フルネーム：", self.display_name_edit)
        user_form.addRow("Teams ID：", self.teams_id_edit)
        user_form.addRow("シフト表上の名前表記：", self.shift_display_name_edit)
        user_form.addRow("タイムシートの名前表記：", self.timesheet_display_name_edit)
        main_layout.addWidget(user_group)

        # ── Teams Webhook ──
        webhook_group = QGroupBox("Teams Webhook")
        webhook_layout = QVBoxLayout(webhook_group)
        wh_row = QHBoxLayout()
        self.webhook_edit = QLineEdit()
        self.webhook_edit.setEchoMode(QLineEdit.Password)
        self.webhook_edit.setPlaceholderText("Webhook URL")
        wh_row.addWidget(self.webhook_edit)
        self.show_webhook_btn = QPushButton("表示")
        self.show_webhook_btn.setFixedWidth(80)
        self.show_webhook_btn.setCheckable(True)
        self.show_webhook_btn.toggled.connect(self._toggle_webhook_visibility)
        wh_row.addWidget(self.show_webhook_btn)
        webhook_layout.addLayout(wh_row)
        main_layout.addWidget(webhook_group)

        # ── フォルダパス ──
        folder_group = QGroupBox("フォルダパス")
        folder_form = QFormLayout(folder_group)
        folder_form.setSpacing(6)

        ts_row = QHBoxLayout()
        self.timesheet_folder_edit = QLineEdit()
        ts_row.addWidget(self.timesheet_folder_edit)
        ts_browse = QPushButton("参照")
        ts_browse.setFixedWidth(80)
        ts_browse.clicked.connect(lambda: self._browse_folder(self.timesheet_folder_edit))
        ts_row.addWidget(ts_browse)
        folder_form.addRow("タイムシートフォルダ：", ts_row)

        out_row = QHBoxLayout()
        self.output_folder_edit = QLineEdit()
        out_row.addWidget(self.output_folder_edit)
        out_browse = QPushButton("参照")
        out_browse.setFixedWidth(80)
        out_browse.clicked.connect(lambda: self._browse_folder(self.output_folder_edit))
        out_row.addWidget(out_browse)
        folder_form.addRow("CSV出力先：", out_row)

        main_layout.addWidget(folder_group)

        # ── 管理職 ──
        mgr_group = QGroupBox("管理職")
        mgr_layout = QVBoxLayout(mgr_group)
        self.managers_table = QTableWidget(0, 2)
        self.managers_table.setHorizontalHeaderLabels(["名前", "Teams ID"])
        self.managers_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.managers_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.managers_table.setFixedHeight(180)
        self.managers_table.setAlternatingRowColors(True)
        self.managers_table.verticalHeader().setVisible(False)
        mgr_layout.addWidget(self.managers_table)
        mgr_btn_row = QHBoxLayout()
        add_mgr_btn = QPushButton("行を追加")
        add_mgr_btn.clicked.connect(self._add_manager_row)
        del_mgr_btn = QPushButton("行を削除")
        del_mgr_btn.clicked.connect(self._delete_manager_row)
        mgr_btn_row.addWidget(add_mgr_btn)
        mgr_btn_row.addWidget(del_mgr_btn)
        mgr_btn_row.addStretch()
        mgr_layout.addLayout(mgr_btn_row)
        main_layout.addWidget(mgr_group)

        # ── プロキシ ──
        proxy_group = QGroupBox("プロキシ")
        proxy_form = QFormLayout(proxy_group)
        proxy_sh_row = QHBoxLayout()
        self.proxy_sh_edit = QLineEdit()
        self.proxy_sh_edit.setPlaceholderText("proxy.sh のパス（空欄でプロキシなし）")
        proxy_sh_row.addWidget(self.proxy_sh_edit)
        proxy_sh_browse = QPushButton("参照")
        proxy_sh_browse.setFixedWidth(80)
        proxy_sh_browse.clicked.connect(self._browse_proxy_sh)
        proxy_sh_row.addWidget(proxy_sh_browse)
        proxy_form.addRow("proxy.shパス：", proxy_sh_row)
        main_layout.addWidget(proxy_group)

        # ── テーマ ──
        theme_group = QGroupBox("テーマ")
        theme_layout = QHBoxLayout(theme_group)
        self.theme_combo = QComboBox()
        self.theme_combo.addItems(_THEME_LABELS)
        self.theme_combo.currentIndexChanged.connect(self._on_theme_changed)
        theme_layout.addWidget(self.theme_combo)
        theme_layout.addStretch()
        main_layout.addWidget(theme_group)

        # ── 保存ボタン ──
        save_btn = QPushButton("設定を保存")
        save_btn.setMinimumHeight(38)
        save_btn.clicked.connect(self.save_settings)
        main_layout.addWidget(save_btn)
        main_layout.addStretch()

    # ─────────────── ヘルパー ───────────────

    def _toggle_webhook_visibility(self, checked: bool) -> None:
        if checked:
            self.webhook_edit.setEchoMode(QLineEdit.Normal)
            self.show_webhook_btn.setText("非表示")
        else:
            self.webhook_edit.setEchoMode(QLineEdit.Password)
            self.show_webhook_btn.setText("表示")

    def _browse_folder(self, line_edit: QLineEdit) -> None:
        folder = QFileDialog.getExistingDirectory(self, "フォルダを選択", line_edit.text())
        if folder:
            line_edit.setText(folder)

    def _browse_proxy_sh(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "proxy.sh を選択", self.proxy_sh_edit.text(), "Shell Script (*.sh);;All Files (*)")
        if path:
            self.proxy_sh_edit.setText(path)

    def _add_manager_row(self) -> None:
        row = self.managers_table.rowCount()
        self.managers_table.insertRow(row)
        self.managers_table.setItem(row, 0, QTableWidgetItem(""))
        self.managers_table.setItem(row, 1, QTableWidgetItem(""))

    def _delete_manager_row(self) -> None:
        row = self.managers_table.currentRow()
        if row >= 0:
            self.managers_table.removeRow(row)

    def _on_theme_changed(self) -> None:
        theme = _THEME_TO_KEY.get(self.theme_combo.currentText(), "light")
        if self.config:
            self.config.theme = theme
        if self.main_window:
            self.main_window.apply_theme()

    # ─────────────── データ読み書き ───────────────

    def load_from_config(self) -> None:
        """設定をウィジェットに反映する"""
        if not self.config:
            return
        self.ad_name_edit.setText(self.config.ad_name or "")
        self.display_name_edit.setText(self.config.display_name or "")
        self.teams_id_edit.setText(self.config.teams_user_id or "")
        self.shift_display_name_edit.setText(getattr(self.config, "shift_display_name", "") or "")
        self.timesheet_display_name_edit.setText(getattr(self.config, "timesheet_display_name", "") or "")
        self.webhook_edit.setText(self.config.webhook_url or "")
        self.timesheet_folder_edit.setText(self.config.timesheet_folder or "")
        self.output_folder_edit.setText(self.config.output_folder or "attendance_data")
        self.proxy_sh_edit.setText(getattr(self.config, "proxy_sh", "") or "")

        label = _KEY_TO_THEME.get(self.config.theme, "ライト")
        idx = self.theme_combo.findText(label)
        self.theme_combo.setCurrentIndex(idx if idx >= 0 else 0)

        # 管理職テーブル
        self.managers_table.setRowCount(0)
        for mgr in (self.config.managers or []):
            row = self.managers_table.rowCount()
            self.managers_table.insertRow(row)
            self.managers_table.setItem(row, 0, QTableWidgetItem(mgr.get("name", "")))
            self.managers_table.setItem(row, 1, QTableWidgetItem(mgr.get("teams_id", "")))

    def save_settings(self) -> None:
        """ウィジェットの値を設定に保存する"""
        if not self.config:
            return
        self.config.ad_name = self.ad_name_edit.text().strip()
        self.config.display_name = self.display_name_edit.text().strip()
        self.config.teams_user_id = self.teams_id_edit.text().strip()
        self.config.shift_display_name = self.shift_display_name_edit.text().strip()
        self.config.timesheet_display_name = self.timesheet_display_name_edit.text().strip()
        self.config.webhook_url = self.webhook_edit.text().strip()
        self.config.timesheet_folder = self.timesheet_folder_edit.text().strip()
        self.config.output_folder = self.output_folder_edit.text().strip()
        self.config.proxy_sh = self.proxy_sh_edit.text().strip()
        self.config.theme = _THEME_TO_KEY.get(self.theme_combo.currentText(), "light")

        # 管理職
        managers = []
        for row in range(self.managers_table.rowCount()):
            name_item = self.managers_table.item(row, 0)
            id_item = self.managers_table.item(row, 1)
            name = name_item.text().strip() if name_item else ""
            teams_id = id_item.text().strip() if id_item else ""
            if name or teams_id:
                managers.append({"name": name, "teams_id": teams_id})
        self.config.managers = managers

        self.config.save("configs/settings.json")

        QMessageBox.information(self, "保存完了", "設定を保存しました。")
