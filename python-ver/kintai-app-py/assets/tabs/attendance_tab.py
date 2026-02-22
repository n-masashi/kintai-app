"""打刻タブ"""
import random
from datetime import date, datetime
from pathlib import Path
from typing import List, Optional

try:
    from assets.timesheet_helpers import get_today
except ImportError:
    def get_today():
        return date.today()

_WEEKDAY_JA = ["月", "火", "水", "木", "金", "土", "日"]


def _format_date_badge(d: date) -> str:
    """選択日バッジ用フォーマット: '選択日：YYYY年M月D日（曜）'"""
    return f"選択日：{d.year}年{d.month}月{d.day}日（{_WEEKDAY_JA[d.weekday()]}）"

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGroupBox,
    QPushButton, QComboBox, QRadioButton, QCheckBox,
    QLabel, QTextEdit, QButtonGroup, QSizePolicy, QFrame,
    QMessageBox, QApplication, QDialog
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QPainter, QFont, QPixmap


class LoadingOverlay(QWidget):
    """処理中オーバーレイ: 親ウィンドウ全体を半透明で暗くして「処理中...」を表示する"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAttribute(Qt.WA_TransparentForMouseEvents, False)

    def paintEvent(self, event):
        p = QPainter(self)
        p.fillRect(self.rect(), QColor(0, 0, 0, 150))
        p.setPen(QColor(255, 255, 255))
        font = QFont()
        font.setPointSize(16)
        font.setBold(True)
        p.setFont(font)
        p.drawText(self.rect(), Qt.AlignCenter, "処理中...")
        p.end()

try:
    from assets.calendar_widget import CalendarWidget
except ImportError:
    CalendarWidget = None

try:
    import assets.timesheet_actions as ta
    from assets.timesheet_actions import (
        TimesheetNotFoundError, TimesheetLockedError, TimesheetWriteError,
        UnknownShiftTypeError
    )
    from assets.timesheet_constants import REALTIME_SHIFTS
except ImportError:
    ta = None
    TimesheetNotFoundError = None
    TimesheetLockedError = None
    TimesheetWriteError = None
    UnknownShiftTypeError = None
    REALTIME_SHIFTS = []

_BATCH_ALLOWED_SHIFTS = set(REALTIME_SHIFTS) | {"シフト休", "1.0日有給"}


class AttendanceTab(QWidget):
    """打刻タブ"""

    def __init__(self, config, parent=None):
        super().__init__(parent)
        self.config = config
        self._batch_dates: List[date] = []
        self._init_ui()

    def _init_ui(self) -> None:
        # ── 最外ペイン: 左右分割 ──
        root_layout = QHBoxLayout(self)
        root_layout.setSpacing(12)
        root_layout.setContentsMargins(12, 12, 12, 12)

        # ════════════════════════════
        # 左ペイン: カレンダー
        # ════════════════════════════
        cal_group = QWidget()
        cal_layout = QVBoxLayout(cal_group)
        cal_layout.setSpacing(4)

        if CalendarWidget:
            self.calendar = CalendarWidget()
            self.calendar.date_selected.connect(self._on_date_selected)
            self.calendar.date_double_clicked.connect(self._on_date_double_clicked)
            cal_layout.addWidget(self.calendar)
        else:
            self.calendar = None
            cal_layout.addWidget(QLabel("カレンダー読み込みエラー"))

        self.selected_date_label = QLabel(_format_date_badge(get_today()))
        self.selected_date_label.setAlignment(Qt.AlignCenter)
        self.selected_date_label.setObjectName("date_badge")
        cal_layout.addStretch()
        cal_layout.addWidget(self.selected_date_label)
        cal_layout.addSpacing(10)

        root_layout.addWidget(cal_group, stretch=1)

        # ════════════════════════════
        # 右ペイン: コントロール群
        # ════════════════════════════
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setSpacing(8)
        right_layout.setContentsMargins(0, 0, 0, 0)

        # ── 打刻コントロール ──
        ctrl_group = QWidget()
        ctrl_layout = QVBoxLayout(ctrl_group)
        ctrl_layout.setSpacing(8)

        # 出勤形態 + Timesheet Check ボタン
        row1 = QHBoxLayout()
        row1.addWidget(QLabel("出勤形態："))
        self.shift_combo = QComboBox()
        self.shift_combo.addItems(self.config.shift_types if self.config else [])
        self.shift_combo.setCurrentIndex(-1)  # デフォルトは空選択
        self.shift_combo.setMinimumWidth(130)
        self.shift_combo.currentTextChanged.connect(self._on_shift_changed)
        row1.addWidget(self.shift_combo)
        row1.addSpacing(8)
        self.timesheet_check_btn = QPushButton("Timesheet Check")
        self.timesheet_check_btn.setObjectName("secondary_btn")
        self.timesheet_check_btn.clicked.connect(self._on_timesheet_check)
        row1.addWidget(self.timesheet_check_btn, stretch=1)
        ctrl_layout.addLayout(row1)

        # 出勤形式 & オプション（横並び）
        style_opt_row = QHBoxLayout()
        style_opt_row.setSpacing(12)

        # 出勤形式
        style_box = QGroupBox("出勤形式")
        style_box_layout = QVBoxLayout(style_box)
        style_box_layout.setSpacing(2)
        self.remote_radio = QRadioButton("リモート")
        self.office_radio = QRadioButton("出社")
        self.remote_radio.setChecked(True)
        work_style_group = QButtonGroup(self)
        work_style_group.addButton(self.remote_radio)
        work_style_group.addButton(self.office_radio)
        style_box_layout.addWidget(self.remote_radio)
        style_box_layout.addWidget(self.office_radio)
        style_opt_row.addWidget(style_box)

        # オプション
        opt_box = QGroupBox("オプション")
        opt_box_layout = QVBoxLayout(opt_box)
        opt_box_layout.setSpacing(2)
        self.no_post_check = QCheckBox("TeamsPostなし")
        self.assumed_check = QCheckBox("想定入力")
        self.assumed_check.stateChanged.connect(
            lambda: self._on_shift_changed(self.shift_combo.currentText())
        )
        opt_box_layout.addWidget(self.no_post_check)
        opt_box_layout.addWidget(self.assumed_check)
        style_opt_row.addWidget(opt_box)

        ctrl_layout.addLayout(style_opt_row)

        # 出勤・退勤ボタン
        ctrl_layout.addSpacing(12)
        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)
        self.clock_in_btn = QPushButton("出  勤")
        self.clock_in_btn.setObjectName("clock_in_btn")
        self.clock_in_btn.setMinimumHeight(50)
        self.clock_in_btn.clicked.connect(self.on_clock_in)

        self.clock_out_btn = QPushButton("退  勤")
        self.clock_out_btn.setObjectName("clock_out_btn")
        self.clock_out_btn.setMinimumHeight(50)
        self.clock_out_btn.clicked.connect(self.on_clock_out)

        btn_row.addWidget(self.clock_in_btn, stretch=1)
        btn_row.addWidget(self.clock_out_btn, stretch=1)
        ctrl_layout.addLayout(btn_row)

        right_layout.addWidget(ctrl_group)
        right_layout.addStretch()

        # ── 一括記入 ──
        batch_group = QGroupBox("一括記入")
        batch_layout = QVBoxLayout(batch_group)
        batch_layout.setSpacing(6)

        self.batch_error_label = QLabel()
        self.batch_error_label.setWordWrap(True)
        self.batch_error_label.setStyleSheet("color: #EF4444; font-size: 12px;")
        self.batch_error_label.setVisible(False)
        batch_layout.addWidget(self.batch_error_label)

        self.batch_dates_edit = QTextEdit()
        self.batch_dates_edit.setReadOnly(True)
        self.batch_dates_edit.setMaximumHeight(70)
        self.batch_dates_edit.setPlaceholderText("ダブルクリックまたは「追加」で日付を追加…")
        batch_layout.addWidget(self.batch_dates_edit)

        batch_btn_row = QHBoxLayout()
        batch_btn_row.setSpacing(6)
        add_btn = QPushButton("追加")
        add_btn.setObjectName("info_btn")
        add_btn.clicked.connect(self._add_to_batch)
        clear_btn = QPushButton("クリア")
        clear_btn.setObjectName("secondary_btn")
        clear_btn.clicked.connect(self._clear_batch)
        self.batch_write_btn = QPushButton("一括記入")
        self.batch_write_btn.setObjectName("warning_btn")
        self.batch_write_btn.clicked.connect(self.on_batch_write)
        self.batch_write_btn.setEnabled(False)
        batch_btn_row.addWidget(add_btn)
        batch_btn_row.addWidget(clear_btn)
        batch_btn_row.addWidget(self.batch_write_btn)
        batch_layout.addLayout(batch_btn_row)

        right_layout.addWidget(batch_group)

        root_layout.addWidget(right_widget, stretch=1)

        self._on_shift_changed("")  # 全ウィジェット生成後に初期状態を設定

    # ─────────────── スロット ───────────────

    def _on_shift_changed(self, shift: str) -> None:
        """シフト選択変更時にボタンのラベルと有効状態を更新する"""
        is_realtime = shift in REALTIME_SHIFTS
        self.remote_radio.setEnabled(is_realtime)
        self.office_radio.setEnabled(is_realtime)

        if not self._batch_dates:
            if is_realtime:
                is_assumed = self.assumed_check.isChecked()
                self.clock_in_btn.setEnabled(True)
                self.clock_out_btn.setText(f"{shift}  退勤")
                if is_assumed:
                    self.clock_in_btn.setText(f"{shift}  出勤(想定)")
                    self.clock_out_btn.setEnabled(False)
                else:
                    self.clock_in_btn.setText(f"{shift}  出勤")
                    self.clock_out_btn.setEnabled(True)
            elif shift:
                self.clock_in_btn.setText(shift)
                self.clock_in_btn.setEnabled(True)
                self.clock_out_btn.setText("退  勤")
                self.clock_out_btn.setEnabled(False)
            else:
                self.clock_in_btn.setText("出  勤")
                self.clock_in_btn.setEnabled(True)
                self.clock_out_btn.setText("退  勤")
                self.clock_out_btn.setEnabled(True)
        self._update_batch_ui_state()

    def _on_date_selected(self, d: date) -> None:
        self.selected_date_label.setText(_format_date_badge(d))

    def _on_date_double_clicked(self, d: date) -> None:
        self._add_date_to_batch(d)

    def _add_to_batch(self) -> None:
        if self.calendar:
            d = self.calendar.get_selected_date()
            if d:
                self._add_date_to_batch(d)

    def _add_date_to_batch(self, d: date) -> None:
        if d not in self._batch_dates:
            self._batch_dates.append(d)
            self._batch_dates.sort()
            self._refresh_batch_display()
            self._update_batch_ui_state()

    def _clear_batch(self) -> None:
        self._batch_dates.clear()
        self._refresh_batch_display()
        self._on_shift_changed(self.shift_combo.currentText())

    def _update_batch_ui_state(self) -> None:
        """一括記入リストの状態に応じてUI全体を更新する"""
        shift = self.shift_combo.currentText()
        has_dates = bool(self._batch_dates)

        if has_dates:
            # 想定入力を強制チェック＆ロック
            self.assumed_check.setChecked(True)
            self.assumed_check.setEnabled(False)
            # TeamsPostなしをロック
            self.no_post_check.setEnabled(False)
            # 打刻ボタンを一括モード表示に切替（両方グレーアウト）
            self.clock_in_btn.setText("一括List選択中")
            self.clock_in_btn.setEnabled(False)
            self.clock_out_btn.setEnabled(False)
            # 非対応シフトのエラー表示
            if shift and shift not in _BATCH_ALLOWED_SHIFTS:
                self.batch_error_label.setText(
                    f"{shift} は一括記入に対応していません。"
                )
                self.batch_error_label.setVisible(True)
            else:
                self.batch_error_label.setVisible(False)
        else:
            # ロック解除
            self.assumed_check.setEnabled(True)
            self.no_post_check.setEnabled(True)
            self.batch_error_label.setVisible(False)

        # 一括記入ボタンの有効条件
        can_batch = False
        if has_dates:
            if shift in REALTIME_SHIFTS and self.assumed_check.isChecked():
                can_batch = True
            elif shift in {"シフト休", "1.0日有給"}:
                can_batch = True
        self.batch_write_btn.setEnabled(can_batch)

    def _refresh_batch_display(self) -> None:
        text = "  ".join(d.strftime("%Y/%m/%d") for d in self._batch_dates)
        self.batch_dates_edit.setPlainText(text)

    def _get_selected_date(self) -> date:
        if self.calendar:
            d = self.calendar.get_selected_date()
            if d:
                return d
        return get_today()

    def _get_work_style(self) -> str:
        return "リモート" if self.remote_radio.isChecked() else "出社"

    def _on_timesheet_check(self) -> None:
        """Timesheet Check ボタン（未実装）"""
        QMessageBox.information(self, "Timesheet Check", "この機能は未実装です。")

    # ─────────────── 打刻アクション ───────────────

    def on_clock_in(self) -> None:
        """出勤ボタン押下"""
        if not ta:
            QMessageBox.critical(self, "モジュールエラー",
                "timesheet_actions モジュールが読み込めません。")
            return

        shift = self.shift_combo.currentText()
        if not shift:
            QMessageBox.warning(self, "入力エラー", "業務形態を選択してください。")
            return
        work_style = self._get_work_style()
        target_date = self._get_selected_date()
        is_assumed = self.assumed_check.isChecked()
        no_post = self.no_post_check.isChecked()

        def late_reason_cb():
            from assets.dialogs.late_reason_dialog import LateReasonDialog
            dlg = LateReasonDialog(self)
            if dlg.exec_():
                return dlg.get_reason()
            return None

        def half_day_cb():
            from assets.dialogs.half_day_dialog import HalfDayDialog
            dlg = HalfDayDialog(self)
            if dlg.exec_():
                return {
                    "start": dlg.get_start_time(),
                    "end": dlg.get_end_time(),
                    "remark": dlg.get_remark(),
                }
            return None

        def remark_cb(title="備考入力"):
            from assets.dialogs.remark_dialog import RemarkDialog
            dlg = RemarkDialog(title=title, parent=self)
            if dlg.exec_():
                return dlg.get_remark()
            return None

        try:
            self._show_loading()
            try:
                ok, teams_error = ta.clock_in(
                    config=self.config,
                    shift=shift,
                    work_style=work_style,
                    target_date=target_date,
                    is_assumed=is_assumed,
                    no_post=no_post,
                    late_reason_cb=late_reason_cb,
                    half_day_cb=half_day_cb,
                    remark_cb=remark_cb,
                    status_cb=self.set_status,
                )
            finally:
                self._hide_loading()
            if ok:
                shift_line = shift
                if shift in REALTIME_SHIFTS and not is_assumed:
                    shift_line += f"({work_style})"
                msg = (
                    f"出勤打刻が完了しました。\n\n"
                    f"日付: {target_date.strftime('%Y/%m/%d')}\n"
                    f"シフト: {shift_line}"
                )
                if teams_error:
                    msg += f"\n\n⚠ {teams_error}"
                QMessageBox.information(self, "出勤完了", msg)
        except TimesheetNotFoundError as e:
            QMessageBox.warning(
                self, "タイムシート未検出",
                f"タイムシートが見つかりません。\n\n"
                f"検索パス：{e.folder}\n"
                f"検索タイムシート：{e.year:04d}{e.month:02d}{e.name}.xlsx"
            )
        except TimesheetLockedError as e:
            QMessageBox.critical(
                self, "ファイル書込エラー",
                f"Excelファイルが別のプロセス（Excelなど）によって\n"
                f"開かれているため更新できませんでした。\n\n"
                f"ファイル名: {e.path.name}\n\n"
                f"Excelを閉じてから再度お試しください。"
            )
        except UnknownShiftTypeError as e:
            QMessageBox.warning(self, "未定義の出勤形態", str(e))
        except TimesheetWriteError as e:
            QMessageBox.critical(self, "Excel書込エラー", str(e))
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"エラーが発生しました:\n{e}")

    def on_clock_out(self) -> None:
        """退勤ボタン押下"""
        if not ta:
            QMessageBox.critical(self, "モジュールエラー",
                "timesheet_actions モジュールが読み込めません。")
            return

        shift = self.shift_combo.currentText()
        if not shift:
            QMessageBox.warning(self, "入力エラー", "業務形態を選択してください。")
            return
        work_style = self._get_work_style()
        target_date = self._get_selected_date()
        no_post = self.no_post_check.isChecked()

        from assets.dialogs.clock_out_dialog import ClockOutDialog
        dlg = ClockOutDialog(
            shift_types=self.config.shift_types,
            managers=self.config.managers,
            is_night=(shift == "深夜"),
            parent=self
        )
        if not dlg.exec_():
            return

        is_cross_day = dlg.get_is_cross_day()
        clock_out_info = {
            "next_workday": dlg.get_next_workday(),
            "next_shift": dlg.get_next_shift(),
            "next_work_mode": dlg.get_next_work_mode(),
            "mention": dlg.get_mention(),
            "comment": dlg.get_comment(),
        }

        try:
            self._show_loading()
            try:
                ok, teams_error = ta.clock_out(
                    config=self.config,
                    shift=shift,
                    work_style=work_style,
                    target_date=target_date,
                    no_post=no_post,
                    clock_out_info=clock_out_info,
                    status_cb=self.set_status,
                    is_cross_day=is_cross_day,
                )
            finally:
                self._hide_loading()
            if ok:
                from assets.timesheet_helpers import get_now, round_time as _round_time
                _rounded = _round_time(get_now())
                clock_out_time_str = _rounded.strftime('%Y/%m/%d %H:%M')

                next_workday  = clock_out_info.get("next_workday")
                next_shift    = clock_out_info.get("next_shift", "")
                next_work_mode = clock_out_info.get("next_work_mode", "")
                if next_workday:
                    next_date_str = f"{next_workday.month}/{next_workday.day}({_WEEKDAY_JA[next_workday.weekday()]})"
                    next_line = f"{next_date_str} {next_shift}{next_work_mode}"
                else:
                    next_line = f"{next_shift}{next_work_mode}"

                msg = (
                    f"退勤打刻が完了しました。お疲れさまでした。\n\n"
                    f"退勤時刻: {clock_out_time_str}\n"
                    f"次回の出勤：{next_line}"
                )
                if teams_error:
                    msg += f"\n\n⚠ {teams_error}"

                # カスタムダイアログ（OKボタン左に画像をランダム表示）
                _images_dir = Path(__file__).parent.parent / "images"
                _image_files = (
                    [p for p in _images_dir.glob("*")
                     if p.suffix.lower() in (".png", ".jpg", ".jpeg")]
                    if _images_dir.exists() else []
                )
                _dlg = QDialog(self)
                _dlg.setWindowTitle("退勤完了")
                _dlg.setMinimumWidth(300)
                _vlay = QVBoxLayout(_dlg)
                _vlay.setSpacing(12)
                _vlay.setContentsMargins(16, 16, 16, 16)
                _text_lbl = QLabel(msg)
                _text_lbl.setWordWrap(True)
                _vlay.addWidget(_text_lbl)
                _hlay = QHBoxLayout()
                _hlay.addStretch(1)
                if _image_files:
                    _px = QPixmap(str(random.choice(_image_files)))
                    _img_lbl = QLabel()
                    _img_lbl.setPixmap(
                        _px.scaled(75, 75, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                    )
                    _hlay.addWidget(_img_lbl)
                    _hlay.addStretch(1)
                _ok = QPushButton("OK")
                _ok.setDefault(True)
                _ok.setFixedWidth(80)
                _ok.clicked.connect(_dlg.accept)
                _hlay.addWidget(_ok)
                _vlay.addLayout(_hlay)
                _dlg.exec_()
        except TimesheetNotFoundError as e:
            QMessageBox.warning(
                self, "タイムシート未検出",
                f"タイムシートが見つかりません。\n\n"
                f"検索パス：{e.folder}\n"
                f"検索タイムシート：{e.year:04d}{e.month:02d}{e.name}.xlsx"
            )
        except TimesheetLockedError as e:
            QMessageBox.critical(
                self, "ファイル書込エラー",
                f"Excelファイルが別のプロセス（Excelなど）によって\n"
                f"開かれているため更新できませんでした。\n\n"
                f"ファイル名: {e.path.name}\n\n"
                f"Excelを閉じてから再度お試しください。"
            )
        except TimesheetWriteError as e:
            QMessageBox.critical(self, "Excel書込エラー", str(e))
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"エラーが発生しました:\n{e}")

    def on_batch_write(self) -> None:
        """一括記入ボタン押下"""
        if not self._batch_dates:
            QMessageBox.warning(self, "未選択", "日付を選択してください。")
            return
        if not ta:
            QMessageBox.critical(self, "モジュールエラー",
                "timesheet_actions モジュールが読み込めません。")
            return

        shift = self.shift_combo.currentText()
        if not shift:
            QMessageBox.warning(self, "入力エラー", "業務形態を選択してください。")
            return
        work_style = self._get_work_style()

        def half_day_cb():
            from assets.dialogs.half_day_dialog import HalfDayDialog
            dlg = HalfDayDialog(self)
            if dlg.exec_():
                return {"start": dlg.get_start_time(), "end": dlg.get_end_time(), "remark": dlg.get_remark()}
            return None

        def remark_cb(title="備考入力"):
            from assets.dialogs.remark_dialog import RemarkDialog
            dlg = RemarkDialog(title=title, parent=self)
            if dlg.exec_():
                return dlg.get_remark()
            return None

        errors: list = []

        def batch_status_cb(msg: str, color: str) -> None:
            if color in ("orange", "red"):
                errors.append(msg)
            self.set_status(msg, color)

        try:
            self._show_loading()
            try:
                success, fail = ta.batch_write(
                    config=self.config,
                    dates=list(self._batch_dates),
                    shift=shift,
                    work_style=work_style,
                    half_day_cb=half_day_cb,
                    remark_cb=remark_cb,
                    status_cb=batch_status_cb,
                )
            finally:
                self._hide_loading()
            summary = f"成功: {success} 件 / 失敗: {fail} 件"
            if errors:
                detail = "\n".join(f"・{e}" for e in errors)
                QMessageBox.warning(
                    self, "一括記入完了（一部エラー）",
                    f"{summary}\n\n{detail}"
                )
            else:
                QMessageBox.information(self, "一括記入完了",
                    f"一括記入が完了しました。\n\n{summary}")
        except UnknownShiftTypeError as e:
            QMessageBox.warning(self, "未定義の出勤形態", str(e))
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"エラーが発生しました:\n{e}")

    def _show_loading(self) -> None:
        win = self.window()
        if not hasattr(self, '_overlay'):
            self._overlay = LoadingOverlay(win)
        self._overlay.resize(win.size())
        self._overlay.raise_()
        self._overlay.show()
        QApplication.processEvents()

    def _hide_loading(self) -> None:
        if hasattr(self, '_overlay'):
            self._overlay.hide()

    def update_shift_types(self, shift_types: list) -> None:
        """出勤形態コンボボックスを更新する"""
        current = self.shift_combo.currentText()
        self.shift_combo.clear()
        self.shift_combo.addItems(shift_types)
        idx = self.shift_combo.findText(current)
        # 現在の選択が新リストに存在すれば維持、なければ空選択
        self.shift_combo.setCurrentIndex(idx if idx >= 0 else -1)

    def set_status(self, msg: str, color: str = "black") -> None:
        """ステータスフィードバック（ラベル廃止したので無効化）"""
        pass
