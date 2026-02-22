"""assets/timesheet_actions.py のユニットテスト"""
import pytest
from datetime import datetime, date, timedelta
from pathlib import Path
from unittest.mock import patch, MagicMock, call

from assets.timesheet_helpers import time_to_excel_serial
from assets.timesheet_actions import (
    clock_in, clock_out, batch_write, output_csv, write_to_excel,
    TimesheetNotFoundError, TimesheetLockedError, TimesheetWriteError,
    UnknownShiftTypeError,
)


# ────────── ヘルパー ──────────

def _make_mock_wb(start_serial=0.375, row_for_date=25):
    """openpyxlワークシートのモックを返す"""
    ws = MagicMock()
    ws.cell.return_value.value = start_serial
    wb = MagicMock()
    wb.active = ws
    return wb


def _noop_status(msg, color):
    pass


# ────────── output_csv ──────────

class TestOutputCsv:
    def test_creates_csv(self, tmp_path, base_config):
        base_config.output_folder = str(tmp_path)
        base_config.shift_display_name = "山田"
        output_csv(base_config, "日勤", "出社", date(2026, 2, 21))
        assert (tmp_path / "山田.csv").exists()

    def test_csv_content_office(self, tmp_path, base_config):
        base_config.output_folder = str(tmp_path)
        base_config.shift_display_name = "山田"
        output_csv(base_config, "日勤", "出社", date(2026, 2, 21))
        content = (tmp_path / "山田.csv").read_text(encoding="utf-8")
        assert "日勤" in content
        assert "(ﾃ" not in content

    def test_csv_content_remote(self, tmp_path, base_config):
        base_config.output_folder = str(tmp_path)
        base_config.shift_display_name = "山田"
        output_csv(base_config, "日勤", "リモート", date(2026, 2, 21))
        content = (tmp_path / "山田.csv").read_text(encoding="utf-8")
        assert "日勤(ﾃ" in content

    def test_csv_fallback_when_output_folder_empty(self, base_config):
        """output_folder 空 → _BASE_DIR/attendance_data に作成（エラーにならない）"""
        base_config.output_folder = ""
        base_config.shift_display_name = "山田"
        # エラーが起きなければOK（実際にファイルを作ろうとするが例外は握り潰す）
        output_csv(base_config, "早番", "リモート", date(2026, 2, 21))

    def test_csv_overwrite(self, tmp_path, base_config):
        """2回呼ぶと上書きされる"""
        base_config.output_folder = str(tmp_path)
        base_config.shift_display_name = "山田"
        output_csv(base_config, "日勤", "出社", date(2026, 2, 21))
        output_csv(base_config, "早番", "リモート", date(2026, 2, 21))
        content = (tmp_path / "山田.csv").read_text(encoding="utf-8")
        assert "早番(ﾃ" in content
        assert "日勤" not in content


# ────────── write_to_excel ──────────

class TestWriteToExcel:
    def _make_wb_with_days(self, days_in_col3):
        """C列(col=3)に日付値が入るモックワークブック"""
        ws = MagicMock()
        def cell_factory(row, column):
            c = MagicMock()
            if column == 3:
                c.value = days_in_col3.get(row)
            else:
                c.value = None
            return c
        ws.cell.side_effect = cell_factory
        wb = MagicMock()
        wb.active = ws
        return wb, ws

    @patch("assets.timesheet_actions.openpyxl")
    @patch("assets.timesheet_actions.OPENPYXL_AVAILABLE", True)
    def test_write_sets_shift_label(self, mock_openpyxl, tmp_path):
        wb, ws = self._make_wb_with_days({18 + i: i + 1 for i in range(28)})
        mock_openpyxl.load_workbook.return_value = wb
        row_data = {
            "date": date(2026, 2, 1),
            "shift_label": "日勤",
            "start_time": time_to_excel_serial(10, 0),
            "end_time": None,
            "overtime_type": None,
            "remark": None,
        }
        write_to_excel(tmp_path / "test.xlsx", row_data)
        # E列(5) に "日勤" が書かれたか
        ws.cell.assert_any_call(row=18, column=5)

    @patch("assets.timesheet_actions.openpyxl")
    @patch("assets.timesheet_actions.OPENPYXL_AVAILABLE", True)
    def test_write_raises_on_row_not_found(self, mock_openpyxl, tmp_path):
        wb, _ = self._make_wb_with_days({})  # 空 = 行が見つからない
        mock_openpyxl.load_workbook.return_value = wb
        row_data = {
            "date": date(2026, 2, 1),
            "shift_label": "日勤",
            "start_time": None, "end_time": None,
            "overtime_type": None, "remark": None,
        }
        with pytest.raises(TimesheetWriteError):
            write_to_excel(tmp_path / "test.xlsx", row_data)

    @patch("assets.timesheet_actions.openpyxl")
    @patch("assets.timesheet_actions.OPENPYXL_AVAILABLE", True)
    def test_write_permission_error_raises_locked(self, mock_openpyxl, tmp_path):
        mock_openpyxl.load_workbook.side_effect = PermissionError
        row_data = {
            "date": date(2026, 2, 1),
            "shift_label": "日勤",
            "start_time": None, "end_time": None,
            "overtime_type": None, "remark": None,
        }
        with pytest.raises(TimesheetLockedError):
            write_to_excel(tmp_path / "test.xlsx", row_data)


# ────────── clock_out: ターゲット日付とシリアル値 ──────────

class TestClockOutTargetAndSerial:
    """退勤処理の日付決定・シリアル値計算を集中テスト"""

    def _run_clock_out(self, now_dt, shift, is_cross_day, base_config,
                       start_serial=None):
        """clock_out を最小モックで実行し row_data を返す"""
        if start_serial is None:
            start_serial = time_to_excel_serial(10, 0)  # 10:00

        with patch("assets.timesheet_actions.get_now", return_value=now_dt), \
             patch("assets.timesheet_actions._find_xlsx_or_raise",
                   return_value=Path("/fake/file.xlsx")), \
             patch("assets.timesheet_actions.OPENPYXL_AVAILABLE", True), \
             patch("assets.timesheet_actions.openpyxl") as mock_opx, \
             patch("assets.timesheet_actions.get_row_for_date", return_value=25), \
             patch("assets.timesheet_actions.write_to_excel", return_value=True) as mock_write:

            mock_opx.load_workbook.return_value = _make_mock_wb(start_serial)

            ok, _ = clock_out(
                config=base_config,
                shift=shift,
                work_style="リモート",
                target_date=date(2026, 2, 1),  # 上書きされる
                no_post=True,
                clock_out_info={},
                status_cb=_noop_status,
                is_cross_day=is_cross_day,
            )
            assert ok
            return mock_write.call_args[0][1]  # row_data

    # --- ターゲット日付 ---

    def test_target_normal_shift_no_cross(self, base_config):
        """通常シフト・通常退勤 → 当日"""
        now = datetime(2026, 2, 21, 18, 0)
        row = self._run_clock_out(now, "日勤", False, base_config)
        assert row["date"] == date(2026, 2, 21)

    def test_target_normal_shift_cross_day(self, base_config):
        """通常シフト・日跨ぎ退勤 → 前日"""
        now = datetime(2026, 2, 22, 2, 0)
        row = self._run_clock_out(now, "日勤", True, base_config)
        assert row["date"] == date(2026, 2, 21)

    def test_target_night_shift_no_cross(self, base_config):
        """深夜・通常退勤 → 前日"""
        now = datetime(2026, 2, 22, 1, 0)
        row = self._run_clock_out(now, "深夜", False, base_config,
                                  start_serial=time_to_excel_serial(22, 30))
        assert row["date"] == date(2026, 2, 21)

    def test_target_night_shift_cross_day(self, base_config):
        """深夜・日跨ぎ（翌々日退勤）→ 前々日"""
        now = datetime(2026, 2, 7, 1, 0)
        row = self._run_clock_out(now, "深夜", True, base_config,
                                  start_serial=time_to_excel_serial(22, 30))
        assert row["date"] == date(2026, 2, 5)

    # --- シリアル値 ---

    def test_serial_normal_no_cross(self, base_config):
        """通常・通常退勤 → hour のみ"""
        now = datetime(2026, 2, 21, 18, 0)
        row = self._run_clock_out(now, "日勤", False, base_config)
        assert abs(row["end_time"] - time_to_excel_serial(18, 0)) < 1e-9

    def test_serial_normal_cross_day(self, base_config):
        """通常・日跨ぎ → +24h"""
        now = datetime(2026, 2, 22, 1, 0)
        row = self._run_clock_out(now, "日勤", True, base_config)
        assert abs(row["end_time"] - time_to_excel_serial(25, 0)) < 1e-9

    def test_serial_night_no_cross(self, base_config):
        """深夜・通常退勤 → +24h"""
        now = datetime(2026, 2, 22, 1, 0)
        row = self._run_clock_out(now, "深夜", False, base_config,
                                  start_serial=time_to_excel_serial(22, 30))
        assert abs(row["end_time"] - time_to_excel_serial(25, 0)) < 1e-9

    def test_serial_night_cross_day(self, base_config):
        """深夜・翌々日退勤 → +48h"""
        now = datetime(2026, 2, 7, 1, 0)
        row = self._run_clock_out(now, "深夜", True, base_config,
                                  start_serial=time_to_excel_serial(22, 30))
        assert abs(row["end_time"] - time_to_excel_serial(49, 0)) < 1e-9

    def test_serial_rounding_applied(self, base_config):
        """退勤時刻は15分丸めが適用される"""
        now = datetime(2026, 2, 21, 18, 7)   # 18:07 → 18:00
        row = self._run_clock_out(now, "日勤", False, base_config)
        assert abs(row["end_time"] - time_to_excel_serial(18, 0)) < 1e-9

    def test_overtime_detected(self, base_config):
        """9時間超 → overtime_type = '客先指示'"""
        now = datetime(2026, 2, 21, 20, 0)   # 10:00 〜 20:00 = 10h
        row = self._run_clock_out(now, "日勤", False, base_config,
                                  start_serial=time_to_excel_serial(10, 0))
        assert row["overtime_type"] == "客先指示"

    def test_no_overtime_within_9h(self, base_config):
        """9時間以内 → overtime_type = None"""
        now = datetime(2026, 2, 21, 18, 0)   # 10:00 〜 18:00 = 8h
        row = self._run_clock_out(now, "日勤", False, base_config,
                                  start_serial=time_to_excel_serial(10, 0))
        assert row["overtime_type"] is None


# ────────── clock_in ──────────

class TestClockIn:
    def _run_clock_in(self, shift, is_assumed=False, now_dt=None,
                      late_reason_cb=None, half_day_cb=None, remark_cb=None,
                      base_config=None):
        if now_dt is None:
            now_dt = datetime(2026, 2, 21, 10, 0)
        if late_reason_cb is None:
            late_reason_cb = lambda: None
        if half_day_cb is None:
            half_day_cb = lambda: None
        if remark_cb is None:
            remark_cb = lambda title="": None

        with patch("assets.timesheet_actions.get_now", return_value=now_dt), \
             patch("assets.timesheet_actions.get_today",
                   return_value=now_dt.date()), \
             patch("assets.timesheet_actions._find_xlsx_or_raise",
                   return_value=Path("/fake/file.xlsx")), \
             patch("assets.timesheet_actions.write_to_excel",
                   return_value=True) as mock_write, \
             patch("assets.timesheet_actions.output_csv"):

            ok, _ = clock_in(
                config=base_config,
                shift=shift,
                work_style="リモート",
                target_date=now_dt.date(),
                is_assumed=is_assumed,
                no_post=True,
                late_reason_cb=late_reason_cb,
                half_day_cb=half_day_cb,
                remark_cb=remark_cb,
                status_cb=_noop_status,
            )
            if mock_write.called:
                return ok, mock_write.call_args[0][1]
            return ok, None

    def test_realtime_normal(self, base_config):
        """日勤・通常出勤 → shift_label=日勤, start_time=10:00"""
        ok, row = self._run_clock_in("日勤", base_config=base_config)
        assert ok
        assert row["shift_label"] == "日勤"
        assert abs(row["start_time"] - time_to_excel_serial(10, 0)) < 1e-9
        assert row["end_time"] is None

    def test_realtime_assumed(self, base_config):
        """想定記入 → start + end 両方固定値"""
        ok, row = self._run_clock_in("日勤", is_assumed=True, base_config=base_config)
        assert ok
        assert row["start_time"] is not None
        assert row["end_time"] is not None

    def test_realtime_hayaban(self, base_config):
        """早番 07:00 出勤"""
        ok, row = self._run_clock_in(
            "早番",
            now_dt=datetime(2026, 2, 21, 7, 0),
            base_config=base_config,
        )
        assert ok
        assert row["shift_label"] == "早番"
        assert abs(row["start_time"] - time_to_excel_serial(7, 0)) < 1e-9

    def test_realtime_late(self, base_config):
        """遅刻: shift_label='遅刻', L列に理由"""
        ok, row = self._run_clock_in(
            "日勤",
            now_dt=datetime(2026, 2, 21, 10, 30),  # 10:10超 → 遅刻
            late_reason_cb=lambda: "電車遅延",
            base_config=base_config,
        )
        assert ok
        assert row["shift_label"] == "遅刻"
        assert row["remark"] == "電車遅延"

    def test_realtime_late_cancel(self, base_config):
        """遅刻ダイアログでキャンセル → ok=False"""
        ok, _ = self._run_clock_in(
            "日勤",
            now_dt=datetime(2026, 2, 21, 10, 30),
            late_reason_cb=lambda: None,  # キャンセル
            base_config=base_config,
        )
        assert ok is False

    def test_vacation_fixed_shift_rest(self, base_config):
        """シフト休 → shift_label='シフト休', 時刻なし"""
        ok, row = self._run_clock_in("シフト休", base_config=base_config)
        assert ok
        assert row["shift_label"] == "シフト休"
        assert row["start_time"] is None
        assert row["end_time"] is None

    def test_vacation_fixed_health_checkup(self, base_config):
        """健康診断(半日) → shift_label='0.5日有給', 14:00-18:00"""
        ok, row = self._run_clock_in("健康診断(半日)", base_config=base_config)
        assert ok
        assert row["shift_label"] == "0.5日有給"
        assert abs(row["start_time"] - time_to_excel_serial(14, 0)) < 1e-9
        assert abs(row["end_time"] - time_to_excel_serial(18, 0)) < 1e-9

    def test_vacation_input_furikyu(self, base_config):
        """振休 → remark_cb が呼ばれ備考が入る"""
        ok, row = self._run_clock_in(
            "振休",
            remark_cb=lambda title="": "12/25出社分",
            base_config=base_config,
        )
        assert ok
        assert row["shift_label"] == "シフト休"
        assert row["remark"] == "12/25出社分"

    def test_vacation_input_cancel(self, base_config):
        """振休ダイアログでキャンセル → ok=False"""
        ok, _ = self._run_clock_in(
            "振休",
            remark_cb=lambda title="": None,
            base_config=base_config,
        )
        assert ok is False

    def test_half_day_paid(self, base_config):
        """0.5日有給 → half_day_cb の値を使う"""
        from unittest.mock import MagicMock
        start_q = MagicMock()
        start_q.hour.return_value = 9
        start_q.minute.return_value = 0
        end_q = MagicMock()
        end_q.hour.return_value = 13
        end_q.minute.return_value = 0

        ok, row = self._run_clock_in(
            "0.5日有給",
            half_day_cb=lambda: {"start": start_q, "end": end_q, "remark": ""},
            base_config=base_config,
        )
        assert ok
        assert row["shift_label"] == "0.5日有給"
        assert abs(row["start_time"] - time_to_excel_serial(9, 0)) < 1e-9

    def test_unknown_shift_raises(self, base_config):
        """未定義シフト → UnknownShiftTypeError"""
        with pytest.raises(UnknownShiftTypeError):
            with patch("assets.timesheet_actions.get_now",
                       return_value=datetime(2026, 2, 21, 10, 0)), \
                 patch("assets.timesheet_actions.get_today",
                       return_value=date(2026, 2, 21)), \
                 patch("assets.timesheet_actions._find_xlsx_or_raise",
                       return_value=Path("/fake/file.xlsx")), \
                 patch("assets.timesheet_actions.write_to_excel",
                       return_value=True):
                clock_in(
                    config=base_config,
                    shift="存在しないシフト",
                    work_style="リモート",
                    target_date=date(2026, 2, 21),
                    is_assumed=False,
                    no_post=True,
                    late_reason_cb=lambda: None,
                    half_day_cb=lambda: None,
                    remark_cb=lambda title="": None,
                    status_cb=_noop_status,
                )


# ────────── batch_write ──────────

class TestBatchWrite:
    def _run_batch(self, dates, shift, base_config,
                   half_day_cb=None, remark_cb=None):
        if half_day_cb is None:
            half_day_cb = lambda: None
        if remark_cb is None:
            remark_cb = lambda title="": None

        with patch("assets.timesheet_actions._find_xlsx_or_raise",
                   return_value=Path("/fake/file.xlsx")), \
             patch("assets.timesheet_actions.write_to_excel", return_value=True):
            return batch_write(
                config=base_config,
                dates=dates,
                shift=shift,
                work_style="リモート",
                half_day_cb=half_day_cb,
                remark_cb=remark_cb,
                status_cb=_noop_status,
            )

    def test_batch_realtime_success(self, base_config):
        dates = [date(2026, 2, 10), date(2026, 2, 11), date(2026, 2, 12)]
        success, fail = self._run_batch(dates, "日勤", base_config)
        assert success == 3 and fail == 0

    def test_batch_shift_rest(self, base_config):
        dates = [date(2026, 2, 10), date(2026, 2, 11)]
        success, fail = self._run_batch(dates, "シフト休", base_config)
        assert success == 2 and fail == 0

    def test_batch_partial_failure(self, base_config):
        """一部タイムシート未検出 → 失敗件数カウント"""
        dates = [date(2026, 2, 10), date(2026, 2, 11)]
        with patch("assets.timesheet_actions._find_xlsx_or_raise",
                   side_effect=[
                       Path("/fake/file.xlsx"),
                       TimesheetNotFoundError("/f", "山田", 2026, 2),
                   ]), \
             patch("assets.timesheet_actions.write_to_excel", return_value=True):
            success, fail = batch_write(
                config=base_config,
                dates=dates,
                shift="日勤",
                work_style="リモート",
                half_day_cb=lambda: None,
                remark_cb=lambda title="": None,
                status_cb=_noop_status,
            )
        assert success == 1 and fail == 1

    def test_batch_unknown_shift_raises(self, base_config):
        with pytest.raises(UnknownShiftTypeError):
            self._run_batch([date(2026, 2, 10)], "存在しないシフト", base_config)

    def test_batch_vacation_input_cancel(self, base_config):
        """振休ダイアログキャンセル → (0, 0) を返す"""
        success, fail = self._run_batch(
            [date(2026, 2, 10)], "振休", base_config,
            remark_cb=lambda title="": None,
        )
        assert success == 0 and fail == 0


# ────────── _find_xlsx_or_raise ──────────

class TestFindXlsxOrRaise:
    from assets.timesheet_actions import _find_xlsx_or_raise

    def test_raises_when_no_folder(self, base_config):
        from assets.timesheet_actions import _find_xlsx_or_raise
        base_config.timesheet_folder = ""
        with pytest.raises(TimesheetNotFoundError):
            _find_xlsx_or_raise(base_config, date(2026, 2, 21))

    def test_raises_when_no_name(self, base_config):
        from assets.timesheet_actions import _find_xlsx_or_raise
        base_config.timesheet_display_name = ""
        base_config.display_name = ""
        with pytest.raises(TimesheetNotFoundError):
            _find_xlsx_or_raise(base_config, date(2026, 2, 21))

    def test_raises_when_file_not_found(self, tmp_path, base_config):
        from assets.timesheet_actions import _find_xlsx_or_raise
        base_config.timesheet_folder = str(tmp_path)
        with pytest.raises(TimesheetNotFoundError):
            _find_xlsx_or_raise(base_config, date(2026, 2, 21))

    def test_returns_path_when_found(self, tmp_path, base_config):
        from assets.timesheet_actions import _find_xlsx_or_raise
        base_config.timesheet_folder = str(tmp_path)
        base_config.timesheet_display_name = "山田"
        (tmp_path / "202602山田.xlsx").touch()
        result = _find_xlsx_or_raise(base_config, date(2026, 2, 21))
        assert result.name == "202602山田.xlsx"
