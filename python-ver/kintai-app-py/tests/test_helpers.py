"""assets/timesheet_helpers.py のユニットテスト"""
import os
import pytest
from datetime import datetime, date, timedelta
from pathlib import Path
from unittest.mock import patch, MagicMock

from assets.timesheet_helpers import (
    round_time,
    round_time_night_shift,
    time_to_excel_serial,
    get_holidays,
    find_timesheet,
    is_late,
    format_date_jp,
    get_row_for_date,
    get_now,
    get_today,
)


# ─────────────────────────── round_time ───────────────────────────

class TestRoundTime:
    def test_round_down(self):
        """7分 → 0分に切り捨て"""
        dt = datetime(2026, 2, 21, 9, 7)
        r = round_time(dt)
        assert r.hour == 9 and r.minute == 0

    def test_round_up(self):
        """8分 → 15分に切り上げ"""
        dt = datetime(2026, 2, 21, 9, 8)
        r = round_time(dt)
        assert r.hour == 9 and r.minute == 15

    def test_exact_quarter(self):
        """ちょうど15分は変化なし"""
        dt = datetime(2026, 2, 21, 9, 15)
        r = round_time(dt)
        assert r.hour == 9 and r.minute == 15

    def test_round_to_next_hour(self):
        """53分 → 翌時0分"""
        dt = datetime(2026, 2, 21, 9, 53)
        r = round_time(dt)
        assert r.hour == 10 and r.minute == 0

    def test_midnight_rollover(self):
        """23:53 → 翌日0:00（日付またぎ）"""
        dt = datetime(2026, 2, 21, 23, 53)
        r = round_time(dt)
        assert r.hour == 0 and r.minute == 0

    def test_seconds_cleared(self):
        """秒・マイクロ秒はクリアされる"""
        dt = datetime(2026, 2, 21, 10, 0, 45)
        r = round_time(dt)
        assert r.second == 0 and r.microsecond == 0

    def test_half_point_rounds_up(self):
        """7.5分（ちょうど半分）は切り上げ"""
        dt = datetime(2026, 2, 21, 9, 8)  # round(8/15) = round(0.533) = 1 → 15分
        r = round_time(dt)
        assert r.minute == 15


# ─────────────────────────── round_time_night_shift ───────────────────────────

class TestRoundTimeNightShift:
    def test_after_22_no_offset(self):
        """22:35 → 22:30（+24h なし）"""
        dt = datetime(2026, 2, 21, 22, 35)
        r = round_time_night_shift(dt)
        assert r["hours"] == 22 and r["minutes"] == 30

    def test_before_22_adds_24h(self):
        """01:07 → +24h → 25:00（丸め）"""
        dt = datetime(2026, 2, 22, 1, 7)
        r = round_time_night_shift(dt)
        assert r["hours"] == 25 and r["minutes"] == 0

    def test_exact_22_no_offset(self):
        """22:00 ちょうどは +24h しない"""
        dt = datetime(2026, 2, 21, 22, 0)
        r = round_time_night_shift(dt)
        assert r["hours"] == 22 and r["minutes"] == 0

    def test_early_morning_rounding(self):
        """07:22 → +24h → 31:15（丸め）"""
        dt = datetime(2026, 2, 22, 7, 22)
        r = round_time_night_shift(dt)
        # 7+24=31, 31*60+22=1882, round(1882/15)*15=1875, 1875//60=31, 1875%60=15
        assert r["hours"] == 31 and r["minutes"] == 15


# ─────────────────────────── time_to_excel_serial ───────────────────────────

class TestTimeToExcelSerial:
    def test_9h(self):
        assert abs(time_to_excel_serial(9, 0) - 0.375) < 1e-9

    def test_18h(self):
        assert abs(time_to_excel_serial(18, 0) - 0.75) < 1e-9

    def test_0h(self):
        assert time_to_excel_serial(0, 0) == 0.0

    def test_24h(self):
        assert abs(time_to_excel_serial(24, 0) - 1.0) < 1e-9

    def test_25h(self):
        """深夜通常退勤 (+24h)"""
        assert abs(time_to_excel_serial(25, 0) - 25 / 24) < 1e-9

    def test_49h(self):
        """深夜翌々日退勤 (+48h)"""
        assert abs(time_to_excel_serial(49, 0) - 49 / 24) < 1e-9

    def test_half_hour(self):
        """9:30 → 0.395833..."""
        assert abs(time_to_excel_serial(9, 30) - 9.5 / 24) < 1e-9

    def test_22h30m(self):
        """深夜始業 22:30"""
        assert abs(time_to_excel_serial(22, 30) - 22.5 / 24) < 1e-9


# ─────────────────────────── is_late ───────────────────────────

class TestIsLate:
    def test_nichkin_late(self):
        """日勤 10:00 + 10分 = 10:10 超 → 遅刻"""
        assert is_late("日勤", datetime(2026, 2, 21, 10, 11)) is True

    def test_nichkin_on_time(self):
        """日勤 10:05 → 遅刻でない"""
        assert is_late("日勤", datetime(2026, 2, 21, 10, 5)) is False

    def test_nichkin_exactly_at_margin(self):
        """日勤 10:10 ちょうど → 遅刻でない（超過でない）"""
        assert is_late("日勤", datetime(2026, 2, 21, 10, 10)) is False

    def test_hayaban_late(self):
        """早番 07:00 + 10分超 → 遅刻"""
        assert is_late("早番", datetime(2026, 2, 21, 7, 15)) is True

    def test_hayaban_on_time(self):
        assert is_late("早番", datetime(2026, 2, 21, 7, 0)) is False

    def test_unknown_shift_not_late(self):
        """定義外シフトは常に False"""
        assert is_late("シフト休", datetime(2026, 2, 21, 10, 0)) is False

    def test_night_shift_late(self):
        """深夜 22:30 + 10分超 → 遅刻"""
        assert is_late("深夜", datetime(2026, 2, 21, 22, 45)) is True

    def test_night_shift_on_time(self):
        assert is_late("深夜", datetime(2026, 2, 21, 22, 30)) is False

    def test_night_shift_early_morning_late(self):
        """深夜勤 翌朝1時は前日比較 → 遅刻にはなる"""
        # 1時は前日22:30と比較して24h以上超過しているが、
        # is_late内でhour<12の場合start_dtを-1日して比較
        now = datetime(2026, 2, 22, 1, 0)
        # start_dt = 2/22 22:30 - 1日 = 2/21 22:30, limit = 2/21 22:40
        # now(2/22 01:00) > limit(2/21 22:40) → True
        assert is_late("深夜", now) is True


# ─────────────────────────── format_date_jp ───────────────────────────

class TestFormatDateJp:
    def test_saturday(self):
        d = date(2026, 2, 21)  # 土曜
        assert format_date_jp(d) == "2026年02月21日（土）"

    def test_monday(self):
        d = date(2026, 2, 23)  # 月曜
        assert format_date_jp(d) == "2026年02月23日（月）"

    def test_zero_padded(self):
        d = date(2026, 1, 5)
        assert "01月05日" in format_date_jp(d)


# ─────────────────────────── get_now / get_today ───────────────────────────

class TestGetNowGetToday:
    def test_get_now_without_env(self):
        env = os.environ.copy()
        env.pop("KINTAI_TEST_DATE", None)
        with patch.dict(os.environ, env, clear=True):
            result = get_now()
        assert isinstance(result, datetime)

    def test_get_now_with_env(self):
        with patch.dict(os.environ, {"KINTAI_TEST_DATE": "2026-01-15"}):
            result = get_now()
        assert result.year == 2026 and result.month == 1 and result.day == 15

    def test_get_now_invalid_env_falls_back(self):
        with patch.dict(os.environ, {"KINTAI_TEST_DATE": "invalid"}):
            result = get_now()
        assert isinstance(result, datetime)

    def test_get_today_with_env(self):
        with patch.dict(os.environ, {"KINTAI_TEST_DATE": "2026-03-01"}):
            result = get_today()
        assert result == date(2026, 3, 1)

    def test_get_today_without_env(self):
        env = os.environ.copy()
        env.pop("KINTAI_TEST_DATE", None)
        with patch.dict(os.environ, env, clear=True):
            result = get_today()
        assert isinstance(result, date)


# ─────────────────────────── get_holidays ───────────────────────────

class TestGetHolidays:
    def test_new_year(self):
        h = get_holidays(2026, 1)
        assert date(2026, 1, 1) in h

    def test_coming_of_age_day_second_monday(self):
        """成人の日: 1月第2月曜 (2026/1/12)"""
        h = get_holidays(2026, 1)
        assert date(2026, 1, 12) in h

    def test_national_foundation_day(self):
        h = get_holidays(2026, 2)
        assert date(2026, 2, 11) in h

    def test_emperors_birthday(self):
        h = get_holidays(2026, 2)
        assert date(2026, 2, 23) in h

    def test_constitution_day(self):
        h = get_holidays(2026, 5)
        assert date(2026, 5, 3) in h

    def test_greenery_day(self):
        h = get_holidays(2026, 5)
        assert date(2026, 5, 4) in h

    def test_childrens_day(self):
        h = get_holidays(2026, 5)
        assert date(2026, 5, 5) in h

    def test_culture_day(self):
        h = get_holidays(2026, 11)
        assert date(2026, 11, 3) in h

    def test_labour_thanksgiving(self):
        h = get_holidays(2026, 11)
        assert date(2026, 11, 23) in h

    def test_mountain_day(self):
        h = get_holidays(2026, 8)
        assert date(2026, 8, 11) in h

    def test_no_holidays_in_non_holiday_month(self):
        """6月は祝日なし（振替なければ）"""
        h = get_holidays(2026, 6)
        assert len(h) == 0

    def test_returns_only_specified_month(self):
        """戻り値は指定月のみ"""
        h = get_holidays(2026, 1)
        for d in h:
            assert d.month == 1 and d.year == 2026


# ─────────────────────────── find_timesheet ───────────────────────────

class TestFindTimesheet:
    def test_found(self, tmp_path):
        (tmp_path / "202602山田.xlsx").touch()
        result = find_timesheet(str(tmp_path), "山田", 2026, 2)
        assert result is not None
        assert result.name == "202602山田.xlsx"

    def test_not_found_empty_dir(self, tmp_path):
        result = find_timesheet(str(tmp_path), "山田", 2026, 2)
        assert result is None

    def test_wrong_month(self, tmp_path):
        (tmp_path / "202601山田.xlsx").touch()
        result = find_timesheet(str(tmp_path), "山田", 2026, 2)
        assert result is None

    def test_wrong_name(self, tmp_path):
        (tmp_path / "202602鈴木.xlsx").touch()
        result = find_timesheet(str(tmp_path), "山田", 2026, 2)
        assert result is None

    def test_folder_not_exist(self):
        result = find_timesheet("/nonexistent/path", "山田", 2026, 2)
        assert result is None

    def test_non_xlsx_ignored(self, tmp_path):
        (tmp_path / "202602山田.csv").touch()
        result = find_timesheet(str(tmp_path), "山田", 2026, 2)
        assert result is None

    def test_multiple_files_returns_one(self, tmp_path):
        (tmp_path / "202602山田A.xlsx").touch()
        (tmp_path / "202602山田B.xlsx").touch()
        result = find_timesheet(str(tmp_path), "山田", 2026, 2)
        assert result is not None


# ─────────────────────────── get_row_for_date ───────────────────────────

def _make_ws(day_map: dict):
    """C列（列3）に day_map[行番号]=日 を返すモックワークシート"""
    ws = MagicMock()
    def cell_side_effect(row, column):
        c = MagicMock()
        if column == 3:
            c.value = day_map.get(row)
        return c
    ws.cell.side_effect = cell_side_effect
    return ws


class TestGetRowForDate:
    def _day_map_standard(self):
        """18行〜45行に 1〜28 が入るマップ"""
        return {18 + i: i + 1 for i in range(28)}

    def test_day_1(self):
        ws = _make_ws(self._day_map_standard())
        assert get_row_for_date(ws, 1) == 18

    def test_day_15(self):
        ws = _make_ws(self._day_map_standard())
        assert get_row_for_date(ws, 15) == 32

    def test_day_28(self):
        ws = _make_ws(self._day_map_standard())
        assert get_row_for_date(ws, 28) == 45

    def test_day_29(self):
        """29日 = 28日行 + 1"""
        ws = _make_ws(self._day_map_standard())
        assert get_row_for_date(ws, 29) == 46

    def test_day_31(self):
        """31日 = 28日行 + 3"""
        ws = _make_ws(self._day_map_standard())
        assert get_row_for_date(ws, 31) == 48

    def test_day_not_found(self):
        """該当日なし → None"""
        ws = _make_ws({})
        assert get_row_for_date(ws, 1) is None

    def test_non_numeric_cell_skipped(self):
        """文字列セルは無視して正しい行を返す"""
        day_map = {18: "合計", 19: 1}
        ws = _make_ws(day_map)
        assert get_row_for_date(ws, 1) == 19
