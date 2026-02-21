"""時刻丸め・Excel シリアル値変換・祝日計算・ファイル検索ヘルパー"""
import calendar
import os
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Optional, Set


def get_now() -> datetime:
    """現在日時を返す。環境変数 KINTAI_TEST_DATE (YYYY-MM-DD) が設定されている場合は
    その日付に実時刻を組み合わせて返す（テスト用）。"""
    test_date = os.environ.get("KINTAI_TEST_DATE")
    if test_date:
        try:
            d = datetime.strptime(test_date, "%Y-%m-%d")
            now = datetime.now()
            return d.replace(hour=now.hour, minute=now.minute, second=now.second)
        except ValueError:
            pass
    return datetime.now()


def get_today() -> date:
    """今日の日付を返す。環境変数 KINTAI_TEST_DATE (YYYY-MM-DD) が設定されている場合は
    その日付を返す（テスト用）。"""
    test_date = os.environ.get("KINTAI_TEST_DATE")
    if test_date:
        try:
            return datetime.strptime(test_date, "%Y-%m-%d").date()
        except ValueError:
            pass
    return date.today()

try:
    import openpyxl
    from openpyxl.utils.datetime import from_excel as _from_excel
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False
    _from_excel = None

from assets.timesheet_constants import START_TIME_MAP, LATE_MARGIN_MIN, ROUND_UNIT_MIN


def round_time(dt: datetime, unit: int = ROUND_UNIT_MIN) -> datetime:
    """15分単位で四捨五入する。"""
    dt = dt.replace(second=0, microsecond=0)
    total_minutes = dt.hour * 60 + dt.minute
    rounded_minutes = round(total_minutes / unit) * unit
    new_hours = rounded_minutes // 60
    new_minutes = rounded_minutes % 60
    return dt.replace(hour=new_hours, minute=new_minutes)


def round_time_night_shift(dt: datetime, unit: int = ROUND_UNIT_MIN) -> dict:
    """
    深夜勤務用15分丸め（24h超え対応）。
    現在時刻が22時より前なら翌日とみなし+24h。
    Returns: {"hours": int, "minutes": int}
    """
    hours = dt.hour
    if hours < 22:
        hours += 24
    total_minutes = hours * 60 + dt.minute
    rounded_minutes = round(total_minutes / unit) * unit
    new_hours = rounded_minutes // 60
    new_minutes = rounded_minutes % 60
    return {"hours": new_hours, "minutes": new_minutes}


def time_to_excel_serial(hour: int, minute: int) -> float:
    """
    時刻を Excel のシリアル値（小数）に変換する。
    例: 9:00 → 0.375, 18:00 → 0.75, 0:00 → 0.0
    """
    total_minutes = hour * 60 + minute
    return total_minutes / (24 * 60)


def _nth_weekday(year: int, month: int, weekday: int, n: int) -> date:
    """
    指定した年月の第 n 回目の曜日 (0=月曜...6=日曜) を返す。
    """
    first = date(year, month, 1)
    # first.weekday() が weekday より大きい場合は翌週に補正
    diff = (weekday - first.weekday()) % 7
    first_occurrence = first + timedelta(days=diff)
    return first_occurrence + timedelta(weeks=n - 1)


def _vernal_equinox(year: int) -> int:
    """春分の日（3月）の日を返す（近似計算）"""
    if year <= 1979:
        return int(20.8357 + 0.242194 * (year - 1980) - int((year - 1983) / 4))
    elif year <= 2099:
        return int(20.8431 + 0.242194 * (year - 1980) - int((year - 1980) / 4))
    else:
        return int(21.851 + 0.242194 * (year - 1980) - int((year - 1980) / 4))


def _autumnal_equinox(year: int) -> int:
    """秋分の日（9月）の日を返す（近似計算）"""
    if year <= 1979:
        return int(23.2588 + 0.242194 * (year - 1980) - int((year - 1983) / 4))
    elif year <= 2099:
        return int(23.2488 + 0.242194 * (year - 1980) - int((year - 1980) / 4))
    else:
        return int(24.2488 + 0.242194 * (year - 1980) - int((year - 1980) / 4))


def get_holidays(year: int, month: int) -> Set[date]:
    """
    指定した年月の日本の祝日の集合を返す。
    振替休日（日曜日が祝日の翌月曜日）も含む。
    """
    # まず全祝日を計算
    holidays: Set[date] = set()

    # 固定祝日
    fixed = [
        (1, 1),   # 元旦
        (2, 11),  # 建国記念日
        (2, 23),  # 天皇誕生日（2020年以降の）
        (4, 29),  # 昭和の日
        (5, 3),   # 憲法記念日
        (5, 4),   # みどりの日
        (5, 5),   # こどもの日
        (8, 11),  # 山の日
        (11, 3),  # 文化の日
        (11, 23), # 勤労感謝の日
    ]
    for m, d in fixed:
        try:
            holidays.add(date(year, m, d))
        except ValueError:
            pass

    # 春分の日 (3月)
    try:
        holidays.add(date(year, 3, _vernal_equinox(year)))
    except ValueError:
        pass

    # 秋分の日 (9月)
    try:
        holidays.add(date(year, 9, _autumnal_equinox(year)))
    except ValueError:
        pass

    # ハッピーマンデー
    happy_mondays = [
        (1, 2),   # 成人の日 (1月第2月曜)
        (7, 3),   # 海の日 (7月第3月曜)
        (9, 3),   # 敬老の日 (9月第3月曜)
        (10, 2),  # スポーツの日 (10月第2月曜)
    ]
    for m, n in happy_mondays:
        try:
            holidays.add(_nth_weekday(year, m, 0, n))  # 0=月曜日
        except ValueError:
            pass

    # 振替休日の計算（全祝日の中から日曜日を検出）
    substitute = set()
    for h in holidays:
        if h.weekday() == 6:  # 日曜日
            candidate = h + timedelta(days=1)
            # 振替候補が別の祝日と重なる場合は更に翌日へ
            while candidate in holidays or candidate in substitute:
                candidate += timedelta(days=1)
            substitute.add(candidate)
    holidays.update(substitute)

    # 指定月のみフィルタ
    return {h for h in holidays if h.month == month and h.year == year}


def find_timesheet(folder: str, display_name: str, year: int, month: int) -> Optional[Path]:
    """
    folder 内から display_name と "{year:04d}{month:02d}" を含む .xlsx を検索する。
    例: 202602宮田.xlsx
    見つからない場合は None を返す。
    """
    target_prefix = f"{year:04d}{month:02d}"
    folder_path = Path(folder)
    if not folder_path.exists():
        return None
    for p in folder_path.glob("*.xlsx"):
        if display_name in p.name and target_prefix in p.name:
            return p
    return None


def is_late(shift: str, now: datetime) -> bool:
    """
    出勤形態と現在時刻を比較して遅刻かどうか判定する。
    LATE_MARGIN_MIN 分超過で遅刻と判断する。
    shift が START_TIME_MAP にない場合は False を返す。
    深夜シフトは日付またぎを考慮したdatetime比較を行う。
    """
    if shift not in START_TIME_MAP:
        return False
    start_str = START_TIME_MAP[shift]
    h, m = map(int, start_str.split(":"))

    if shift == "深夜":
        # 深夜: 日付またぎがあるためdatetime比較
        start_dt = now.replace(hour=h, minute=m, second=0, microsecond=0)
        if now.hour < 12:
            start_dt = start_dt - timedelta(days=1)
        limit = start_dt + timedelta(minutes=LATE_MARGIN_MIN)
        return now > limit
    else:
        # 通常: 分の比較
        fixed_start_minutes = h * 60 + m
        current_minutes = now.hour * 60 + now.minute
        return current_minutes > (fixed_start_minutes + LATE_MARGIN_MIN)


def format_date_jp(d: date) -> str:
    """
    date を "2025年01月15日（火）" 形式にフォーマットする。
    """
    weekdays = ["月", "火", "水", "木", "金", "土", "日"]
    return f"{d.year}年{d.month:02d}月{d.day:02d}日（{weekdays[d.weekday()]}）"


def get_row_for_date(ws, target_day: int) -> Optional[int]:
    """
    C列(3列目)を18行目〜48行目まで走査し、
    target_day（日の数値）に一致する行番号を返す。
    見つからない場合はNoneを返す。

    28日以前: C列の数値と直接照合（行位置に依存しない）
    29〜31日: 28日の行を基準に +1/+2/+3 行で特定
    """
    # まず28日の行を探す（29〜31日の基準になる）
    row_28 = None
    for row_num in range(18, 49):
        val = ws.cell(row=row_num, column=3).value
        try:
            if int(val) == 28:
                row_28 = row_num
                break
        except (ValueError, TypeError):
            continue

    if target_day <= 28:
        for row_num in range(18, 49):
            val = ws.cell(row=row_num, column=3).value
            try:
                if int(val) == target_day:
                    return row_num
            except (ValueError, TypeError):
                continue
        return None

    # 29〜31日: 28日行の直後
    if row_28 is None:
        return None
    offset = target_day - 28  # 29→1, 30→2, 31→3
    return row_28 + offset
