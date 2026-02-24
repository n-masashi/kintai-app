"""勤怠アプリ定数定義"""

# ──────────────────────────────────────────────────────────────────
# 出勤形態の一元定義
#
# type の種別:
#   "realtime"       : リアルタイム打刻（日勤・早番・遅番・深夜）
#   "vacation_fixed" : 固定設定の休暇（時刻・備考はここで定義）
#   "vacation_input" : 備考入力が必要な休暇
#   "custom_input"   : ダイアログで出勤形態・時刻・備考を自由入力するタイプ
#
# 新しい出勤形態を追加する場合は、このdictに1エントリ追加するだけでOK。
# timesheet_actions.py の変更は不要（既存グループの範囲内であれば）。
# ──────────────────────────────────────────────────────────────────
SHIFT_DEFINITIONS: dict = {
    # ── リアルタイムシフト ──────────────────────────────────────
    "日勤": {
        "type": "realtime",
        "start": "10:00",
        "end": "19:00",
    },
    "早番": {
        "type": "realtime",
        "start": "07:00",
        "end": "16:00",
    },
    "遅番": {
        "type": "realtime",
        "start": "14:30",
        "end": "23:30",
    },
    "深夜": {
        "type": "realtime",
        "start": "22:30",
        "end": "31:30",
    },

    # ── 固定休暇（時刻・備考は定義済み） ────────────────────────
    "シフト休": {
        "type": "vacation_fixed",
        "shift_label": "シフト休",
        "start": None,
        "end": None,
        "remark": None,
    },
    "健康診断(半日)": {
        "type": "vacation_fixed",
        "shift_label": "0.5日有給",
        "start": "14:00",
        "end": "18:00",
        "remark": "午後健康診断+0.5有給",
    },
    "1日人間ドック": {
        "type": "vacation_fixed",
        "shift_label": "日勤",
        "start": "10:00",
        "end": "18:00",
        "remark": "1日人間ドック",
    },
    "慶弔休暇": {
        "type": "vacation_fixed",
        "shift_label": "慶弔休暇",
        "start": None,
        "end": None,
        "remark": None,
    },

    # ── 備考入力が必要な休暇 ─────────────────────────────────────
    "振休": {
        "type": "vacation_input",
        "shift_label": "シフト休",
        "dialog_prompt": "振休の詳細を入力してください",
        "placeholder": "(例)「12/25出社分」",
    },
    "1.0日有給": {
        "type": "vacation_input",
        "shift_label": "1.0日有給",
        "dialog_prompt": "有給の詳細を入力してください",
        "placeholder": "(例)「私用の為」「体調不良の為」",
        "default_remark": "私用の為",
    },

    # ── カスタム入力（ダイアログで出勤形態・時刻・備考を入力） ──────
    "0.5日有給": {
        "type": "custom_input",
        "shift_label": "0.5日有給",
    },
}

# ──────────────────────────────────────────────────────────────────
# 以下は SHIFT_DEFINITIONS から自動生成（直接編集不要）
# ──────────────────────────────────────────────────────────────────

# リアルタイム（実際の時刻を使用する）出勤形態
REALTIME_SHIFTS = [k for k, v in SHIFT_DEFINITIONS.items() if v["type"] == "realtime"]

# 固定休暇（時刻入力不要）
VACATION_FIXED = [k for k, v in SHIFT_DEFINITIONS.items() if v["type"] == "vacation_fixed"]

# 備考入力が必要な休暇
VACATION_INPUT = [k for k, v in SHIFT_DEFINITIONS.items() if v["type"] == "vacation_input"]

# カスタム入力（ダイアログで時刻・備考を入力するタイプ）
CUSTOM_INPUT = next(k for k, v in SHIFT_DEFINITIONS.items() if v["type"] == "custom_input")

# 出勤形態別の開始時刻マップ
START_TIME_MAP = {k: v["start"] for k, v in SHIFT_DEFINITIONS.items() if v["type"] == "realtime"}

# 出勤形態別の終了時刻マップ
END_TIME_MAP = {k: v["end"] for k, v in SHIFT_DEFINITIONS.items() if v["type"] == "realtime"}

# 休暇種別の固定記入設定
VACATION_CONFIG = {
    k: {
        "shift_label": v["shift_label"],
        "start_time": v["start"],
        "end_time": v["end"],
        "remark": v["remark"],
    }
    for k, v in SHIFT_DEFINITIONS.items() if v["type"] == "vacation_fixed"
}

# 備考入力が必要な休暇の入力設定
VACATION_INPUT_CONFIG = {
    k: {
        "shift_label": v["shift_label"],
        "dialog_prompt": v["dialog_prompt"],
        "placeholder": v.get("placeholder", ""),
        "default_remark": v.get("default_remark", ""),
    }
    for k, v in SHIFT_DEFINITIONS.items() if v["type"] == "vacation_input"
}

# 遅刻許容分数
LATE_MARGIN_MIN = 10

# 丸め単位（分）
ROUND_UNIT_MIN = 15

# 勤務スタイル
WORK_STYLE_REMOTE = "リモート"
WORK_STYLE_OFFICE = "出社"
