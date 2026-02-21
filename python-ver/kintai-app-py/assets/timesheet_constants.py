"""勤怠アプリ定数定義"""

# 出勤形態別の開始時刻マップ
START_TIME_MAP = {
    "日勤": "10:00",
    "早番": "07:00",
    "遅番": "14:30",
    "深夜": "22:30",
}

# 出勤形態別の終了時刻マップ
END_TIME_MAP = {
    "日勤": "19:00",
    "早番": "16:00",
    "遅番": "23:30",
    "深夜": "31:30",
}

# リアルタイム（実際の時刻を使用する）出勤形態
REALTIME_SHIFTS = ["日勤", "早番", "遅番", "深夜"]

# 固定休暇（時刻入力不要）
VACATION_FIXED = ["シフト休", "健康診断(半日)", "1日人間ドック", "慶弔休暇"]

# 備考入力が必要な休暇
VACATION_INPUT = ["振休", "1.0日有給"]

# 半日有給
HALF_DAY_PAID = "0.5日有給"

# 遅刻許容分数
LATE_MARGIN_MIN = 10

# 丸め単位（分）
ROUND_UNIT_MIN = 15

# 勤務スタイル
WORK_STYLE_REMOTE = "リモート"
WORK_STYLE_OFFICE = "出社"

# 休暇種別の固定記入設定
VACATION_CONFIG = {
    "シフト休": {
        "shift_label": "シフト休",
        "start_time": None,
        "end_time": None,
        "remark": None,
    },
    "健康診断(半日)": {
        "shift_label": "0.5日有給",
        "start_time": "14:00",
        "end_time": "18:00",
        "remark": "午後健康診断+0.5有給",
    },
    "1日人間ドック": {
        "shift_label": "日勤",
        "start_time": "10:00",
        "end_time": "18:00",
        "remark": "1日人間ドック",
    },
    "慶弔休暇": {
        "shift_label": "慶弔休暇",
        "start_time": None,
        "end_time": None,
        "remark": None,
    },
}

# 備考入力が必要な休暇の入力設定
VACATION_INPUT_CONFIG = {
    "振休": {
        "shift_label": "シフト休",
        "dialog_prompt": "振休の詳細を入力してください:　(例)「12/25出社分」",
    },
    "1.0日有給": {
        "shift_label": "1.0日有給",
        "dialog_prompt": "有給の詳細を入力してください:　(例)「体調不良の為」「私用の為」",
    },
}
