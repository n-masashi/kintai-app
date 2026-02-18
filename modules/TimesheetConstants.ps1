# ==============================
# タイムシート操作
# ==============================

# 定数定義
$script:StartTimeMap = @{
    "日勤" = @{ Hours = 10; Minutes = 0 }
    "早番" = @{ Hours = 7;  Minutes = 0 }
    "遅番" = @{ Hours = 14; Minutes = 30 }
    "深夜" = @{ Hours = 22; Minutes = 30 }
}

$script:EndTimeMap = @{
    "日勤" = @{ Hours = 19; Minutes = 0 }
    "早番" = @{ Hours = 16; Minutes = 0 }
    "遅番" = @{ Hours = 23; Minutes = 30 }
    "深夜" = @{ Hours = 31; Minutes = 30 }
}

$script:RealtimeShiftGroups = @("日勤", "早番", "遅番", "深夜")

# 休暇グループ定義（ComboBox UI制御用: グレーアウト対象全て）
$script:VacationGroup = @("シフト休", "健康診断(半日)", "1日人間ドック", "慶弔休暇", "振休", "1.0日有給", "0.5日有給")

# 休暇グループ（ユーザー入力型）
$script:VacationInputGroup = @("振休", "1.0日有給")

$script:VacationInputConfig = @{
    "振休" = @{
        ShiftLabel   = "シフト休"
        DialogPrompt = "振休の詳細を入力してください:　(例)「12/25出社分」"
    }
    "1.0日有給" = @{
        ShiftLabel   = "1.0日有給"
        DialogPrompt = "有給の詳細を入力してください:　(例)「体調不良の為」「私用の為」"
    }
}

$script:VacationConfig = @{
    "シフト休" = @{
        ShiftLabel = "シフト休"
        StartTime  = $null
        EndTime    = $null
        Remark     = $null
    }
    "健康診断(半日)" = @{
        ShiftLabel = "0.5日有給"
        StartTime  = @{ Hours = 14; Minutes = 0 }
        EndTime    = @{ Hours = 18; Minutes = 0 }
        Remark     = "午後健康診断+0.5有給"
    }
    "1日人間ドック" = @{
        ShiftLabel = "日勤"
        StartTime  = @{ Hours = 10; Minutes = 0 }
        EndTime    = @{ Hours = 18; Minutes = 0 }
        Remark     = "1日人間ドック"
    }
    "慶弔休暇" = @{
        ShiftLabel = "慶弔休暇"
        StartTime  = $null
        EndTime    = $null
        Remark     = $null
    }
}
