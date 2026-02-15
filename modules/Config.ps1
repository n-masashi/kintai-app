# ==============================
# 共通設定・変数 
# ==============================

# 設定ファイルのパスを保持
$script:settingsPath = Join-Path (Split-Path -Parent $PSScriptRoot) "settings.json"

# 設定の読み込み
function Load-Settings {
    if (Test-Path $script:settingsPath) {
        try {
            $settingsContent = Get-Content -Path $script:settingsPath -Encoding UTF8 -Raw | ConvertFrom-Json

            # shift_types
            $script:shiftTypes = $settingsContent.shift_types

            # settings全体を保持（デフォルト値も含む）
            $script:settings = [PSCustomObject]@{
                shift_types = $settingsContent.shift_types
                user_info = [PSCustomObject]@{
                    ad_username = if ($settingsContent.user_info.ad_username) { $settingsContent.user_info.ad_username } else { "" }
                    full_name = if ($settingsContent.user_info.full_name) { $settingsContent.user_info.full_name } else { "" }
                    teams_principal_id = if ($settingsContent.user_info.teams_principal_id) { $settingsContent.user_info.teams_principal_id } else { "" }
                    shift_display_name = if ($settingsContent.user_info.shift_display_name) { $settingsContent.user_info.shift_display_name } else { "" }
                    timesheet_display_name = if ($settingsContent.user_info.timesheet_display_name) { $settingsContent.user_info.timesheet_display_name } else { "" }
                }
                teams_workflow = [PSCustomObject]@{
                    webhook_url = if ($settingsContent.teams_workflow.webhook_url) { $settingsContent.teams_workflow.webhook_url } else { "" }
                }
                managers = if ($settingsContent.managers) { @($settingsContent.managers) } else { @() }
                theme = if ($settingsContent.theme) { $settingsContent.theme } else { "light" }
                timesheet_folder = if ($settingsContent.timesheet_folder) { $settingsContent.timesheet_folder } else { "" }
                attendance_data_folder = if ($settingsContent.attendance_data_folder) { $settingsContent.attendance_data_folder } else { "" }
            }
        }
        catch {
            # エラー時はデフォルト値
            Initialize-DefaultSettings
        }
    }
    else {
        # ファイルが無い場合はデフォルト値
        Initialize-DefaultSettings
    }
}

# デフォルト設定の初期化
function Initialize-DefaultSettings {
    $script:shiftTypes = @("日勤", "早番","遅番", "深夜")
    $script:settings = [PSCustomObject]@{
        shift_types = $script:shiftTypes
        user_info = [PSCustomObject]@{
            ad_username = ""
            full_name = ""
            teams_principal_id = ""
            shift_display_name = ""
            timesheet_display_name = ""
        }
        teams_workflow = [PSCustomObject]@{
            webhook_url = ""
        }
        managers = @()
        theme = "light"
        timesheet_folder = ""
        attendance_data_folder = ""
    }
}

# 設定の保存
function Save-Settings {
    # 設定オブジェクトを構築
    $settingsObj = [PSCustomObject]@{
        shift_types = $script:settings.shift_types
        user_info = $script:settings.user_info
        teams_workflow = $script:settings.teams_workflow
        managers = $script:settings.managers
        theme = $script:settings.theme
        timesheet_folder = $script:settings.timesheet_folder
        attendance_data_folder = $script:settings.attendance_data_folder
    }

    try {
        # JSON形式で保存
        $json = $settingsObj | ConvertTo-Json -Depth 10
        [System.IO.File]::WriteAllText($script:settingsPath, $json, (New-Object System.Text.UTF8Encoding $true))
    }
    catch {
        [System.Windows.MessageBox]::Show("設定の保存に失敗しました: $_", "エラー", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
}

# 設定を読み込む
Load-Settings

# 状態管理
$script:dispYear    = (Get-Date).Year
$script:dispMonth   = (Get-Date).Month
$script:selectedDay = (Get-Date).Day
$script:dateList    = New-Object System.Collections.ArrayList

# 第N X曜日を求める関数
function Get-NthWeekday {
    param([int]$Year, [int]$Month, [int]$Nth, [int]$DayOfWeek)
    $d = Get-Date -Year $Year -Month $Month -Day 1
    $count = 0
    while ($true) {
        if ([int]$d.DayOfWeek -eq $DayOfWeek) {
            $count++
            if ($count -eq $Nth) {
                return "$Year/$Month/$($d.Day)"
            }
        }
        $d = $d.AddDays(1)
    }
}

# 日本の祝日を取得する関数
function Get-JapaneseHolidays {
    param([int]$Year)
    $holidays = @{}

    # 固定祝日
    $holidays["$Year/1/1"]   = "元日"
    $holidays["$Year/2/11"]  = "建国記念の日"
    $holidays["$Year/2/23"]  = "天皇誕生日"
    $holidays["$Year/4/29"]  = "昭和の日"
    $holidays["$Year/5/3"]   = "憲法記念日"
    $holidays["$Year/5/4"]   = "みどりの日"
    $holidays["$Year/5/5"]   = "こどもの日"
    $holidays["$Year/8/11"]  = "山の日"
    $holidays["$Year/11/3"]  = "文化の日"
    $holidays["$Year/11/23"] = "勤労感謝の日"

    # ハッピーマンデー
    $holidays[(Get-NthWeekday $Year 1 2 1)]  = "成人の日"
    $holidays[(Get-NthWeekday $Year 7 3 1)]  = "海の日"
    $holidays[(Get-NthWeekday $Year 9 3 1)]  = "敬老の日"
    $holidays[(Get-NthWeekday $Year 10 2 1)] = "スポーツの日"

    # 春分の日（概算）
    $springDay = [math]::Floor(20.8431 + 0.242194 * ($Year - 1980) - [math]::Floor(($Year - 1980) / 4))
    $holidays["$Year/3/$springDay"] = "春分の日"

    # 秋分の日（概算）
    $autumnDay = [math]::Floor(23.2488 + 0.242194 * ($Year - 1980) - [math]::Floor(($Year - 1980) / 4))
    $holidays["$Year/9/$autumnDay"] = "秋分の日"

    # 振替休日：祝日が日曜の場合、翌月曜が休み
    $keys = @($holidays.Keys)
    foreach ($key in $keys) {
        $d = [datetime]$key
        if ($d.DayOfWeek -eq [DayOfWeek]::Sunday) {
            $next = $d.AddDays(1)
            $nextKey = "$($next.Year)/$($next.Month)/$($next.Day)"
            if (-not $holidays.ContainsKey($nextKey)) {
                $holidays[$nextKey] = "振替休日"
            }
        }
    }

    return $holidays
}
