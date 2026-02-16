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

# ==============================
# ヘルパー関数
# ==============================

# タイムシートファイルパスを取得
function Get-TimesheetPath {
    param([int]$Year, [int]$Month)

    $folder = $script:settings.timesheet_folder
    if ([string]::IsNullOrWhiteSpace($folder)) {
        return $null
    }

    # 相対パスの場合、プロジェクトルートを基準に解決
    if (-not [System.IO.Path]::IsPathRooted($folder)) {
        $folder = Join-Path $script:ProjectRoot $folder
    }

    $displayName = $script:settings.user_info.timesheet_display_name
    $yyyymm = "{0:D4}{1:D2}" -f $Year, $Month
    $fileName = "${yyyymm}${displayName}.xlsx"
    return Join-Path $folder $fileName
}

# Excelシリアル値に変換 (時刻 → 時/24 + 分/1440)
function ConvertTo-ExcelSerial {
    param([int]$Hours, [int]$Minutes)
    return $Hours / 24.0 + $Minutes / 1440.0
}

# 15分単位に丸める（四捨五入）
function Get-RoundedQuarter {
    param([datetime]$DateTime)

    $totalMinutes = $DateTime.Hour * 60 + $DateTime.Minute
    $roundedMinutes = [Math]::Round($totalMinutes / 15.0) * 15
    $newHours = [Math]::Floor($roundedMinutes / 60)
    $newMinutes = $roundedMinutes % 60
    return @{ Hours = [int]$newHours; Minutes = [int]$newMinutes }
}

# 深夜勤務用15分丸め（24h超え対応）
# 現在時刻が22時より前なら翌日とみなし+24h
function Get-NightShiftRoundedQuarter {
    param([datetime]$DateTime)

    $hours = $DateTime.Hour
    if ($hours -lt 22) { $hours += 24 }
    $totalMinutes = $hours * 60 + $DateTime.Minute
    $roundedMinutes = [Math]::Round($totalMinutes / 15.0) * 15
    $newHours = [Math]::Floor($roundedMinutes / 60)
    $newMinutes = $roundedMinutes % 60
    return @{ Hours = [int]$newHours; Minutes = [int]$newMinutes }
}

# C列を走査して対象日の行番号を返す
function Find-DayRow {
    param($Worksheet, [int]$Day)

    for ($row = 18; $row -le 48; $row++) {
        $cellValue = $Worksheet.Cells.Item($row, 3).Value2
        if ($null -ne $cellValue) {
            try {
                if ([int]$cellValue -eq $Day) {
                    return $row
                }
            } catch {
                # 数値でないセルはスキップ
            }
        }
    }
    return -1
}

# 遅刻理由入力ダイアログ
function Show-LateReasonDialog {
    param($OwnerWindow)

    $dialogXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="遅刻理由" Width="400" Height="180"
        WindowStartupLocation="CenterOwner" ResizeMode="NoResize">
    <StackPanel Margin="20">
        <TextBlock Text="遅刻理由を入力してください:　(例) 寝坊の為" FontSize="12" Margin="0,0,0,10"/>
        <TextBox x:Name="TxtLateReason" FontSize="12" Height="30" TextWrapping="Wrap" AcceptsReturn="True" VerticalScrollBarVisibility="Auto"/>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,15,0,0">
            <Button x:Name="BtnOk" Content="OK" Width="80" Height="30" Margin="0,0,10,0" IsDefault="True"/>
            <Button x:Name="BtnCancel" Content="キャンセル" Width="80" Height="30" IsCancel="True"/>
        </StackPanel>
    </StackPanel>
</Window>
"@

    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($dialogXaml))
    $dialog = [System.Windows.Markup.XamlReader]::Load($reader)
    $reader.Close()

    $dialog.Owner = $OwnerWindow

    $txtReason = $dialog.FindName("TxtLateReason")
    $btnOk = $dialog.FindName("BtnOk")
    $btnCancel = $dialog.FindName("BtnCancel")

    # 結果格納用のオブジェクト（クロージャのスコープ問題回避）
    $resultHolder = [PSCustomObject]@{ Value = $null }

    $btnOk.Add_Click({
        $resultHolder.Value = $txtReason.Text
        $dialog.DialogResult = $true
    }.GetNewClosure())

    $btnCancel.Add_Click({
        $dialog.DialogResult = $false
    }.GetNewClosure())

    $result = $dialog.ShowDialog()
    if ($result) {
        return $resultHolder.Value
    }
    return $null
}

# ==============================
# メイン関数
# ==============================

# 出勤処理
function Write-ClockIn {
    param($Window)

    # UI値取得
    $cmbShiftType = $Window.FindName("CmbShiftType")
    $chkEstimatedInput = $Window.FindName("ChkEstimatedInput")
    $shiftType = $cmbShiftType.SelectedItem

    # 0.5日有給（時間+備考ユーザー入力型）
    if ($shiftType -eq "0.5日有給") {
        Write-HalfDayLeaveClockIn -Window $Window
        return
    }

    # 休暇グループ（ユーザー入力型）の場合は専用処理
    if ($shiftType -in $script:VacationInputGroup) {
        Write-VacationInputClockIn -Window $Window -ShiftType $shiftType
        return
    }

    # 休暇グループ（固定記載）の場合は専用処理
    if ($shiftType -in $script:VacationGroup) {
        Write-VacationClockIn -Window $Window -ShiftType $shiftType
        return
    }

    # リアルタイム勤務グループか判定
    if ($shiftType -notin $script:RealtimeShiftGroups) {
        [System.Windows.MessageBox]::Show("「${shiftType}」の出勤処理は未実装です。", "情報", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::None)
        return
    }

    $isEstimated = $chkEstimatedInput.IsChecked

    # 日付取得
    if ($isEstimated) {
        $targetYear = $script:dispYear
        $targetMonth = $script:dispMonth
        $targetDay = $script:selectedDay
    } else {
        $now = Get-Date
        $targetYear = $now.Year
        $targetMonth = $now.Month
        $targetDay = $now.Day
    }

    # 固定始業・終業時刻を取得
    $startTime = $script:StartTimeMap[$shiftType]
    $endTime = $script:EndTimeMap[$shiftType]

    # 遅刻判定＆理由入力（Excel COMを開く前に実施）
    $isLate = $false
    $lateReason = $null
    $rounded = $null
    if (-not $isEstimated) {
        $now = Get-Date
        $isLateDetected = $false

        if ($shiftType -eq "深夜") {
            # 深夜: 日付またぎがあるためdatetime比較
            $targetStart = Get-Date -Year $now.Year -Month $now.Month -Day $now.Day -Hour $startTime.Hours -Minute $startTime.Minutes -Second 0
            # 現在時刻が午前（日付をまたいだ後）なら基準は前日の22:30
            if ($now.Hour -lt 12) {
                $targetStart = $targetStart.AddDays(-1)
            }
            $isLateDetected = ($now -gt $targetStart.AddMinutes(10))
        } else {
            $fixedStartMinutes = $startTime.Hours * 60 + $startTime.Minutes
            $currentMinutes = $now.Hour * 60 + $now.Minute
            $isLateDetected = ($currentMinutes -gt ($fixedStartMinutes + 10))
        }

        if ($isLateDetected) {
            $isLate = $true
            $lateReason = Show-LateReasonDialog -OwnerWindow $Window
            if ($null -eq $lateReason) {
                return
            }
            if ($shiftType -eq "深夜") {
                $rounded = Get-NightShiftRoundedQuarter -DateTime $now
            } else {
                $rounded = Get-RoundedQuarter -DateTime $now
            }
        }
    }

    # タイムシートパス取得
    $tsPath = Get-TimesheetPath -Year $targetYear -Month $targetMonth
    if ($null -eq $tsPath) {
        [System.Windows.MessageBox]::Show("タイムシートフォルダが設定されていません。`n設定タブでフォルダパスを設定してください。", "エラー", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return
    }
    if (-not (Test-Path $tsPath)) {
        [System.Windows.MessageBox]::Show("タイムシートが見つかりません。`n$tsPath", "エラー", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return
    }

    # Excel COM操作
    $excel = $null
    $workbook = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($tsPath)
        $worksheet = $workbook.Sheets.Item(1)

        # 対象行を特定
        $row = Find-DayRow -Worksheet $worksheet -Day $targetDay
        if ($row -eq -1) {
            [System.Windows.MessageBox]::Show("タイムシートに${targetDay}日の行が見つかりません。", "エラー", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return
        }

        # NumberFormat: 深夜は24h超えのため [h]:mm、それ以外は h:mm
        $timeFmt = if ($shiftType -eq "深夜") { "[h]:mm" } else { "h:mm" }

        if ($isEstimated) {
            # 想定記入: E列に出勤形態, F列に固定始業, G列に固定終業
            $cellE = $worksheet.Cells.Item($row, 5)
            $cellE.Value2 = [string]$shiftType

            $cellF = $worksheet.Cells.Item($row, 6)
            $cellF.NumberFormat = $timeFmt
            $cellF.Value2 = [double](ConvertTo-ExcelSerial -Hours $startTime.Hours -Minutes $startTime.Minutes)

            $cellG = $worksheet.Cells.Item($row, 7)
            $cellG.NumberFormat = $timeFmt
            $cellG.Value2 = [double](ConvertTo-ExcelSerial -Hours $endTime.Hours -Minutes $endTime.Minutes)
        } elseif ($isLate) {
            # 遅刻処理
            $cellE = $worksheet.Cells.Item($row, 5)
            $cellE.Value2 = [string]"遅刻"

            $cellF = $worksheet.Cells.Item($row, 6)
            $cellF.NumberFormat = $timeFmt
            $cellF.Value2 = [double](ConvertTo-ExcelSerial -Hours $rounded.Hours -Minutes $rounded.Minutes)

            $cellL = $worksheet.Cells.Item($row, 12)
            $cellL.Value2 = [string]$lateReason
        } else {
            # 通常出勤
            $cellE = $worksheet.Cells.Item($row, 5)
            $cellE.Value2 = [string]$shiftType

            $cellF = $worksheet.Cells.Item($row, 6)
            $cellF.NumberFormat = $timeFmt
            $cellF.Value2 = [double](ConvertTo-ExcelSerial -Hours $startTime.Hours -Minutes $startTime.Minutes)
        }

        $workbook.Save()

        # Teams Post判定（出勤時）
        $teamsError = $null
        $chkNoTeamsPost = $Window.FindName("ChkNoTeamsPost")
        $today = Get-Date
        $isCalendarToday = ($script:dispYear -eq $today.Year -and $script:dispMonth -eq $today.Month -and $script:selectedDay -eq $today.Day)
        $shouldPost = ((-not $chkNoTeamsPost.IsChecked) -and (-not $isEstimated) -and ($shiftType -in $script:RealtimeShiftGroups) -and $isCalendarToday)
        if ($shouldPost) {
            $radioRemote = $Window.FindName("RadioRemote")
            $workMode = if ($radioRemote.IsChecked) { "リモート" } else { "出社" }
            try {
                Send-TeamsPost -CheckType "出勤" -WorkMode $workMode -MentionData @() -Comment ""
            } catch {
                $teamsError = $_
            }

            # 出勤データCSV出力
            try {
                $csvFolder = $script:settings.attendance_data_folder
                if ([string]::IsNullOrWhiteSpace($csvFolder)) {
                    $csvFolder = Join-Path (Split-Path -Parent $PSScriptRoot) "attendance_data"
                }
                if (-not (Test-Path $csvFolder)) {
                    New-Item -Path $csvFolder -ItemType Directory -Force | Out-Null
                }
                $csvName = $script:settings.user_info.shift_display_name
                $csvShift = if ($workMode -eq "リモート") { "${shiftType}(ﾃ" } else { $shiftType }
                $csvPath = Join-Path $csvFolder "${csvName}.csv"
                $csvContent = "name,shift`r`n${csvName},${csvShift}"
                [System.IO.File]::WriteAllText($csvPath, $csvContent, (New-Object System.Text.UTF8Encoding $false))
            } catch {
                # CSV出力失敗は警告のみ
            }
        }

        $msg = "出勤を記録しました。`n${targetYear}/${targetMonth}/${targetDay} ${shiftType}"
        if ($teamsError) {
            $msg += "`n`nTeams投稿に失敗しました: $teamsError"
            [System.Windows.MessageBox]::Show($msg, "完了（Teams投稿エラー）", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        } else {
            [System.Windows.MessageBox]::Show($msg, "完了", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::None)
        }

    } catch {
        [System.Windows.MessageBox]::Show("タイムシートへの記載に失敗しました。`n$_", "エラー", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    } finally {
        # COM cleanup
        if ($workbook) {
            try { $workbook.Close($false) } catch {}
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        }
        if ($excel) {
            try { $excel.Quit() } catch {}
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

# 退勤処理
function Write-ClockOut {
    param($Window)

    # UI値取得
    $cmbShiftType = $Window.FindName("CmbShiftType")
    $shiftType = $cmbShiftType.SelectedItem
    $isNightShift = ($shiftType -eq "深夜")

    # 退勤情報ダイアログ表示（Excel COM前に表示）
    $clockOutInfo = Show-ClockOutDialog -OwnerWindow $Window -IsNightShift $isNightShift
    if ($null -eq $clockOutInfo) { return }

    # 日付取得（常にリアルタイム）
    $now = Get-Date
    $isCrossDay = ($clockOutInfo.ClockOutType -eq "crossday")

    # ターゲット日付の決定
    if ($isCrossDay) {
        if ($isNightShift) {
            # 深夜＋日跨ぎ: 2日前の行（翌々日退勤 = シフト開始日）
            $targetDate = $now.AddDays(-2)
        } else {
            # 通常シフト＋日跨ぎ: 前日の行
            $targetDate = $now.AddDays(-1)
        }
    } else {
        if ($isNightShift) {
            # 深夜＋通常退勤: 前日の行
            $targetDate = $now.AddDays(-1)
        } else {
            # 通常シフト＋通常退勤: 本日の行
            $targetDate = $now
        }
    }
    $targetYear = $targetDate.Year
    $targetMonth = $targetDate.Month
    $targetDay = $targetDate.Day

    # タイムシートパス取得
    $tsPath = Get-TimesheetPath -Year $targetYear -Month $targetMonth
    if ($null -eq $tsPath) {
        [System.Windows.MessageBox]::Show("タイムシートフォルダが設定されていません。`n設定タブでフォルダパスを設定してください。", "エラー", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return
    }
    if (-not (Test-Path $tsPath)) {
        [System.Windows.MessageBox]::Show("タイムシートが見つかりません。`n$tsPath", "エラー", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return
    }

    # Excel COM操作
    $excel = $null
    $workbook = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($tsPath)
        $worksheet = $workbook.Sheets.Item(1)

        # 対象行を特定
        $row = Find-DayRow -Worksheet $worksheet -Day $targetDay
        if ($row -eq -1) {
            [System.Windows.MessageBox]::Show("タイムシートに${targetDay}日の行が見つかりません。", "エラー", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return
        }

        # F列(始業時刻)チェック
        $startValue = $worksheet.Cells.Item($row, 6).Value2
        if ($null -eq $startValue -or $startValue -eq "") {
            [System.Windows.MessageBox]::Show("対象日の始業時間が記載されてません。`n確認して再度退勤をしてください。", "エラー", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return
        }

        # 時刻丸め＆シリアル値計算
        if ($isCrossDay) {
            # 日跨ぎ: 通常丸め + 24h加算（深夜は+48h）
            $rounded = Get-RoundedQuarter -DateTime $now
            if ($isNightShift) {
                $rounded.Hours += 48
            } else {
                $rounded.Hours += 24
            }
            $timeFmt = "[h]:mm"
        } else {
            if ($isNightShift) {
                $rounded = Get-NightShiftRoundedQuarter -DateTime $now
                $timeFmt = "[h]:mm"
            } else {
                $rounded = Get-RoundedQuarter -DateTime $now
                $timeFmt = "h:mm"
            }
        }
        $endSerial = [double](ConvertTo-ExcelSerial -Hours $rounded.Hours -Minutes $rounded.Minutes)

        # G列に終業時刻を記載
        $cellG = $worksheet.Cells.Item($row, 7)
        $cellG.NumberFormat = $timeFmt
        $cellG.Value2 = $endSerial

        # 実働時間計算 (G列 - F列)
        $workDuration = $endSerial - [double]$startValue

        # 9時間超 → K列に「客先指示」
        $nineHoursSerial = 9.0 / 24.0
        if ($workDuration -gt $nineHoursSerial) {
            $cellK = $worksheet.Cells.Item($row, 11)
            $cellK.Value2 = [string]"客先指示"
        }

        $workbook.Save()

        # Teams Post判定（退勤時）
        $teamsError = $null
        $chkNoTeamsPost = $Window.FindName("ChkNoTeamsPost")
        $shouldPost = ((-not $chkNoTeamsPost.IsChecked) -and ($shiftType -in $script:RealtimeShiftGroups))
        if ($shouldPost) {
            try {
                Send-TeamsPost -CheckType "退勤" `
                    -WorkMode "" `
                    -NextDateText $clockOutInfo.NextDate `
                    -NextShift $clockOutInfo.NextShift `
                    -NextWorkMode $clockOutInfo.NextWorkMode `
                    -MentionData $clockOutInfo.MentionIds `
                    -Comment $clockOutInfo.Comment
            } catch {
                $teamsError = $_
            }
        }

        $endTimeStr = "{0}:{1:D2}" -f $rounded.Hours, $rounded.Minutes
        $displayDate = "${targetYear}/${targetMonth}/${targetDay}"
        $msg = "退勤を記録しました。`n${displayDate} 退勤時刻: ${endTimeStr}"
        if ($teamsError) {
            $msg += "`n`nTeams投稿に失敗しました: $teamsError"
            [System.Windows.MessageBox]::Show($msg, "完了（Teams投稿エラー）", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        } else {
            [System.Windows.MessageBox]::Show($msg, "完了", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::None)
        }

    } catch {
        [System.Windows.MessageBox]::Show("タイムシートへの記載に失敗しました。`n$_", "エラー", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    } finally {
        # COM cleanup
        if ($workbook) {
            try { $workbook.Close($false) } catch {}
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        }
        if ($excel) {
            try { $excel.Quit() } catch {}
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

# 休暇グループ出勤処理
function Write-VacationClockIn {
    param($Window, [string]$ShiftType)

    $config = $script:VacationConfig[$ShiftType]

    # 日付はカレンダー選択日を使用
    $targetYear = $script:dispYear
    $targetMonth = $script:dispMonth
    $targetDay = $script:selectedDay

    # タイムシートパス取得
    $tsPath = Get-TimesheetPath -Year $targetYear -Month $targetMonth
    if ($null -eq $tsPath) {
        [System.Windows.MessageBox]::Show("タイムシートフォルダが設定されていません。`n設定タブでフォルダパスを設定してください。", "エラー", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return
    }
    if (-not (Test-Path $tsPath)) {
        [System.Windows.MessageBox]::Show("タイムシートが見つかりません。`n$tsPath", "エラー", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return
    }

    # Excel COM操作
    $excel = $null
    $workbook = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($tsPath)
        $worksheet = $workbook.Sheets.Item(1)

        # 対象行を特定
        $row = Find-DayRow -Worksheet $worksheet -Day $targetDay
        if ($row -eq -1) {
            [System.Windows.MessageBox]::Show("タイムシートに${targetDay}日の行が見つかりません。", "エラー", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return
        }

        # E列: 出勤形態
        $cellE = $worksheet.Cells.Item($row, 5)
        $cellE.Value2 = [string]$config.ShiftLabel

        # F列: 始業時刻
        if ($null -ne $config.StartTime) {
            $cellF = $worksheet.Cells.Item($row, 6)
            $cellF.NumberFormat = "h:mm"
            $cellF.Value2 = [double](ConvertTo-ExcelSerial -Hours $config.StartTime.Hours -Minutes $config.StartTime.Minutes)
        }

        # G列: 終業時刻
        if ($null -ne $config.EndTime) {
            $cellG = $worksheet.Cells.Item($row, 7)
            $cellG.NumberFormat = "h:mm"
            $cellG.Value2 = [double](ConvertTo-ExcelSerial -Hours $config.EndTime.Hours -Minutes $config.EndTime.Minutes)
        }

        # L列: 備考
        if ($null -ne $config.Remark) {
            $cellL = $worksheet.Cells.Item($row, 12)
            $cellL.Value2 = [string]$config.Remark
        }

        $workbook.Save()
        [System.Windows.MessageBox]::Show("記録しました。`n${targetYear}/${targetMonth}/${targetDay} ${ShiftType}", "完了", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::None)

    } catch {
        [System.Windows.MessageBox]::Show("タイムシートへの記載に失敗しました。`n$_", "エラー", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    } finally {
        if ($workbook) {
            try { $workbook.Close($false) } catch {}
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        }
        if ($excel) {
            try { $excel.Quit() } catch {}
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

# 備考入力ダイアログ（汎用）
function Show-RemarkInputDialog {
    param($OwnerWindow, [string]$PromptText)

    $dialogXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="備考入力" Width="450" Height="180"
        WindowStartupLocation="CenterOwner" ResizeMode="NoResize">
    <StackPanel Margin="20">
        <TextBlock x:Name="TxtPrompt" TextWrapping="Wrap" FontSize="12" Margin="0,0,0,10"/>
        <TextBox x:Name="TxtRemark" FontSize="12" Height="30" TextWrapping="Wrap"/>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,15,0,0">
            <Button x:Name="BtnOk" Content="OK" Width="80" Height="30" Margin="0,0,10,0" IsDefault="True"/>
            <Button x:Name="BtnCancel" Content="キャンセル" Width="80" Height="30" IsCancel="True"/>
        </StackPanel>
    </StackPanel>
</Window>
"@

    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($dialogXaml))
    $dialog = [System.Windows.Markup.XamlReader]::Load($reader)
    $reader.Close()

    $dialog.Owner = $OwnerWindow

    $txtPrompt = $dialog.FindName("TxtPrompt")
    $txtPrompt.Text = $PromptText

    $txtRemark = $dialog.FindName("TxtRemark")
    $btnOk = $dialog.FindName("BtnOk")
    $btnCancel = $dialog.FindName("BtnCancel")

    $resultHolder = [PSCustomObject]@{ Value = $null }

    $btnOk.Add_Click({
        $resultHolder.Value = $txtRemark.Text
        $dialog.DialogResult = $true
    }.GetNewClosure())

    $btnCancel.Add_Click({
        $dialog.DialogResult = $false
    }.GetNewClosure())

    $result = $dialog.ShowDialog()
    if ($result) {
        return $resultHolder.Value
    }
    return $null
}

# 休暇グループ（ユーザー入力型）出勤処理
function Write-VacationInputClockIn {
    param($Window, [string]$ShiftType)

    $config = $script:VacationInputConfig[$ShiftType]

    # 備考入力ダイアログ（Excel COMを開く前に表示）
    $remark = Show-RemarkInputDialog -OwnerWindow $Window -PromptText $config.DialogPrompt
    if ($null -eq $remark) {
        return
    }

    # 日付はカレンダー選択日を使用
    $targetYear = $script:dispYear
    $targetMonth = $script:dispMonth
    $targetDay = $script:selectedDay

    # タイムシートパス取得
    $tsPath = Get-TimesheetPath -Year $targetYear -Month $targetMonth
    if ($null -eq $tsPath) {
        [System.Windows.MessageBox]::Show("タイムシートフォルダが設定されていません。`n設定タブでフォルダパスを設定してください。", "エラー", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return
    }
    if (-not (Test-Path $tsPath)) {
        [System.Windows.MessageBox]::Show("タイムシートが見つかりません。`n$tsPath", "エラー", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return
    }

    # Excel COM操作
    $excel = $null
    $workbook = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($tsPath)
        $worksheet = $workbook.Sheets.Item(1)

        # 対象行を特定
        $row = Find-DayRow -Worksheet $worksheet -Day $targetDay
        if ($row -eq -1) {
            [System.Windows.MessageBox]::Show("タイムシートに${targetDay}日の行が見つかりません。", "エラー", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return
        }

        # E列: 出勤形態
        $cellE = $worksheet.Cells.Item($row, 5)
        $cellE.Value2 = [string]$config.ShiftLabel

        # L列: 備考（ユーザー入力）
        $cellL = $worksheet.Cells.Item($row, 12)
        $cellL.Value2 = [string]$remark

        $workbook.Save()
        [System.Windows.MessageBox]::Show("記録しました。`n${targetYear}/${targetMonth}/${targetDay} ${ShiftType}", "完了", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::None)

    } catch {
        [System.Windows.MessageBox]::Show("タイムシートへの記載に失敗しました。`n$_", "エラー", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    } finally {
        if ($workbook) {
            try { $workbook.Close($false) } catch {}
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        }
        if ($excel) {
            try { $excel.Quit() } catch {}
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

# 0.5日有給入力ダイアログ（始業・終業・備考）
function Show-HalfDayLeaveDialog {
    param($OwnerWindow)

    $dialogXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="0.5日有給" Width="400" Height="280"
        WindowStartupLocation="CenterOwner" ResizeMode="NoResize">
    <StackPanel Margin="20">
        <TextBlock Text="出勤時刻 (例:14:00):" FontSize="12" Margin="0,0,0,3"/>
        <TextBox x:Name="TxtStartTime" FontSize="12" Height="26" Margin="0,0,0,8"/>

        <TextBlock Text="退勤時刻 (例:18:00):" FontSize="12" Margin="0,0,0,3"/>
        <TextBox x:Name="TxtEndTime" FontSize="12" Height="26" Margin="0,0,0,8"/>

        <TextBlock Text="有給の詳細を入力してください:　(例)「体調不良の為」「私用の為」" FontSize="12" TextWrapping="Wrap" Margin="0,0,0,3"/>
        <TextBox x:Name="TxtRemark" FontSize="12" Height="30" TextWrapping="Wrap"/>

        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,15,0,0">
            <Button x:Name="BtnOk" Content="OK" Width="80" Height="30" Margin="0,0,10,0" IsDefault="True"/>
            <Button x:Name="BtnCancel" Content="キャンセル" Width="80" Height="30" IsCancel="True"/>
        </StackPanel>
    </StackPanel>
</Window>
"@

    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($dialogXaml))
    $dialog = [System.Windows.Markup.XamlReader]::Load($reader)
    $reader.Close()

    $dialog.Owner = $OwnerWindow

    $txtStart = $dialog.FindName("TxtStartTime")
    $txtEnd = $dialog.FindName("TxtEndTime")
    $txtRemark = $dialog.FindName("TxtRemark")
    $btnOk = $dialog.FindName("BtnOk")
    $btnCancel = $dialog.FindName("BtnCancel")

    $resultHolder = [PSCustomObject]@{ StartTime = $null; EndTime = $null; Remark = $null }

    $btnOk.Add_Click({
        # hh:mm または h:mm 形式チェック
        $timePattern = '^\d{1,2}:\d{2}$'
        $startText = $txtStart.Text.Trim()
        $endText = $txtEnd.Text.Trim()

        if ($startText -notmatch $timePattern) {
            [System.Windows.MessageBox]::Show("始業時刻の形式が正しくありません。`nhh:mm 形式で入力してください。", "入力エラー", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        if ($endText -notmatch $timePattern) {
            [System.Windows.MessageBox]::Show("終業時刻の形式が正しくありません。`nhh:mm 形式で入力してください。", "入力エラー", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }

        $sParts = $startText -split ':'
        $eParts = $endText -split ':'
        $resultHolder.StartTime = @{ Hours = [int]$sParts[0]; Minutes = [int]$sParts[1] }
        $resultHolder.EndTime = @{ Hours = [int]$eParts[0]; Minutes = [int]$eParts[1] }
        $resultHolder.Remark = $txtRemark.Text
        $dialog.DialogResult = $true
    }.GetNewClosure())

    $btnCancel.Add_Click({
        $dialog.DialogResult = $false
    }.GetNewClosure())

    $result = $dialog.ShowDialog()
    if ($result) {
        return $resultHolder
    }
    return $null
}

# 一括記入処理
function Write-BulkInput {
    param($Window)

    # UI値取得
    $cmbShiftType = $Window.FindName("CmbShiftType")
    $shiftType = $cmbShiftType.SelectedItem

    if ($script:dateList.Count -eq 0) {
        [System.Windows.MessageBox]::Show("日付リストが空です。", "警告", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }

    # 対象日付をソートしてコピー
    $targetDates = @($script:dateList | Sort-Object { [datetime]$_ })

    # 確認ダイアログ
    $datesStr = $targetDates -join ", "
    $confirmMsg = "以下の日付に「${shiftType}」を一括記入します。`n`n${datesStr}`n`nよろしいですか？"
    $confirmResult = [System.Windows.MessageBox]::Show($confirmMsg, "一括記入確認", [System.Windows.MessageBoxButton]::OKCancel, [System.Windows.MessageBoxImage]::None)
    if ($confirmResult -ne [System.Windows.MessageBoxResult]::OK) {
        return
    }

    # 月ごとにグループ化（タイムシートは月単位のため）
    $datesByMonth = @{}
    foreach ($dateStr in $targetDates) {
        $dt = [datetime]$dateStr
        $key = "{0:D4}{1:D2}" -f $dt.Year, $dt.Month
        if (-not $datesByMonth.ContainsKey($key)) {
            $datesByMonth[$key] = @{
                Year  = $dt.Year
                Month = $dt.Month
                Days  = @()
            }
        }
        $datesByMonth[$key].Days += $dt.Day
    }

    # 固定時刻取得（リアルタイム勤務グループの場合）
    $startTime = $script:StartTimeMap[$shiftType]
    $endTime = $script:EndTimeMap[$shiftType]
    $timeFmt = if ($shiftType -eq "深夜") { "[h]:mm" } else { "h:mm" }

    $successCount = 0
    $errorMessages = @()

    # 月ごとにExcelを開いて処理
    foreach ($monthKey in ($datesByMonth.Keys | Sort-Object)) {
        $monthData = $datesByMonth[$monthKey]

        $tsPath = Get-TimesheetPath -Year $monthData.Year -Month $monthData.Month
        if ($null -eq $tsPath) {
            $errorMessages += "$($monthData.Year)/$($monthData.Month): タイムシートフォルダ未設定"
            continue
        }
        if (-not (Test-Path $tsPath)) {
            $errorMessages += "$($monthData.Year)/$($monthData.Month): タイムシートが見つかりません"
            continue
        }

        $excel = $null
        $workbook = $null
        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            $workbook = $excel.Workbooks.Open($tsPath)
            $worksheet = $workbook.Sheets.Item(1)

            foreach ($day in $monthData.Days) {
                $row = Find-DayRow -Worksheet $worksheet -Day $day
                if ($row -eq -1) {
                    $errorMessages += "$($monthData.Year)/$($monthData.Month)/${day}: 行が見つかりません"
                    continue
                }

                if ($shiftType -in $script:RealtimeShiftGroups) {
                    # リアルタイム勤務グループ（想定記入）: E列=出勤形態, F列=固定始業, G列=固定終業
                    $cellE = $worksheet.Cells.Item($row, 5)
                    $cellE.Value2 = [string]$shiftType

                    $cellF = $worksheet.Cells.Item($row, 6)
                    $cellF.NumberFormat = $timeFmt
                    $cellF.Value2 = [double](ConvertTo-ExcelSerial -Hours $startTime.Hours -Minutes $startTime.Minutes)

                    $cellG = $worksheet.Cells.Item($row, 7)
                    $cellG.NumberFormat = $timeFmt
                    $cellG.Value2 = [double](ConvertTo-ExcelSerial -Hours $endTime.Hours -Minutes $endTime.Minutes)
                }
                elseif ($shiftType -eq "シフト休") {
                    # シフト休: E列のみ
                    $cellE = $worksheet.Cells.Item($row, 5)
                    $cellE.Value2 = [string]"シフト休"
                }
                elseif ($shiftType -eq "1.0日有給") {
                    # 1.0日有給: E列 + L列（私用の為）
                    $cellE = $worksheet.Cells.Item($row, 5)
                    $cellE.Value2 = [string]"1.0日有給"

                    $cellL = $worksheet.Cells.Item($row, 12)
                    $cellL.Value2 = [string]"私用の為"
                }

                $successCount++
            }

            $workbook.Save()

        } catch {
            $errorMessages += "$($monthData.Year)/$($monthData.Month): $_"
        } finally {
            if ($workbook) {
                try { $workbook.Close($false) } catch {}
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
            }
            if ($excel) {
                try { $excel.Quit() } catch {}
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
            }
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    }

    # 結果表示
    if ($errorMessages.Count -eq 0) {
        [System.Windows.MessageBox]::Show("一括記入が完了しました。`n${successCount}件を記録しました。", "完了", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::None)
    } else {
        $errStr = $errorMessages -join "`n"
        [System.Windows.MessageBox]::Show("一括記入が完了しました。`n成功: ${successCount}件`n`nエラー:`n${errStr}", "一括記入結果", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
    }
}

# 0.5日有給出勤処理
function Write-HalfDayLeaveClockIn {
    param($Window)

    # ダイアログ表示（Excel COMを開く前）
    $input = Show-HalfDayLeaveDialog -OwnerWindow $Window
    if ($null -eq $input) {
        return
    }

    # 日付はカレンダー選択日を使用
    $targetYear = $script:dispYear
    $targetMonth = $script:dispMonth
    $targetDay = $script:selectedDay

    # タイムシートパス取得
    $tsPath = Get-TimesheetPath -Year $targetYear -Month $targetMonth
    if ($null -eq $tsPath) {
        [System.Windows.MessageBox]::Show("タイムシートフォルダが設定されていません。`n設定タブでフォルダパスを設定してください。", "エラー", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return
    }
    if (-not (Test-Path $tsPath)) {
        [System.Windows.MessageBox]::Show("タイムシートが見つかりません。`n$tsPath", "エラー", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return
    }

    # Excel COM操作
    $excel = $null
    $workbook = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($tsPath)
        $worksheet = $workbook.Sheets.Item(1)

        # 対象行を特定
        $row = Find-DayRow -Worksheet $worksheet -Day $targetDay
        if ($row -eq -1) {
            [System.Windows.MessageBox]::Show("タイムシートに${targetDay}日の行が見つかりません。", "エラー", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return
        }

        # E列: 出勤形態
        $cellE = $worksheet.Cells.Item($row, 5)
        $cellE.Value2 = [string]"0.5日有給"

        # F列: 始業時刻
        $cellF = $worksheet.Cells.Item($row, 6)
        $cellF.NumberFormat = "h:mm"
        $cellF.Value2 = [double](ConvertTo-ExcelSerial -Hours $input.StartTime.Hours -Minutes $input.StartTime.Minutes)

        # G列: 終業時刻
        $cellG = $worksheet.Cells.Item($row, 7)
        $cellG.NumberFormat = "h:mm"
        $cellG.Value2 = [double](ConvertTo-ExcelSerial -Hours $input.EndTime.Hours -Minutes $input.EndTime.Minutes)

        # L列: 備考
        if (-not [string]::IsNullOrWhiteSpace($input.Remark)) {
            $cellL = $worksheet.Cells.Item($row, 12)
            $cellL.Value2 = [string]$input.Remark
        }

        $workbook.Save()
        [System.Windows.MessageBox]::Show("記録しました。`n${targetYear}/${targetMonth}/${targetDay} 0.5日有給", "完了", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::None)

    } catch {
        [System.Windows.MessageBox]::Show("タイムシートへの記載に失敗しました。`n$_", "エラー", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    } finally {
        if ($workbook) {
            try { $workbook.Close($false) } catch {}
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        }
        if ($excel) {
            try { $excel.Quit() } catch {}
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

# ==============================
# 退勤情報入力ダイアログ
# ==============================
function Show-ClockOutDialog {
    param($OwnerWindow, [bool]$IsNightShift = $false)

    $dialogXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="退勤情報" Width="460" Height="500"
        WindowStartupLocation="CenterOwner" ResizeMode="NoResize">
    <StackPanel Margin="20">
        <TextBlock Text="退勤種別:" FontSize="12" FontWeight="Bold" Margin="0,0,0,5"/>
        <StackPanel Orientation="Horizontal" Margin="0,0,0,12">
            <RadioButton x:Name="RadioNormal" Content="通常退勤" IsChecked="True" Margin="0,0,20,0" FontSize="12"/>
            <RadioButton x:Name="RadioCrossDay" FontSize="12"/>
        </StackPanel>

        <TextBlock Text="次回出勤日 (例: 2/16):" FontSize="12" Margin="0,0,0,3"/>
        <TextBox x:Name="TxtNextDate" FontSize="12" Height="26" Margin="0,0,0,10"/>

        <Grid Margin="0,0,0,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="12"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <StackPanel Grid.Column="0">
                <TextBlock Text="次回シフト:" FontSize="12" Margin="0,0,0,3"/>
                <ComboBox x:Name="CmbNextShift" FontSize="12" Height="28"/>
            </StackPanel>
            <StackPanel Grid.Column="2">
                <TextBlock Text="次回出勤形式:" FontSize="12" Margin="0,0,0,3"/>
                <ComboBox x:Name="CmbNextWorkMode" FontSize="12" Height="28"/>
            </StackPanel>
        </Grid>

        <TextBlock Text="コメント (任意):" FontSize="12" Margin="0,0,0,3"/>
        <TextBox x:Name="TxtComment" FontSize="12" Height="30" TextWrapping="Wrap" Margin="0,0,0,10"/>

        <TextBlock Text="メンション先:" FontSize="12" Margin="0,0,0,3"/>
        <ComboBox x:Name="CmbMention" FontSize="12" Height="28" Margin="0,0,0,18"/>

        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="BtnOk" Content="OK" Width="80" Height="30" Margin="0,0,10,0" IsDefault="True"/>
            <Button x:Name="BtnCancel" Content="キャンセル" Width="80" Height="30" IsCancel="True"/>
        </StackPanel>
    </StackPanel>
</Window>
"@

    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($dialogXaml))
    $dialog = [System.Windows.Markup.XamlReader]::Load($reader)
    $reader.Close()
    $dialog.Owner = $OwnerWindow

    # コントロール取得
    $radioNormal   = $dialog.FindName("RadioNormal")
    $radioCrossDay = $dialog.FindName("RadioCrossDay")
    $txtNextDate   = $dialog.FindName("TxtNextDate")
    $cmbNextShift  = $dialog.FindName("CmbNextShift")
    $cmbNextWorkMode = $dialog.FindName("CmbNextWorkMode")
    $txtComment    = $dialog.FindName("TxtComment")
    $cmbMention    = $dialog.FindName("CmbMention")
    $btnOk         = $dialog.FindName("BtnOk")
    $btnCancel     = $dialog.FindName("BtnCancel")

    # 日跨ぎラジオのラベル設定
    if ($IsNightShift) {
        $radioCrossDay.Content = "日を跨ぐ退勤（翌々日の退勤）"
    } else {
        $radioCrossDay.Content = "日を跨ぐ退勤"
    }

    # 次回シフト ComboBox
    foreach ($s in @("日勤", "早番", "遅番", "深夜")) {
        [void]$cmbNextShift.Items.Add($s)
    }
    $cmbNextShift.SelectedIndex = 0

    # 次回出勤形式 ComboBox
    foreach ($m in @("リモート", "出社")) {
        [void]$cmbNextWorkMode.Items.Add($m)
    }
    $cmbNextWorkMode.SelectedIndex = 0

    # メンション先 ComboBox
    [void]$cmbMention.Items.Add("")  # 空白
    $managers = $script:settings.managers
    foreach ($mgr in $managers) {
        [void]$cmbMention.Items.Add("@$($mgr.name)さん")
    }
    [void]$cmbMention.Items.Add("@All管理職")
    $cmbMention.SelectedIndex = 0

    # 結果格納
    $resultHolder = [PSCustomObject]@{
        ClockOutType = $null
        NextDate     = $null
        NextShift    = $null
        NextWorkMode = $null
        Comment      = $null
        MentionIds   = $null
    }

    $btnOk.Add_Click({
        # 次回出勤日バリデーション（M/D または MM/DD）
        $dateText = $txtNextDate.Text.Trim()
        if ($dateText -notmatch '^\d{1,2}/\d{1,2}$') {
            [System.Windows.MessageBox]::Show("次回出勤日はM/D形式で入力してください。`n例: 2/16", "入力エラー", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }

        $resultHolder.ClockOutType = if ($radioNormal.IsChecked) { "normal" } else { "crossday" }
        $resultHolder.NextDate     = $dateText
        $resultHolder.NextShift    = $cmbNextShift.SelectedItem
        $resultHolder.NextWorkMode = $cmbNextWorkMode.SelectedItem
        $resultHolder.Comment      = $txtComment.Text

        # メンション先の解決
        $mentionSelection = $cmbMention.SelectedItem
        if ([string]::IsNullOrEmpty($mentionSelection)) {
            $resultHolder.MentionIds = @()
        } elseif ($mentionSelection -eq "@All管理職") {
            $resultHolder.MentionIds = @($managers | ForEach-Object { $_.teams_principal_id })
        } else {
            $mentionName = $mentionSelection -replace '^@(.+)さん$', '$1'
            $matched = $managers | Where-Object { $_.name -eq $mentionName }
            if ($matched) {
                $resultHolder.MentionIds = @($matched.teams_principal_id)
            } else {
                $resultHolder.MentionIds = @()
            }
        }

        $dialog.DialogResult = $true
    }.GetNewClosure())

    $btnCancel.Add_Click({
        $dialog.DialogResult = $false
    }.GetNewClosure())

    $result = $dialog.ShowDialog()
    if ($result) {
        return $resultHolder
    }
    return $null
}

# ==============================
# Teams Webhook 投稿
# ==============================
function Get-ProxyCredential {
    param([string]$Path)
    
    if (Test-Path $Path) {
        try {
            return Import-Clixml -Path $Path
        } catch {
            Write-Host "保存された認証情報の読み込みに失敗しました"
            Remove-Item $Path -ErrorAction SilentlyContinue
        }
    }
    
    # 新規入力
    $cred = Get-Credential -Message "プロキシ認証情報を入力してください"
    if ($null -eq $cred) {
        throw "認証情報の入力がキャンセルされました"
    }
    $cred | Export-Clixml -Path $Path
    return $cred
}

function Invoke-RestMethodWithAutoProxy {
    param(
        [string]$Uri,
        [string]$Method,
        [byte[]]$Body,
        [string]$ContentType
    )
    
    # プロキシ検出
    $systemProxy = [System.Net.WebRequest]::GetSystemWebProxy()
    $proxyUri = $systemProxy.GetProxy($Uri)
    
    # プロキシが必要かチェック（正しい比較方法）
    $needsProxy = ($proxyUri.AbsoluteUri -ne $Uri)
    
    # 基本パラメータ
    $params = @{
        Uri = $Uri
        Method = $Method
        Body = $Body
        ContentType = $ContentType
    }
    
    # プロキシが必要な場合のみ認証情報を取得して追加
    if ($needsProxy) {
        Write-Host "プロキシ経由で接続します: $($proxyUri.AbsoluteUri)"
        $params['Proxy'] = $proxyUri.AbsoluteUri  # 文字列として渡す
        $params['ProxyCredential'] = Get-ProxyCredential -Path $credPath
    } else {
        Write-Host "プロキシなしで直接接続します"
    }
    
    # 実行
    try {
        Invoke-RestMethod @params
    } catch {
        # 407エラーかつプロキシ使用時のみリトライ
        if ($needsProxy -and $_.Exception.Response.StatusCode -eq 407) {
            Write-Host "プロキシ認証に失敗しました。認証情報を再入力してください。"
            Remove-Item $credPath -ErrorAction SilentlyContinue
            
            $params['ProxyCredential'] = Get-ProxyCredential -Path $credPath
            
            # リトライ
            try {
                Invoke-RestMethod @params
            } catch {
                Write-Host "リトライも失敗しました: $($_.Exception.Message)"
                throw
            }
        } else {
            throw
        }
    }
}

function Send-TeamsPost {
    param(
        [string]$CheckType,
        [string]$WorkMode       = "",
        [string]$NextDateText   = "",
        [string]$NextShift      = "",
        [string]$NextWorkMode   = "",
        [array]$MentionData     = @(),
        [string]$Comment        = ""
    )

    $webhookUrl = $script:settings.teams_workflow.webhook_url
    if ([string]::IsNullOrWhiteSpace($webhookUrl)) {
        throw "WebhookURLが設定されていません。"
    }

    $userName = $script:settings.user_info.full_name
    $userId   = $script:settings.user_info.teams_principal_id

    # column_obj
    $columnObj = @{
        type  = "Column"
        width = "stretch"
        items = @(
            @{
                type    = "TextBlock"
                text    = "${userName}が${CheckType}しました"
                size    = "Medium"
                wrap    = $true
                weight  = "Bolder"
                verticalContentAlignment = "Center"
            }
        )
    }

    # message_obj
    if ($CheckType -eq "出勤") {
        $messageObj = @{
            type    = "TextBlock"
            text    = "業務を開始します(${WorkMode})"
            size    = "Medium"
            wrap    = $true
            spacing = "None"
        }
    } else {
        $messageObj = @(
            @{
                type    = "TextBlock"
                text    = "退勤します。次回は${NextDateText} ${NextWorkMode}(${NextShift})です。"
                size    = "Medium"
                wrap    = $true
                spacing = "None"
            },
            @{
                type    = "TextBlock"
                text    = "お疲れさまでした。"
                wrap    = $true
                spacing = "None"
            }
        )
    }

    # comment_obj
    if ([string]::IsNullOrWhiteSpace($Comment)) {
        $commentObj = @{}
    } elseif ($MentionData -and $MentionData.Count -gt 0) {
        $commentObj = @{
            type    = "TextBlock"
            text    = "コメント: ${Comment}"
            wrap    = $true
            spacing = "None"
        }
    } else {
        $commentObj = @{
            type      = "TextBlock"
            text      = "コメント: ${Comment}"
            wrap      = $true
            spacing   = "Small"
            separator = $true
        }
    }

    # payload 組み立て
    $mentionArr = if ($MentionData -and $MentionData.Count -gt 0) { $MentionData } else { @() }
    $payload = @{
        mention_data = $mentionArr
        userId       = $userId
        column       = (ConvertTo-Json -InputObject $columnObj -Depth 10 -Compress)
        message      = (ConvertTo-Json -InputObject $messageObj -Depth 10 -Compress)
        comment      = (ConvertTo-Json -InputObject $commentObj -Depth 10 -Compress)
    }

    $body = ConvertTo-Json -InputObject $payload -Depth 10
    # デバッグ用: Postデータをファイル出力
    #$debugPath = Join-Path $script:settings.timesheet_folder "teams_post_debug.json"
    #$body | Out-File -FilePath $debugPath -Encoding utf8 -Force

    #Invoke-RestMethod -Uri $webhookUrl -Method Post -Body ([System.Text.Encoding]::UTF8.GetBytes($body)) -ContentType "application/json; charset=utf-8"
    $credPath = "$env:USERPROFILE\.proxy_cred.xml"
    try {
        Invoke-RestMethodWithAutoProxy `
            -Uri $webhookUrl `
            -Method Post `
            -Body ([System.Text.Encoding]::UTF8.GetBytes($body)) `
            -ContentType "application/json; charset=utf-8"

        Write-Host "送信成功"
    } catch {
        Write-Host "送信失敗: $($_.Exception.Message)"
        exit 1
    }

}


