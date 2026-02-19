# ==============================
# タイムシート アクション（出退勤・一括記入）
# ==============================

# 出勤処理
function Write-ClockIn {
    param($Window)

    # --- ローディング開始 ---
    $btnClockIn = $Window.FindName("BtnClockIn")
    $script:processingOverlay = $Window.FindName("ProcessingOverlay")
    $btnClockIn.IsEnabled = $false
    $btnClockIn.Content = "処理中..."
    if ($script:processingOverlay) { $script:processingOverlay.Visibility = [System.Windows.Visibility]::Visible }
    $Window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Render, [action]{})

    $excel = $null
    $workbook = $null
    try {
        # UI値取得
        $cmbShiftType = $Window.FindName("CmbShiftType")
        $chkEstimatedInput = $Window.FindName("ChkEstimatedInput")
        $shiftType = $cmbShiftType.SelectedItem

        # 業務形態未選択チェック
        if ($null -eq $shiftType) {
            [System.Windows.MessageBox]::Show("業務形態を選択してください。", "警告", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }

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

        $timeFmt = if ($shiftType -eq "深夜") { "[h]:mm" } else { "h:mm" }

        if ($isEstimated) {
            # 想定記入: E列に出勤形態, F列に固定始業, G列に固定終業
            $cellE = $worksheet.Cells.Item($row, 5)
            $cellE.Value2 = [string]$shiftType

            $cellF = $worksheet.Cells.Item($row, 6)
            $cellF.Value2 = [double](ConvertTo-ExcelSerial -Hours $startTime.Hours -Minutes $startTime.Minutes)

            $cellG = $worksheet.Cells.Item($row, 7)
            $cellG.Value2 = [double](ConvertTo-ExcelSerial -Hours $endTime.Hours -Minutes $endTime.Minutes)
        } elseif ($isLate) {
            # 遅刻処理
            $cellE = $worksheet.Cells.Item($row, 5)
            $cellE.Value2 = [string]"遅刻"

            $cellF = $worksheet.Cells.Item($row, 6)
            $cellF.Value2 = [double](ConvertTo-ExcelSerial -Hours $rounded.Hours -Minutes $rounded.Minutes)

            $cellL = $worksheet.Cells.Item($row, 12)
            $cellL.Value2 = [string]$lateReason
        } else {
            # 通常出勤
            $cellE = $worksheet.Cells.Item($row, 5)
            $cellE.Value2 = [string]$shiftType

            $cellF = $worksheet.Cells.Item($row, 6)
            $cellF.Value2 = [double](ConvertTo-ExcelSerial -Hours $startTime.Hours -Minutes $startTime.Minutes)
        }

        $workbook.Save()

        # Teams Post判定 / CSV出力判定（出勤時）
        $teamsError = $null
        $chkNoTeamsPost = $Window.FindName("ChkNoTeamsPost")
        $today = Get-Date
        $isCalendarToday = ($script:dispYear -eq $today.Year -and $script:dispMonth -eq $today.Month -and $script:selectedDay -eq $today.Day)
        # TeamsPostチェックに関係なく、リアルタイム勤務かつ今日の場合はCSV出力する
        $shouldPost      = ((-not $chkNoTeamsPost.IsChecked) -and (-not $isEstimated) -and ($shiftType -in $script:RealtimeShiftGroups) -and $isCalendarToday)
        $shouldOutputCsv = ((-not $isEstimated) -and ($shiftType -in $script:RealtimeShiftGroups) -and $isCalendarToday)

        if ($shouldPost -or $shouldOutputCsv) {
            $radioRemote = $Window.FindName("RadioRemote")
            $workMode = if ($radioRemote.IsChecked) { "リモート" } else { "出社" }

            if ($shouldPost) {
                try {
                    Send-TeamsPost -CheckType "出勤" -WorkMode $workMode -MentionData @() -Comment ""
                } catch {
                    $teamsError = $_
                }
            }

            # 出勤データCSV出力（TeamsPostなしでも出力する）
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
        # --- ローディング終了（常に実行）---
        if ($script:processingOverlay) { $script:processingOverlay.Visibility = [System.Windows.Visibility]::Collapsed }
        $btnClockIn.IsEnabled = $true
        $btnClockIn.Content = "出 勤"
    }
}

# 退勤処理
function Write-ClockOut {
    param($Window)

    # --- ローディング開始 ---
    $btnClockOut = $Window.FindName("BtnClockOut")
    $script:processingOverlay = $Window.FindName("ProcessingOverlay")
    $btnClockOut.IsEnabled = $false
    $btnClockOut.Content = "処理中..."
    if ($script:processingOverlay) { $script:processingOverlay.Visibility = [System.Windows.Visibility]::Visible }
    $Window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Render, [action]{})

    $excel = $null
    $workbook = $null
    try {
        # UI値取得
        $cmbShiftType = $Window.FindName("CmbShiftType")
        $shiftType = $cmbShiftType.SelectedItem
        $isNightShift = ($shiftType -eq "深夜")

        # 退勤情報ダイアログ表示（Excel COM前に表示）
        $clockOutInfo = Show-ClockOutDialog -OwnerWindow $Window -IsNightShift $isNightShift
        if ($null -eq $clockOutInfo) {
            return
        }

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
        # --- ローディング終了（常に実行）---
        if ($script:processingOverlay) { $script:processingOverlay.Visibility = [System.Windows.Visibility]::Collapsed }
        $btnClockOut.IsEnabled = $true
        $btnClockOut.Content = "退 勤"
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
            $cellF.Value2 = [double](ConvertTo-ExcelSerial -Hours $config.StartTime.Hours -Minutes $config.StartTime.Minutes)
        }

        # G列: 終業時刻
        if ($null -ne $config.EndTime) {
            $cellG = $worksheet.Cells.Item($row, 7)
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

    # --- ローディング開始 ---
    $btnBulkInput = $Window.FindName("BtnBulkInput")
    $script:processingOverlay = $Window.FindName("ProcessingOverlay")
    $btnBulkInput.IsEnabled = $false
    if ($script:processingOverlay) { $script:processingOverlay.Visibility = [System.Windows.Visibility]::Visible }
    $Window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Render, [action]{})

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
                    $cellF.Value2 = [double](ConvertTo-ExcelSerial -Hours $startTime.Hours -Minutes $startTime.Minutes)

                    $cellG = $worksheet.Cells.Item($row, 7)
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

    # --- ローディング終了 ---
    if ($script:processingOverlay) { $script:processingOverlay.Visibility = [System.Windows.Visibility]::Collapsed }
    $btnBulkInput.IsEnabled = $true

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
        $cellF.Value2 = [double](ConvertTo-ExcelSerial -Hours $input.StartTime.Hours -Minutes $input.StartTime.Minutes)

        # G列: 終業時刻
        $cellG = $worksheet.Cells.Item($row, 7)
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
