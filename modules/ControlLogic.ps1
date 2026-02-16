# ==============================
# コントロールパネルロジック (テーマ対応)
# ==============================

# 一括記入ボタン＋出勤/退勤ボタンの有効/無効を更新する関数
function Update-BulkInputState {
    $w = $script:controlWindow
    if ($null -eq $w) { return }

    $btnBulkInput    = $w.FindName("BtnBulkInput")
    $bulkInputBorder = $w.FindName("BulkInputBorder")
    $cmbShiftType    = $w.FindName("CmbShiftType")
    $chkEstimated    = $w.FindName("ChkEstimatedInput")
    $btnClockIn      = $w.FindName("BtnClockIn")
    $btnClockOut     = $w.FindName("BtnClockOut")
    $clockInBorder   = $w.FindName("ClockInBorder")
    $clockOutBorder  = $w.FindName("ClockOutBorder")

    $hasDateList = ($script:dateList.Count -gt 0)
    $selected    = $cmbShiftType.SelectedItem
    $isEstimated = $chkEstimated.IsChecked

    $realtimeGroups    = $script:RealtimeShiftGroups
    $bulkNoCheckGroups = @("シフト休", "1.0日有給")

    $canBulk = $false
    if ($hasDateList) {
        if (($selected -in $realtimeGroups) -and $isEstimated) {
            $canBulk = $true
        }
        if ($selected -in $bulkNoCheckGroups) {
            $canBulk = $true
        }
    }

    # 一括記入ボタン
    $btnBulkInput.IsEnabled = $canBulk
    $bulkInputBorder.Opacity = if ($canBulk) { 1.0 } else { 0.4 }

    # 出勤/退勤ボタン制御
    if ($canBulk) {
        $btnClockIn.IsEnabled = $false
        $clockInBorder.Opacity = 0.4
        $btnClockIn.Content = "一括List選択中"
        $btnClockOut.IsEnabled = $false
        $clockOutBorder.Opacity = 0.4
    } else {
        $isVacation = ($selected -in $script:VacationGroup)
        $btnClockIn.IsEnabled = $true
        $clockInBorder.Opacity = 1.0
        if ($isVacation) {
            $btnClockIn.Content = $selected
        } else {
            $btnClockIn.Content = "出 勤"
        }
        $btnClockOut.IsEnabled = (-not $isVacation)
        $clockOutBorder.Opacity = if ($isVacation) { 0.4 } else { 1.0 }
    }
}

function Initialize-ControlPanel {
    param($window)

    # ウィンドウ参照を保存（Update-BulkInputState で使用）
    $script:controlWindow = $window

    # コントロール取得
    $txtDate = $window.FindName("TxtDate")
    $script:txtTime = $window.FindName("TxtTime")
    $btnTimesheet = $window.FindName("BtnTimesheet")
    $cmbShiftType = $window.FindName("CmbShiftType")
    $radioRemote = $window.FindName("RadioRemote")
    $radioOffice = $window.FindName("RadioOffice")
    $chkNoTeamsPost = $window.FindName("ChkNoTeamsPost")
    $chkEstimatedInput = $window.FindName("ChkEstimatedInput")
    $btnClockIn = $window.FindName("BtnClockIn")
    $btnClockOut = $window.FindName("BtnClockOut")
    $script:txtDateList = $window.FindName("TxtDateList")
    $btnAddDate = $window.FindName("BtnAddDate")
    $btnClearDates = $window.FindName("BtnClearDates")
    $btnBulkInput = $window.FindName("BtnBulkInput")

    # Border取得（シャドウアニメーション用）
    $clockInBorder = $window.FindName("ClockInBorder")
    $clockOutBorder = $window.FindName("ClockOutBorder")
    $addDateBorder = $window.FindName("AddDateBorder")
    $clearDatesBorder = $window.FindName("ClearDatesBorder")
    $bulkInputBorder = $window.FindName("BulkInputBorder")
    $timesheetBtnBorder = $window.FindName("TimesheetBtnBorder")

    # 日付表示
    $txtDate.Text = (Get-Date).ToString("yyyy/M/d (ddd)")

    # 時計表示
    $script:txtTime.Text = (Get-Date).ToString("HH:mm:ss")

    # DispatcherTimerで時計更新
    $script:timer = New-Object System.Windows.Threading.DispatcherTimer
    $script:timer.Interval = [TimeSpan]::FromSeconds(1)
    $script:timer.Add_Tick({
        $script:txtTime.Text = (Get-Date).ToString("HH:mm:ss")
    })
    $script:timer.Start()

    # 出勤形態コンボボックス
    foreach ($shiftType in $script:shiftTypes) {
        [void]$cmbShiftType.Items.Add($shiftType)
    }
    $cmbShiftType.SelectedIndex = 0

    # 一括記入ボタンのデフォルトグレーアウト
    $btnBulkInput.IsEnabled = $false
    $bulkInputBorder.Opacity = 0.4

    # ComboBox選択変更時: 休暇グループならUI制御 + 一括記入状態更新
    $vacationGroup = $script:VacationGroup
    $cmbShiftType.Add_SelectionChanged({
        $selected = $cmbShiftType.SelectedItem
        $isVacation = ($selected -in $vacationGroup)

        # オプション欄のグレーアウト
        $chkNoTeamsPost.IsEnabled = (-not $isVacation)
        $chkEstimatedInput.IsEnabled = (-not $isVacation)

        # 出勤形式欄のグレーアウト
        $radioRemote.IsEnabled = (-not $isVacation)
        $radioOffice.IsEnabled = (-not $isVacation)

        # 一括記入 + 出勤/退勤ボタン状態更新
        Update-BulkInputState
    }.GetNewClosure())

    # 想定記入チェックボックス変更時
    $chkEstimatedInput.Add_Checked({
        Update-BulkInputState
    })
    $chkEstimatedInput.Add_Unchecked({
        Update-BulkInputState
    })

    # 日付リストTextBox変更時: 追加/クリアどちらでも発火
    $script:txtDateList.Add_TextChanged({
        Update-BulkInputState
    })

    # 関数参照を変数に保存（クロージャーで使用するため）
    $writeClockInFunc = ${function:Write-ClockIn}
    $writeClockOutFunc = ${function:Write-ClockOut}
    $writeBulkInputFunc = ${function:Write-BulkInput}

    # Timesheetボタン
    $btnTimesheet.Add_Click({
        [System.Windows.MessageBox]::Show("Timesheet機能は未実装です。", "情報", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
    })

    # 出勤ボタン
    $btnClockIn.Add_Click({
        & $writeClockInFunc -Window $window
    }.GetNewClosure())

    # 退勤ボタン
    $btnClockOut.Add_Click({
        & $writeClockOutFunc -Window $window
    }.GetNewClosure())

    # 追加ボタン
    $btnAddDate.Add_Click({
        Add-DateToList -Day $script:selectedDay
    })

    # クリアボタン
    $btnClearDates.Add_Click({
        $script:dateList.Clear()
        $script:txtDateList.Text = ""
    })

    # 一括記入ボタン
    $btnBulkInput.Add_Click({
        & $writeBulkInputFunc -Window $window
    }.GetNewClosure())

    # ボタン+Border ホバーアニメーション設定（translateY + シャドウ変更）
    $script:ApplyButtonHoverAnimation = {
        param($button, $borderElement, [string]$shadowKey, [string]$hoverShadowKey)
    
        # New-ShadowEffect関数への参照を変数に保存
        $newShadowFunc = ${function:New-ShadowEffect}
    
        # TranslateTransformの準備（ボタン自体に適用）
        $button.RenderTransformOrigin = New-Object System.Windows.Point(0.5, 0.5)
        $translateTransform = New-Object System.Windows.Media.TranslateTransform(0, 0)
        $button.RenderTransform = $translateTransform
    
        # MouseEnter
        $button.Add_MouseEnter({
            param($sender, $e)
    
            $currentTheme = if ($script:settings.theme) { $script:settings.theme } else { "light" }
            $theme = $global:ThemeColors[$currentTheme]
    
            # Y軸移動アニメーション
            $moveAnim = New-Object System.Windows.Media.Animation.DoubleAnimation
            $moveAnim.To = -2
            $moveAnim.Duration = [System.Windows.Duration]::new([System.TimeSpan]::FromMilliseconds(300))
            $moveAnim.EasingFunction = New-Object System.Windows.Media.Animation.CubicEase
            $sender.RenderTransform.BeginAnimation([System.Windows.Media.TranslateTransform]::YProperty, $moveAnim)
    
            # シャドウ変更（Borderがある場合）
            if ($borderElement -and $hoverShadowKey -and $theme[$hoverShadowKey]) {
                $borderElement.Effect = & $newShadowFunc $theme[$hoverShadowKey]
            }
        }.GetNewClosure())

        # MouseLeave
        $button.Add_MouseLeave({
            param($sender, $e)

            $currentTheme = if ($script:settings.theme) { $script:settings.theme } else { "light" }
            $theme = $global:ThemeColors[$currentTheme]

            # Y軸を元に戻す
            $moveAnim = New-Object System.Windows.Media.Animation.DoubleAnimation
            $moveAnim.To = 0
            $moveAnim.Duration = [System.Windows.Duration]::new([System.TimeSpan]::FromMilliseconds(300))
            $moveAnim.EasingFunction = New-Object System.Windows.Media.Animation.CubicEase
            $sender.RenderTransform.BeginAnimation([System.Windows.Media.TranslateTransform]::YProperty, $moveAnim)

            # シャドウを元に戻す
            if ($borderElement -and $shadowKey -and $theme[$shadowKey]) {
                $borderElement.Effect = & $newShadowFunc $theme[$shadowKey]
            }
        }.GetNewClosure())
    }

    # ボタンにホバーアニメーション＋シャドウアニメーション適用
    & $script:ApplyButtonHoverAnimation $btnClockIn $clockInBorder "BtnSuccessShadow" "BtnSuccessHoverShadow"
    & $script:ApplyButtonHoverAnimation $btnClockOut $clockOutBorder "BtnDangerShadow" "BtnDangerHoverShadow"
    & $script:ApplyButtonHoverAnimation $btnAddDate $addDateBorder "BtnInfoShadow" "BtnInfoHoverShadow"
    & $script:ApplyButtonHoverAnimation $btnClearDates $clearDatesBorder "BtnSecondaryShadow" "BtnSecondaryHoverShadow"
    & $script:ApplyButtonHoverAnimation $btnBulkInput $bulkInputBorder "BtnWarningShadow" "BtnWarningHoverShadow"
    & $script:ApplyButtonHoverAnimation $btnTimesheet $timesheetBtnBorder $null $null
}
