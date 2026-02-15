# ==============================
# カレンダー
# ==============================

function Initialize-Calendar {
    param($window)

    # コントロール取得
    $btnPrevMonth = $window.FindName("BtnPrevMonth")
    $btnNextMonth = $window.FindName("BtnNextMonth")
    $script:txtMonthTitle = $window.FindName("TxtMonthTitle")
    $script:dayOfWeekHeader = $window.FindName("DayOfWeekHeader")
    $script:calendarGrid = $window.FindName("CalendarGrid")

    # カレンダーコンテナ
    $script:calendarContainer = $window.FindName("CalendarContainer")

    # 選択中のボタンを保持
    $script:selectedButton = $null
    $script:selectedBrush = $null
    $script:selectedBorderBrush = $null

    # 曜日ヘッダー作成関数
    $script:RebuildWeekdayHeader = {
        $currentTheme = if ($script:settings.theme) { $script:settings.theme } else { "light" }
        $theme = $global:ThemeColors[$currentTheme]

        $script:dayOfWeekHeader.Children.Clear()

        $dayNames = @("日", "月", "火", "水", "木", "金", "土")

        for ($i = 0; $i -lt 7; $i++) {
            $tb = New-Object System.Windows.Controls.TextBlock
            $tb.Text = $dayNames[$i]
            $tb.FontSize = 11
            $tb.FontWeight = "Bold"
            $tb.HorizontalAlignment = "Center"
            $tb.VerticalAlignment = "Center"
            $tb.Padding = New-Object System.Windows.Thickness(0, 10, 0, 0)

            # 曜日ごとの色
            if ($i -eq 0) {
                # 日曜日
                $tb.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.CalHolidayText))
            }
            elseif ($i -eq 6) {
                # 土曜日
                $tb.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.CalSaturdayText))
            }
            else {
                # 平日
                $tb.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.CalWeekdayText))
            }

            $tb.TextAlignment = "Center"
            $border = New-Object System.Windows.Controls.Border
            $border.Child = $tb
            [void]$script:dayOfWeekHeader.Children.Add($border)
        }
    }

    # 曜日ヘッダー初期構築
    & $script:RebuildWeekdayHeader

    # カレンダー描画関数
    $script:DrawCalendar = {
        $script:calendarGrid.Children.Clear()
        $script:selectedButton = $null

        $currentTheme = if ($script:settings.theme) { $script:settings.theme } else { "light" }
        $theme = $global:ThemeColors[$currentTheme]

        # 曜日ヘッダー再構築
        & $script:RebuildWeekdayHeader

        $y = $script:dispYear
        $m = $script:dispMonth
        $firstDay = Get-Date -Year $y -Month $m -Day 1
        $daysInMonth = [DateTime]::DaysInMonth($y, $m)
        $todayDate = Get-Date

        # 月タイトル更新
        $script:txtMonthTitle.Text = "${y}年 ${m}月"
        $script:txtMonthTitle.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.TextPrimary))

        # 祝日取得
        $holidays = Get-JapaneseHolidays -Year $y

        # 必要な行数を動的に計算
        $startDow = [int]$firstDay.DayOfWeek
        $totalCells = $startDow + $daysInMonth
        $rowCount = [math]::Ceiling($totalCells / 7)
        $script:calendarGrid.Rows = $rowCount

        # 前月の日付（フィラー）
        if ($startDow -gt 0) {
            $prevMonth = $firstDay.AddMonths(-1)
            $prevDaysInMonth = [DateTime]::DaysInMonth($prevMonth.Year, $prevMonth.Month)
            for ($i = 0; $i -lt $startDow; $i++) {
                $prevD = $prevDaysInMonth - $startDow + 1 + $i
                $tb = New-Object System.Windows.Controls.TextBlock
                $tb.Text = "$prevD"
                $tb.FontSize = 11
                $tb.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.TextDisabled))
                $tb.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.InputBg))
                $tb.HorizontalAlignment = "Stretch"
                $tb.VerticalAlignment = "Stretch"
                $tb.TextAlignment = "Center"
                $tb.Padding = New-Object System.Windows.Thickness(0, 10, 0, 0)
                [void]$script:calendarGrid.Children.Add($tb)
            }
        }

        # 当月の日付セル
        for ($d = 1; $d -le $daysInMonth; $d++) {
            $col = ($startDow + $d - 1) % 7

            $isToday = ($d -eq $todayDate.Day -and $y -eq $todayDate.Year -and $m -eq $todayDate.Month)
            $isSelected = ($d -eq $script:selectedDay)
            $holidayKey = "$y/$m/$d"
            $isHoliday = $holidays.ContainsKey($holidayKey)
            $isSunday = ($col -eq 0)
            $isSaturday = ($col -eq 6)

            # ボーダーで角丸ボタンを作成
            $border = New-Object System.Windows.Controls.Border
            $border.CornerRadius = $theme.CornerRadiusSmall
            $border.BorderThickness = if ($currentTheme -eq "dark") { 1 } else { 1 }
            $border.Margin = New-Object System.Windows.Thickness(2)
            $border.Cursor = "Hand"
            $border.Tag = $d

            # ScaleTransformの準備（ホバーアニメーション用）
            $border.RenderTransformOrigin = New-Object System.Windows.Point(0.5, 0.5)
            $scaleTransform = New-Object System.Windows.Media.ScaleTransform(1, 1)
            $border.RenderTransform = $scaleTransform

            # まず通常時の背景・ボーダー・テキスト色を決定（選択状態に関係なく）
            if ($isToday) {
                $normalBg = New-GradientBrush -Colors $theme.CalTodayBg
                $normalBorderBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.CalTodayBg[0]))
                $normalTextColor = $theme.CalTodayText
            }
            elseif ($isHoliday) {
                $normalBg = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.CalHolidayBg))
                $normalBorderBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.InputBorder))
                $normalTextColor = $theme.CalHolidayText
            }
            elseif ($isSunday) {
                $normalBg = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.CalHolidayBg))
                $normalBorderBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.InputBorder))
                $normalTextColor = $theme.CalHolidayText
            }
            elseif ($isSaturday) {
                $normalBg = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.CalSaturdayBg))
                $normalBorderBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.InputBorder))
                $normalTextColor = $theme.CalSaturdayText
            }
            else {
                $normalBg = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.CalDayBg))
                $normalBorderBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.InputBorder))
                $normalTextColor = $theme.CalDayText
            }

            # 背景色とボーダー色の設定（選択中なら選択色を上書き適用）
            if ($isSelected) {
                $border.Background = New-GradientBrush -Colors $theme.CalSelectedBg
                $border.BorderBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.CalSelectedBg[0]))
                $textColor = $theme.CalSelectedText
            }
            else {
                $border.Background = $normalBg
                $border.BorderBrush = $normalBorderBrush
                $textColor = $normalTextColor
            }

            # 元の背景・ボーダーを保存（常に通常色を保存。選択色ではなく本来の色）
            $border.Tag = @{
                Day = $d
                OriginalBackground = $normalBg
                OriginalBorderBrush = $normalBorderBrush
                IsSelected = $isSelected
                IsToday = $isToday
                IsHoliday = $isHoliday
                IsSunday = $isSunday
                IsSaturday = $isSaturday
            }

            # テキスト
            $tb = New-Object System.Windows.Controls.TextBlock
            $tb.Text = "$d"
            $tb.FontSize = 12
            $tb.HorizontalAlignment = "Center"
            $tb.VerticalAlignment = "Center"
            $tb.Padding = New-Object System.Windows.Thickness(0, 10, 0, 0)
            $tb.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($textColor))

            if ($isToday -or $isSelected) {
                $tb.FontWeight = "Bold"
            }

            $border.Child = $tb

            # 祝日ツールチップ
            if ($isHoliday) {
                $tooltip = New-Object System.Windows.Controls.ToolTip
                $tooltip.Content = $holidays[$holidayKey]
                $border.ToolTip = $tooltip
            }

            # ホバーアニメーション
            $border.Add_MouseEnter({
                param($sender, $e)

                $currentTheme = if ($script:settings.theme) { $script:settings.theme } else { "light" }
                $theme = $global:ThemeColors[$currentTheme]

                # スケールアニメーション
                $scaleAnim = New-Object System.Windows.Media.Animation.DoubleAnimation
                $scaleAnim.To = 1.08
                $scaleAnim.Duration = [System.Windows.Duration]::new([System.TimeSpan]::FromMilliseconds(200))
                $scaleAnim.EasingFunction = New-Object System.Windows.Media.Animation.CubicEase
                $sender.RenderTransform.BeginAnimation([System.Windows.Media.ScaleTransform]::ScaleXProperty, $scaleAnim)
                $sender.RenderTransform.BeginAnimation([System.Windows.Media.ScaleTransform]::ScaleYProperty, $scaleAnim)

                # 背景色変更（選択中以外）
                if (-not $sender.Tag.IsSelected) {
                    $sender.Background = New-GradientBrush -Colors $theme.CalDayHoverBg
                    $sender.Child.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.CalDayHoverText))
                }
            })

            $border.Add_MouseLeave({
                param($sender, $e)

                # スケールを戻す
                $scaleAnim = New-Object System.Windows.Media.Animation.DoubleAnimation
                $scaleAnim.To = 1.0
                $scaleAnim.Duration = [System.Windows.Duration]::new([System.TimeSpan]::FromMilliseconds(200))
                $scaleAnim.EasingFunction = New-Object System.Windows.Media.Animation.CubicEase
                $sender.RenderTransform.BeginAnimation([System.Windows.Media.ScaleTransform]::ScaleXProperty, $scaleAnim)
                $sender.RenderTransform.BeginAnimation([System.Windows.Media.ScaleTransform]::ScaleYProperty, $scaleAnim)

                # 背景色を元に戻す
                if (-not $sender.Tag.IsSelected) {
                    $sender.Background = $sender.Tag.OriginalBackground
                    $sender.BorderBrush = $sender.Tag.OriginalBorderBrush

                    # テキスト色も元に戻す
                    $currentTheme = if ($script:settings.theme) { $script:settings.theme } else { "light" }
                    $theme = $global:ThemeColors[$currentTheme]

                    if ($sender.Tag.IsToday) {
                        $sender.Child.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.CalTodayText))
                    }
                    elseif ($sender.Tag.IsHoliday -or $sender.Tag.IsSunday) {
                        $sender.Child.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.CalHolidayText))
                    }
                    elseif ($sender.Tag.IsSaturday) {
                        $sender.Child.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.CalSaturdayText))
                    }
                    else {
                        $sender.Child.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.CalDayText))
                    }
                }
            })

            # クリック処理: シングルクリックで選択、ダブルクリックで日付追加
            $border.Add_MouseLeftButtonDown({
                param($sender, $e)

                if ($e.ClickCount -eq 2) {
                    # ダブルクリック: 日付リストに追加
                    $day = $sender.Tag.Day
                    Add-DateToList -Day $day
                }
                else {
                    # シングルクリック: 選択
                    $currentTheme = if ($script:settings.theme) { $script:settings.theme } else { "light" }
                    $theme = $global:ThemeColors[$currentTheme]

                    # 前の選択を元に戻す
                    if ($script:selectedButton -ne $null) {
                        $script:selectedButton.Background = $script:selectedBrush
                        $script:selectedButton.BorderBrush = $script:selectedBorderBrush
                        $script:selectedButton.Tag.IsSelected = $false

                        # テキスト色・FontWeightも元に戻す
                        $prev = $script:selectedButton.Tag
                        if ($prev.IsToday) {
                            $script:selectedButton.Child.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.CalTodayText))
                            $script:selectedButton.Child.FontWeight = "Bold"
                        }
                        elseif ($prev.IsHoliday -or $prev.IsSunday) {
                            $script:selectedButton.Child.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.CalHolidayText))
                            $script:selectedButton.Child.FontWeight = "Normal"
                        }
                        elseif ($prev.IsSaturday) {
                            $script:selectedButton.Child.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.CalSaturdayText))
                            $script:selectedButton.Child.FontWeight = "Normal"
                        }
                        else {
                            $script:selectedButton.Child.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.CalDayText))
                            $script:selectedButton.Child.FontWeight = "Normal"
                        }
                    }

                    # 新しい選択を適用
                    $script:selectedBrush = $sender.Tag.OriginalBackground
                    $script:selectedBorderBrush = $sender.Tag.OriginalBorderBrush
                    $script:selectedDay = $sender.Tag.Day
                    $script:selectedButton = $sender
                    $sender.Tag.IsSelected = $true

                    $sender.Background = New-GradientBrush -Colors $theme.CalSelectedBg
                    $sender.BorderBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.CalSelectedBg[0]))
                    $sender.Child.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.CalSelectedText))
                    $sender.Child.FontWeight = "Bold"
                }

                $e.Handled = $true
            })

            # 選択中の日をマーク
            if ($isSelected) {
                $script:selectedButton = $border
                $script:selectedBrush = $border.Tag.OriginalBackground
                $script:selectedBorderBrush = $border.Tag.OriginalBorderBrush
            }

            [void]$script:calendarGrid.Children.Add($border)
        }

        # 次月の日付（フィラー）
        $lastCol = ($startDow + $daysInMonth - 1) % 7
        if ($lastCol -lt 6) {
            $nextD = 1
            for ($i = $lastCol + 1; $i -le 6; $i++) {
                $tb = New-Object System.Windows.Controls.TextBlock
                $tb.Text = "$nextD"
                $tb.FontSize = 11
                $tb.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.TextDisabled))
                $tb.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.InputBg))
                $tb.HorizontalAlignment = "Stretch"
                $tb.VerticalAlignment = "Stretch"
                $tb.TextAlignment = "Center"
                $tb.Padding = New-Object System.Windows.Thickness(0, 10, 0, 0)
                [void]$script:calendarGrid.Children.Add($tb)
                $nextD++
            }
        }
    }

    # ナビゲーションボタンにテーマ背景＋ホバーアニメーション（translateY + シャドウ）
    $script:ApplyNavButtonThemeAndHover = {
        param($button)

        $currentTheme = if ($script:settings.theme) { $script:settings.theme } else { "light" }
        $theme = $global:ThemeColors[$currentTheme]

        # 背景グラデーション適用（ライトモード: グラデーション、ダーク: transparent）
        $button.Background = New-GradientBrush -Colors $theme.CalNavBg
        $button.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.CalNavText))

        # 初期シャドウ
        if ($theme.NavBtnShadow.Opacity -gt 0) {
            $button.Effect = New-ShadowEffect $theme.NavBtnShadow
        }

        # TranslateTransformの準備
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

            # ホバーシャドウ
            $sender.Effect = New-ShadowEffect $theme.NavBtnHoverShadow
        })

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
            if ($theme.NavBtnShadow.Opacity -gt 0) {
                $sender.Effect = New-ShadowEffect $theme.NavBtnShadow
            } else {
                $sender.Effect = $null
            }
        })
    }

    # 前月・次月ボタンにテーマ＋ホバー適用
    & $script:ApplyNavButtonThemeAndHover $btnPrevMonth
    & $script:ApplyNavButtonThemeAndHover $btnNextMonth

    # 前月ボタン
    $btnPrevMonth.Add_Click({
        $script:dispMonth--
        if ($script:dispMonth -lt 1) {
            $script:dispMonth = 12
            $script:dispYear--
        }
        $script:selectedDay = 0
        & $script:DrawCalendar
    })

    # 次月ボタン
    $btnNextMonth.Add_Click({
        $script:dispMonth++
        if ($script:dispMonth -gt 12) {
            $script:dispMonth = 1
            $script:dispYear++
        }
        $script:selectedDay = 0
        & $script:DrawCalendar
    })

    # 初回描画
    & $script:DrawCalendar
}

# 日付をリストに追加する共通関数
function Add-DateToList {
    param($Day)

    if ($Day -le 0) { return }

    $dateStr = "{0}/{1}/{2}" -f $script:dispYear, $script:dispMonth, $Day
    if (-not $script:dateList.Contains($dateStr)) {
        [void]$script:dateList.Add($dateStr)
        $sorted = $script:dateList | Sort-Object { [datetime]$_ }
        $script:txtDateList.Text = ($sorted) -join "  "
    }
}
