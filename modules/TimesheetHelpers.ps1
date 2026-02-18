# ==============================
# タイムシート ヘルパー関数・ダイアログ
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
