# ==============================
# 設定タブ
# ==============================

# 管理職リストのデータソース(このリストは退勤時のメンション処理時につかう)
$script:managersCollection = New-Object System.Collections.ObjectModel.ObservableCollection[PSCustomObject]

function Initialize-SettingsTab {
    param($window)

    # コントロール取得（イベントハンドラから参照するため $script: に格納）
    $script:txtAdUsername = $window.FindName("txtAdUsername")
    $script:txtFullName = $window.FindName("txtFullName")
    $script:txtTeamsPrincipalId = $window.FindName("txtTeamsPrincipalId")
    $script:txtShiftDisplayName = $window.FindName("txtShiftDisplayName")
    $script:txtTimesheetDisplayName = $window.FindName("txtTimesheetDisplayName")
    $script:txtWebhookUrl = $window.FindName("txtWebhookUrl")
    $script:txtTimesheetFolder = $window.FindName("txtTimesheetFolder")
    $script:txtAttendanceDataFolder = $window.FindName("txtAttendanceDataFolder")
    $script:lvManagers = $window.FindName("lvManagers")
    $script:txtNewManagerName = $window.FindName("txtNewManagerName")
    $script:txtNewManagerId = $window.FindName("txtNewManagerId")
    $btnAddManager = $window.FindName("btnAddManager")
    $btnDeleteManager = $window.FindName("btnDeleteManager")
    $btnSaveSettings = $window.FindName("btnSaveSettings")
    $script:cmbTheme = $window.FindName("cmbTheme")
    $script:settingsWindow = $window

    # 設定値をUIに反映
    $script:txtAdUsername.Text = $script:settings.user_info.ad_username
    $script:txtFullName.Text = $script:settings.user_info.full_name
    $script:txtTeamsPrincipalId.Text = $script:settings.user_info.teams_principal_id
    $script:txtShiftDisplayName.Text = $script:settings.user_info.shift_display_name
    $script:txtTimesheetDisplayName.Text = $script:settings.user_info.timesheet_display_name
    $script:txtWebhookUrl.Password = $script:settings.teams_workflow.webhook_url
    $script:txtTimesheetFolder.Text = $script:settings.timesheet_folder
    $script:txtAttendanceDataFolder.Text = $script:settings.attendance_data_folder

    # テーマコンボボックス設定
    if ($script:cmbTheme) {
        [void]$script:cmbTheme.Items.Add("light")
        [void]$script:cmbTheme.Items.Add("dark")
        $currentTheme = if ($script:settings.theme) { $script:settings.theme } else { "light" }
        $script:cmbTheme.SelectedItem = $currentTheme

        # テーマ変更時のイベント
        $script:cmbTheme.Add_SelectionChanged({
            $newTheme = $script:cmbTheme.SelectedItem
            if ($newTheme) {
                Apply-Theme -window $script:settingsWindow -themeName $newTheme
            }
        })
    }

    # 管理職リストを読み込み
    $script:managersCollection.Clear()
    foreach ($manager in $script:settings.managers) {
        $script:managersCollection.Add([PSCustomObject]@{
            Name = $manager.name
            TeamsPrincipalId = $manager.teams_principal_id
        })
    }
    $script:lvManagers.ItemsSource = $script:managersCollection

    # プレースホルダー定数
    $script:placeholderName = "例：例名"
    $script:placeholderId = "例：tarou@example.com"

    # 初期プレースホルダー表示
    $currentTheme = if ($script:settings.theme) { $script:settings.theme } else { "light" }
    $theme = $global:ThemeColors[$currentTheme]
    $script:txtNewManagerName.Text = $script:placeholderName
    $script:txtNewManagerName.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.TextDisabled))
    $script:txtNewManagerId.Text = $script:placeholderId
    $script:txtNewManagerId.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.TextDisabled))

    # プレースホルダー処理（名前欄）
    $script:txtNewManagerName.Add_GotFocus({
        if ($script:txtNewManagerName.Text -eq $script:placeholderName) {
            $script:txtNewManagerName.Text = ""
            $currentTheme = if ($script:settings.theme) { $script:settings.theme } else { "light" }
            $theme = $global:ThemeColors[$currentTheme]
            $script:txtNewManagerName.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.InputText))
        }
    })
    $script:txtNewManagerName.Add_LostFocus({
        if ([string]::IsNullOrWhiteSpace($script:txtNewManagerName.Text)) {
            $script:txtNewManagerName.Text = $script:placeholderName
            $currentTheme = if ($script:settings.theme) { $script:settings.theme } else { "light" }
            $theme = $global:ThemeColors[$currentTheme]
            $script:txtNewManagerName.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.TextDisabled))
        }
    })

    # プレースホルダー処理（ID欄）
    $script:txtNewManagerId.Add_GotFocus({
        if ($script:txtNewManagerId.Text -eq $script:placeholderId) {
            $script:txtNewManagerId.Text = ""
            $currentTheme = if ($script:settings.theme) { $script:settings.theme } else { "light" }
            $theme = $global:ThemeColors[$currentTheme]
            $script:txtNewManagerId.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.InputText))
        }
    })
    $script:txtNewManagerId.Add_LostFocus({
        if ([string]::IsNullOrWhiteSpace($script:txtNewManagerId.Text)) {
            $script:txtNewManagerId.Text = $script:placeholderId
            $currentTheme = if ($script:settings.theme) { $script:settings.theme } else { "light" }
            $theme = $global:ThemeColors[$currentTheme]
            $script:txtNewManagerId.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.TextDisabled))
        }
    })

    # 管理職追加ボタンのイベント
    $btnAddManager.Add_Click({
        $name = $script:txtNewManagerName.Text.Trim()
        $id = $script:txtNewManagerId.Text.Trim()

        # プレースホルダーの場合は空とみなす
        if ($name -eq $script:placeholderName) { $name = "" }
        if ($id -eq $script:placeholderId) { $id = "" }

        if ([string]::IsNullOrEmpty($name) -or [string]::IsNullOrEmpty($id)) {
            [System.Windows.MessageBox]::Show("名前とTeamsPrincipalIDを入力してください。", "入力エラー", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }

        $script:managersCollection.Add([PSCustomObject]@{
            Name = $name
            TeamsPrincipalId = $id
        })

        # プレースホルダーに戻す
        $currentTheme = if ($script:settings.theme) { $script:settings.theme } else { "light" }
        $theme = $global:ThemeColors[$currentTheme]
        $script:txtNewManagerName.Text = $script:placeholderName
        $script:txtNewManagerName.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.TextDisabled))
        $script:txtNewManagerId.Text = $script:placeholderId
        $script:txtNewManagerId.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.TextDisabled))
    })

    # 管理職削除ボタンのイベント
    $btnDeleteManager.Add_Click({
        $selected = $script:lvManagers.SelectedItem
        if ($null -eq $selected) {
            [System.Windows.MessageBox]::Show("削除する項目を選択してください。", "選択エラー", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }

        $script:managersCollection.Remove($selected)
    })

    # 保存ボタンのイベント
    $btnSaveSettings.Add_Click({
        # UIから設定を取得
        $script:settings.user_info.ad_username = $script:txtAdUsername.Text
        $script:settings.user_info.full_name = $script:txtFullName.Text
        $script:settings.user_info.teams_principal_id = $script:txtTeamsPrincipalId.Text
        $script:settings.user_info.shift_display_name = $script:txtShiftDisplayName.Text
        $script:settings.user_info.timesheet_display_name = $script:txtTimesheetDisplayName.Text
        $script:settings.teams_workflow.webhook_url = $script:txtWebhookUrl.Password
        $script:settings.timesheet_folder = $script:txtTimesheetFolder.Text
        $script:settings.attendance_data_folder = $script:txtAttendanceDataFolder.Text

        # テーマ設定を更新
        if ($script:cmbTheme -and $script:cmbTheme.SelectedItem) {
            $script:settings.theme = $script:cmbTheme.SelectedItem
        }

        # 管理職リストを更新
        $script:settings.managers = @($script:managersCollection | ForEach-Object {
            [PSCustomObject]@{
                name = $_.Name
                teams_principal_id = $_.TeamsPrincipalId
            }
        })

        # 設定を保存
        Save-Settings

        [System.Windows.MessageBox]::Show("設定を保存しました。", "保存完了", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
    })
}
