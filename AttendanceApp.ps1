# ==============================
# 勤怠打刻アプリめいん
# ==============================

# WPFアセンブリ読み込み
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:ProjectRoot = $PSScriptRoot

# モジュール読み込み
. "$scriptDir\modules\Config.ps1"
. "$scriptDir\modules\ThemeColors.ps1"
. "$scriptDir\modules\ThemeHelpers.ps1"
. "$scriptDir\modules\ThemeEngine.ps1"
. "$scriptDir\modules\CalendarLogic.ps1"
. "$scriptDir\modules\ControlLogic.ps1"
. "$scriptDir\modules\TimesheetConstants.ps1"
. "$scriptDir\modules\TimesheetHelpers.ps1"
. "$scriptDir\modules\TeamsWebhook.ps1"
. "$scriptDir\modules\TimesheetActions.ps1"
. "$scriptDir\modules\SettingsLogic.ps1"

# XAML読み込み（外部ファイルから）
$xamlPath = Join-Path $scriptDir "MainWindow.xaml"
$xamlContent = [System.IO.File]::ReadAllText($xamlPath)
$reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xamlContent))
$window = [System.Windows.Markup.XamlReader]::Load($reader)
$reader.Close()

# アイコン設定(用意してるなら有効化)
#$iconPath = Join-Path $scriptDir "logo.ico"
#if (Test-Path $iconPath) {
#    $window.Icon = $iconPath
#}

# カレンダー初期化
Initialize-Calendar -window $window

# コントロールパネル初期化
Initialize-ControlPanel -window $window

# 設定タブ初期化
Initialize-SettingsTab -window $window

# 出勤形態タブ初期化
Initialize-ShiftTypeTab -window $window

# 初期テーマ適用
$currentTheme = if ($script:settings.theme) { $script:settings.theme } else { "light" }
Apply-Theme -window $window -themeName $currentTheme

# ウィンドウ表示
[void]$window.ShowDialog()

# 後片付け
if ($script:timer -ne $null) {
    $script:timer.Stop()
}
