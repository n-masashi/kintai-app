# ==============================
# テーマエンジン
# ==============================

# タイトルバーのダークモード切替用 DWM API
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class DwmHelper {
    [DllImport("dwmapi.dll", PreserveSig = true)]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

    public static void SetImmersiveDarkMode(IntPtr hwnd, bool enabled) {
        int value = enabled ? 1 : 0;
        // DWMWA_USE_IMMERSIVE_DARK_MODE = 20 (Windows 11 / Windows 10 1903+)
        DwmSetWindowAttribute(hwnd, 20, ref value, sizeof(int));
    }
}
"@

# ==============================
# 関数
# ==============================

# グラデーションブラシを作成
function New-GradientBrush {
    param(
        [string[]]$Colors,
        [double]$Angle = 135
    )

    if ($Colors.Count -eq 1) {
        return New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($Colors[0]))
    }

    $brush = New-Object System.Windows.Media.LinearGradientBrush
    $brush.StartPoint = New-Object System.Windows.Point(0, 0)
    $brush.EndPoint = New-Object System.Windows.Point(1, 1)

    $brush.GradientStops.Add((New-Object System.Windows.Media.GradientStop([System.Windows.Media.ColorConverter]::ConvertFromString($Colors[0]), 0)))
    $brush.GradientStops.Add((New-Object System.Windows.Media.GradientStop([System.Windows.Media.ColorConverter]::ConvertFromString($Colors[1]), 1)))

    return $brush
}

# DropShadowEffect作成ヘルパー
function New-ShadowEffect {
    param($ShadowDef)
    $effect = New-Object System.Windows.Media.Effects.DropShadowEffect
    $effect.Color = [System.Windows.Media.ColorConverter]::ConvertFromString($ShadowDef.Color)
    $effect.BlurRadius = $ShadowDef.Blur
    $effect.ShadowDepth = $ShadowDef.Depth
    $effect.Opacity = $ShadowDef.Opacity
    return $effect
}

# ComboBoxにテーマ対応テンプレートを適用
function Apply-ComboBoxTemplate {
    param($comboBox, $theme)

    $bgColor = $theme.InputBg
    $fgColor = $theme.InputText
    $borderColor = $theme.InputBorder
    $hoverColor = $theme.CardBorder

    $templateXaml = @"
<ControlTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                 xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
                 TargetType="ComboBox">
    <Grid>
        <ToggleButton Name="ToggleButton"
                      IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                      Focusable="False"
                      ClickMode="Press">
            <ToggleButton.Template>
                <ControlTemplate TargetType="ToggleButton">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition/>
                            <ColumnDefinition Width="24"/>
                        </Grid.ColumnDefinitions>
                        <Border x:Name="Border" Grid.ColumnSpan="2" Background="$bgColor" BorderBrush="$borderColor" BorderThickness="1" CornerRadius="4"/>
                        <Border x:Name="SplitBorder" Grid.Column="1" Background="Transparent" Margin="0,2,2,2" CornerRadius="0,2,2,0"/>
                        <Path x:Name="Arrow" Grid.Column="1" Fill="$fgColor" HorizontalAlignment="Center" VerticalAlignment="Center"
                              Data="M 0 0 L 4 4 L 8 0 Z"/>
                    </Grid>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsMouseOver" Value="True">
                            <Setter TargetName="Border" Property="BorderBrush" Value="$hoverColor"/>
                        </Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </ToggleButton.Template>
        </ToggleButton>
        <ContentPresenter Name="ContentSite" IsHitTestVisible="False"
                          Content="{TemplateBinding SelectionBoxItem}"
                          ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                          Margin="6,3,24,3"
                          VerticalAlignment="Center"
                          HorizontalAlignment="Left">
            <ContentPresenter.Resources>
                <Style TargetType="TextBlock">
                    <Setter Property="Foreground" Value="$fgColor"/>
                </Style>
            </ContentPresenter.Resources>
        </ContentPresenter>
        <Popup Name="Popup" IsOpen="{TemplateBinding IsDropDownOpen}" Placement="Bottom"
               Focusable="False" AllowsTransparency="True" PopupAnimation="Slide">
            <Grid Name="DropDown" MinWidth="{TemplateBinding ActualWidth}"
                  MaxHeight="{TemplateBinding MaxDropDownHeight}" SnapsToDevicePixels="True">
                <Border x:Name="DropDownBorder" Background="$bgColor" BorderBrush="$borderColor"
                        BorderThickness="1" CornerRadius="4" Margin="0,1,0,0"/>
                <ScrollViewer Margin="2,4,2,4" SnapsToDevicePixels="True">
                    <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained"/>
                </ScrollViewer>
            </Grid>
        </Popup>
    </Grid>
</ControlTemplate>
"@

    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($templateXaml))
    $template = [System.Windows.Markup.XamlReader]::Load($reader)
    $reader.Close()
    $comboBox.Template = $template
    $comboBox.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($fgColor))
}

# ==============================
# テーマ適用関数
# ==============================

function Apply-Theme {
    param($window, [string]$themeName)

    $theme = $global:ThemeColors[$themeName]
    if (-not $theme) { $themeName = "light"; $theme = $global:ThemeColors["light"] }

    # テーマ名を即座に反映（DrawCalendar等が参照するため）
    $script:settings.theme = $themeName

    # ===== タイトルバー ダークモード切替 =====
    $hwndSource = [System.Windows.Interop.WindowInteropHelper]::new($window)
    if ($hwndSource.Handle -ne [IntPtr]::Zero) {
        [DwmHelper]::SetImmersiveDarkMode($hwndSource.Handle, ($themeName -eq "dark"))
    }

    # ===== Window & Main Backgrounds =====
    $window.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.ContentBackground))

    # 打刻タブのグリッド背景
    $mainTabGrid = $window.FindName("MainTabGrid")
    if ($mainTabGrid) {
        $mainTabGrid.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.CardBackground))
    }

    # ===== Calendar Border =====
    $calendarBorder = $window.FindName("CalendarBorder")
    if ($calendarBorder) {
        $calendarBorder.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.CardBackground))
        $calendarBorder.BorderBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.CardBorder))
        $calendarBorder.CornerRadius = $theme.CornerRadiusMedium
        $calendarBorder.Effect = New-ShadowEffect $theme.CardShadow
    }

    # ===== Control Panel Border =====
    $controlPanelBorder = $window.FindName("ControlPanelBorder")
    if ($controlPanelBorder) {
        $controlPanelBorder.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.InputBg))
        $controlPanelBorder.BorderBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.CardBorder))
        $controlPanelBorder.CornerRadius = $theme.CornerRadiusMedium
        $controlPanelBorder.Effect = New-ShadowEffect $theme.CardShadow
    }

    # ===== Time Display Border =====
    $timeDisplayBorder = $window.FindName("TimeDisplayBorder")
    if ($timeDisplayBorder) {
        $timeDisplayBorder.Background = New-GradientBrush -Colors $theme.TimeDisplayBg
        $timeDisplayBorder.CornerRadius = $theme.TimeDisplayCornerRadius
        $tdPad = $theme.TimeDisplayPadding
        $timeDisplayBorder.Padding = New-Object System.Windows.Thickness($tdPad[0], $tdPad[1], $tdPad[0], $tdPad[1])
        $timeDisplayBorder.Effect = New-ShadowEffect $theme.TimeDisplayShadow
        if ($theme.TimeDisplayBorderColor) {
            $timeDisplayBorder.BorderBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.TimeDisplayBorderColor))
            $timeDisplayBorder.BorderThickness = $theme.TimeDisplayBorderThickness
        } else {
            $timeDisplayBorder.BorderThickness = 0
        }
    }

    # ===== Text Elements =====
    $txtDate = $window.FindName("TxtDate")
    if ($txtDate) {
        $txtDate.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.TimeDateText))
        $txtDate.FontSize = $theme.TimeDateSize
        $txtDate.FontWeight = $theme.TimeDateWeight
        $dm = $theme.TimeDateMargin
        $txtDate.Margin = New-Object System.Windows.Thickness($dm[0], $dm[1], $dm[2], $dm[3])
    }

    $txtTime = $window.FindName("TxtTime")
    if ($txtTime) {
        if ($theme.TimeClockText.Count -gt 1) {
            $txtTime.Foreground = New-GradientBrush -Colors $theme.TimeClockText
        } else {
            $txtTime.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.TimeClockText[0]))
        }
        $txtTime.FontSize = $theme.TimeClockSize
        $txtTime.FontWeight = "Bold"
    }

    $lblShiftType = $window.FindName("LblShiftType")
    if ($lblShiftType) {
        $lblShiftType.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.TextPrimary))
    }

    $lblWorkStyle = $window.FindName("LblWorkStyle")
    if ($lblWorkStyle) {
        $lblWorkStyle.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.TextPrimary))
    }

    $lblOptions = $window.FindName("LblOptions")
    if ($lblOptions) {
        $lblOptions.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.TextPrimary))
    }

    # RadioButtons & CheckBoxes
    $radioRemote = $window.FindName("RadioRemote")
    if ($radioRemote) {
        $radioRemote.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.TextPrimary))
    }

    $radioOffice = $window.FindName("RadioOffice")
    if ($radioOffice) {
        $radioOffice.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.TextPrimary))
    }

    $chkNoTeamsPost = $window.FindName("ChkNoTeamsPost")
    if ($chkNoTeamsPost) {
        $chkNoTeamsPost.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.TextPrimary))
    }

    $chkEstimatedInput = $window.FindName("ChkEstimatedInput")
    if ($chkEstimatedInput) {
        $chkEstimatedInput.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.TextPrimary))
    }

    # ===== Work Style & Options Borders =====
    $workStyleBorder = $window.FindName("WorkStyleBorder")
    if ($workStyleBorder) {
        $workStyleBorder.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.InputBg))
        $workStyleBorder.CornerRadius = $theme.CornerRadiusInput
    }

    $optionsBorder = $window.FindName("OptionsBorder")
    if ($optionsBorder) {
        $optionsBorder.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.InputBg))
        $optionsBorder.CornerRadius = $theme.CornerRadiusInput
    }

    # ===== Timesheet Button =====
    $btnTimesheet = $window.FindName("BtnTimesheet")
    if ($btnTimesheet) {
        $btnBg = $theme.TimesheetBtnBg
        $btnFg = $theme.TimesheetBtnText
        $btnBorderColor = $theme.TimesheetBtnBorderColor
        $btnBorderThick = $theme.TimesheetBtnBorderThickness
        $btnCorner = $theme.TimesheetBtnCornerRadius
        $btnPad = $theme.TimesheetBtnPadding
        $btnMargin = $theme.TimesheetBtnMargin
        $hoverBgColors = $theme.TimesheetBtnHoverBg
        $hoverFg = $theme.TimesheetBtnHoverText

        if ($hoverBgColors.Count -gt 1) {
            $hoverBgSetter = @"
            <Setter Property="Background">
                <Setter.Value>
                    <LinearGradientBrush StartPoint="0,0" EndPoint="1,0">
                        <GradientStop Color="$($hoverBgColors[0])" Offset="0"/>
                        <GradientStop Color="$($hoverBgColors[1])" Offset="1"/>
                    </LinearGradientBrush>
                </Setter.Value>
            </Setter>
"@
        } else {
            $hoverBgSetter = "<Setter Property=""Background"" Value=""$($hoverBgColors[0])""/>"
        }

        $hoverBorderSetter = ""
        if ($themeName -eq "dark") {
            $hoverBorderSetter = "<Setter Property=""BorderBrush"" Value=""#60A5FA""/>"
        }

        $templateXaml = @"
<ControlTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                 xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
                 TargetType="Button">
    <Border x:Name="BtnBorder" Background="$btnBg"
            BorderBrush="$btnBorderColor" BorderThickness="$btnBorderThick"
            CornerRadius="$btnCorner" Padding="$($btnPad[0]),$($btnPad[1])">
        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
    </Border>
    <ControlTemplate.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
            $hoverBgSetter
            <Setter Property="Foreground" Value="$hoverFg"/>
            $hoverBorderSetter
        </Trigger>
    </ControlTemplate.Triggers>
</ControlTemplate>
"@
        $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($templateXaml))
        $template = [System.Windows.Markup.XamlReader]::Load($reader)
        $reader.Close()
        $btnTimesheet.Template = $template
        $btnTimesheet.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($btnFg))
        $btnTimesheet.FontSize = 11
        $btnTimesheet.FontWeight = "Bold"
        $btnTimesheet.Margin = New-Object System.Windows.Thickness($btnMargin[0], $btnMargin[1], $btnMargin[2], $btnMargin[3])
    }

    # ===== Clock In/Out Buttons =====
    $clockInBorder = $window.FindName("ClockInBorder")
    if ($clockInBorder) {
        $clockInBorder.Background = New-GradientBrush -Colors $theme.BtnSuccess
        $clockInBorder.CornerRadius = $theme.CornerRadiusButton
        $clockInBorder.Effect = New-ShadowEffect $theme.BtnSuccessShadow
    }

    $clockOutBorder = $window.FindName("ClockOutBorder")
    if ($clockOutBorder) {
        $clockOutBorder.Background = New-GradientBrush -Colors $theme.BtnDanger
        $clockOutBorder.CornerRadius = $theme.CornerRadiusButton
        $clockOutBorder.Effect = New-ShadowEffect $theme.BtnDangerShadow
    }

    # ===== Small Action Buttons =====
    $addDateBorder = $window.FindName("AddDateBorder")
    if ($addDateBorder) {
        $addDateBorder.Background = New-GradientBrush -Colors $theme.BtnInfo
        $addDateBorder.CornerRadius = $theme.CornerRadiusInput
        $addDateBorder.Effect = New-ShadowEffect $theme.BtnInfoShadow
    }

    $clearDatesBorder = $window.FindName("ClearDatesBorder")
    if ($clearDatesBorder) {
        $clearDatesBorder.Background = New-GradientBrush -Colors $theme.BtnSecondary
        $clearDatesBorder.CornerRadius = $theme.CornerRadiusInput
        $clearDatesBorder.Effect = New-ShadowEffect $theme.BtnSecondaryShadow
    }

    $bulkInputBorder = $window.FindName("BulkInputBorder")
    if ($bulkInputBorder) {
        $bulkInputBorder.Background = New-GradientBrush -Colors $theme.BtnWarning
        $bulkInputBorder.CornerRadius = $theme.CornerRadiusInput
        $bulkInputBorder.Effect = New-ShadowEffect $theme.BtnWarningShadow
    }

    # ===== Date List TextBox =====
    $txtDateList = $window.FindName("TxtDateList")
    if ($txtDateList) {
        $txtDateList.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.InputBg))
        $txtDateList.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.TextSecondary))
    }

    # ===== Settings Tab =====
    $settingsScrollViewer = $window.FindName("SettingsScrollViewer")
    if ($settingsScrollViewer) {
        $settingsScrollViewer.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.CardBackground))
    }

    # Settings Section Titles
    foreach ($name in @("LblAppearance", "LblUserInfo", "LblTeamsWorkflow", "LblManagers")) {
        $lbl = $window.FindName($name)
        if ($lbl) {
            $lbl.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.TabActiveText))
        }
    }

    # Settings Section Borders
    foreach ($name in @("ThemeSectionBorder", "UserInfoBorder", "TeamsWorkflowBorder", "ManagersBorder")) {
        $border = $window.FindName($name)
        if ($border) {
            $border.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.InputBg))
            $border.BorderBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.CardBorder))
            $border.CornerRadius = $theme.CornerRadiusInput
        }
    }

    # Settings Buttons
    $addManagerBorder = $window.FindName("AddManagerBorder")
    if ($addManagerBorder) {
        $addManagerBorder.Background = New-GradientBrush -Colors $theme.BtnInfo
        $addManagerBorder.CornerRadius = $theme.CornerRadiusInput
    }

    $deleteManagerBorder = $window.FindName("DeleteManagerBorder")
    if ($deleteManagerBorder) {
        $deleteManagerBorder.Background = New-GradientBrush -Colors $theme.BtnDanger
        $deleteManagerBorder.CornerRadius = $theme.CornerRadiusInput
    }

    $saveSettingsBorder = $window.FindName("SaveSettingsBorder")
    if ($saveSettingsBorder) {
        $saveSettingsBorder.Background = New-GradientBrush -Colors $theme.BtnSuccess
        $saveSettingsBorder.CornerRadius = $theme.CornerRadiusButton
    }

    # ===== TabItem テーマ動的適用 =====
    $mainTabControl = $window.FindName("MainTabControl")
    if ($mainTabControl) {
        $activeBgColor = $theme.TabActive
        $activeTextColor = $theme.TabActiveText
        $inactiveBgColor = "Transparent"
        $inactiveTextColor = $theme.TabInactive
        $hoverBgColor = $theme.TabHoverBg
        $hoverTextColor = $theme.TabHoverText
        $cornerRadius = $theme.CornerRadiusTab

        foreach ($tabItem in $mainTabControl.Items) {
            $templateXaml = @"
<ControlTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                 xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
                 TargetType="TabItem">
    <Border Name="Border" Background="Transparent" CornerRadius="$cornerRadius,$cornerRadius,0,0" Padding="20,10" Margin="2,0">
        <ContentPresenter x:Name="ContentSite" ContentSource="Header" VerticalAlignment="Center" HorizontalAlignment="Center"/>
    </Border>
    <ControlTemplate.Triggers>
        <Trigger Property="IsSelected" Value="True">
            <Setter TargetName="Border" Property="Background" Value="$($activeBgColor[0])"/>
            <Setter Property="Foreground" Value="$activeTextColor"/>
        </Trigger>
        <Trigger Property="IsSelected" Value="False">
            <Setter TargetName="Border" Property="Background" Value="$inactiveBgColor"/>
            <Setter Property="Foreground" Value="$inactiveTextColor"/>
        </Trigger>
        <Trigger Property="IsMouseOver" Value="True">
            <Setter TargetName="Border" Property="Background" Value="$hoverBgColor"/>
            <Setter Property="Foreground" Value="$hoverTextColor"/>
        </Trigger>
    </ControlTemplate.Triggers>
</ControlTemplate>
"@
            $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($templateXaml))
            $template = [System.Windows.Markup.XamlReader]::Load($reader)
            $reader.Close()
            $tabItem.Template = $template
        }
    }

    # ===== 設定タブ内テキスト要素のテーマ適用 =====
    $lblTheme = $window.FindName("LblTheme")
    if ($lblTheme) {
        $lblTheme.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.TextPrimary))
    }

    # 設定タブ内の全要素をツリー走査してテーマ適用
    $settingsScrollViewer = $window.FindName("SettingsScrollViewer")
    if ($settingsScrollViewer) {
        $stack = New-Object System.Collections.Stack
        $stack.Push($settingsScrollViewer)
        while ($stack.Count -gt 0) {
            $element = $stack.Pop()
            if ($element -is [System.Windows.Controls.TextBlock] -and -not $element.Name) {
                $element.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.TextPrimary))
            }
            if ($element -is [System.Windows.Controls.TextBox]) {
                $element.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.InputBg))
                $element.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.InputText))
                $element.BorderBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.InputBorder))
            }
            if ($element -is [System.Windows.Controls.ComboBox]) {
                Apply-ComboBoxTemplate -comboBox $element -theme $theme
            }
            if ($element -is [System.Windows.Controls.ListView]) {
                $element.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.InputBg))
                $element.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.InputText))
                $element.BorderBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.InputBorder))

                $lvBgColor = $theme.InputBg
                $lvFgColor = $theme.InputText
                $lvHoverBg = $theme.CardBorder
                $lvSelectedBg = if ($themeName -eq "dark") { "#1E3A8A" } else { "#DBEAFE" }
                $styleXaml = @"
<Style xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
       xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
       TargetType="ListViewItem">
    <Setter Property="Background" Value="$lvBgColor"/>
    <Setter Property="Foreground" Value="$lvFgColor"/>
    <Setter Property="BorderThickness" Value="0"/>
    <Setter Property="Padding" Value="4,2"/>
    <Setter Property="Template">
        <Setter.Value>
            <ControlTemplate TargetType="ListViewItem">
                <Border x:Name="Bd" Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}" BorderThickness="0">
                    <GridViewRowPresenter/>
                </Border>
                <ControlTemplate.Triggers>
                    <Trigger Property="IsMouseOver" Value="True">
                        <Setter TargetName="Bd" Property="Background" Value="$lvHoverBg"/>
                    </Trigger>
                    <Trigger Property="IsSelected" Value="True">
                        <Setter TargetName="Bd" Property="Background" Value="$lvSelectedBg"/>
                    </Trigger>
                </ControlTemplate.Triggers>
            </ControlTemplate>
        </Setter.Value>
    </Setter>
</Style>
"@
                $reader2 = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($styleXaml))
                $style = [System.Windows.Markup.XamlReader]::Load($reader2)
                $reader2.Close()
                $element.ItemContainerStyle = $style
            }
            # 子要素を探索
            if ($element -is [System.Windows.Controls.Panel]) {
                foreach ($child in $element.Children) { $stack.Push($child) }
            }
            elseif ($element -is [System.Windows.Controls.ContentControl]) {
                if ($element.Content -is [System.Windows.UIElement]) { $stack.Push($element.Content) }
            }
            elseif ($element -is [System.Windows.Controls.Decorator]) {
                if ($element.Child) { $stack.Push($element.Child) }
            }
            elseif ($element -is [System.Windows.Controls.ItemsControl]) {
                foreach ($item in $element.Items) {
                    if ($item -is [System.Windows.UIElement]) { $stack.Push($item) }
                }
            }
        }
    }

    # ===== 打刻タブ内のComboBox テーマ適用 =====
    $cmbShiftType = $window.FindName("CmbShiftType")
    if ($cmbShiftType) {
        Apply-ComboBoxTemplate -comboBox $cmbShiftType -theme $theme
    }

    # ===== 月タイトルのテーマ適用 =====
    $txtMonthTitle = $window.FindName("TxtMonthTitle")
    if ($txtMonthTitle) {
        $txtMonthTitle.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.TextPrimary))
    }

    # ===== ナビゲーションボタンのテーマ適用 =====
    $btnPrevMonth = $window.FindName("BtnPrevMonth")
    if ($btnPrevMonth) {
        $btnPrevMonth.Background = New-GradientBrush -Colors $theme.CalNavBg
        $btnPrevMonth.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.CalNavText))
        if ($theme.NavBtnShadow.Opacity -gt 0) {
            $btnPrevMonth.Effect = New-ShadowEffect $theme.NavBtnShadow
        } else {
            $btnPrevMonth.Effect = $null
        }
    }
    $btnNextMonth = $window.FindName("BtnNextMonth")
    if ($btnNextMonth) {
        $btnNextMonth.Background = New-GradientBrush -Colors $theme.CalNavBg
        $btnNextMonth.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($theme.CalNavText))
        if ($theme.NavBtnShadow.Opacity -gt 0) {
            $btnNextMonth.Effect = New-ShadowEffect $theme.NavBtnShadow
        } else {
            $btnNextMonth.Effect = $null
        }
    }

    # カレンダー再描画（テーマ適用済みで）
    if ($script:DrawCalendar) {
        & $script:DrawCalendar
    }
}
