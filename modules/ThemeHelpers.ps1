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
