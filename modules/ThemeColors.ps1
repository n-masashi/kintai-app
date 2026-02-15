# ==============================
# テーマカラー定義
# ==============================

$global:ThemeColors = @{
    light = @{
        # Window & Background
        WindowBackground = "#FFFFFF"
        BodyBackground = @("#FDFBFB", "#EBEDEE")
        ContentBackground = "#FFFFFF"
        CardBackground = "#FFFFFF"
        CardBorder = "#F3F4F6"

        # Tab
        TabActive = @("#FCE7F3", "#E0E7FF")
        TabActiveText = "#7C3AED"
        TabInactive = "#9CA3AF"
        TabHoverBg = "#F9FAFB"
        TabHoverText = "#6B7280"

        # Time Display
        TimeDisplayBg = @("#DBEAFE", "#FDE7F3")
        TimeDisplayCornerRadius = 18
        TimeDisplayPadding = @(28, 20)
        TimeDisplayBorderColor = ""
        TimeDisplayBorderThickness = 0
        TimeDisplayShadow = @{ Color = "#3B82F6"; Blur = 16; Depth = 4; Opacity = 0.12 }
        TimeDateText = "#1E40AF"
        TimeDateSize = 18
        TimeDateWeight = "SemiBold"
        TimeDateMargin = @(0, 0, 16, 0)
        TimeClockText = @("#3B82F6", "#8B5CF6")
        TimeClockSize = 18
        TimesheetBtnBg = "#FFFFFF"
        TimesheetBtnText = "#3B82F6"
        TimesheetBtnBorderColor = "#3B82F6"
        TimesheetBtnBorderThickness = 2
        TimesheetBtnCornerRadius = 12
        TimesheetBtnPadding = @(24, 7)
        TimesheetBtnMargin = @(0, 12, 0, 0)
        TimesheetBtnHoverBg = @("#3B82F6", "#60A5FA")
        TimesheetBtnHoverText = "#FFFFFF"

        # Text
        TextPrimary = "#374151"
        TextSecondary = "#6B7280"
        TextDisabled = "#9CA3AF"

        # Buttons
        BtnSuccess = @("#10B981", "#34D399")
        BtnDanger = @("#EF4444", "#F87171")
        BtnInfo = @("#3B82F6", "#60A5FA")
        BtnWarning = @("#F59E0B", "#FBBF24")
        BtnSecondary = @("#6B7280", "#9CA3AF")

        # Calendar
        CalNavBg = @("#FCE7F3", "#E0E7FF")
        CalNavText = "#7C3AED"
        CalDayBg = "#FAFAFA"
        CalDayText = "#374151"
        CalDayHoverBg = @("#FCE7F3", "#E0E7FF")
        CalDayHoverText = "#7C3AED"
        CalTodayBg = @("#DBEAFE", "#DDD6FE")
        CalTodayText = "#7C3AED"
        CalSelectedBg = @("#7C3AED", "#A78BFA")
        CalSelectedText = "#FFFFFF"
        CalHolidayBg = "#FEF2F2"
        CalHolidayText = "#F87171"
        CalSaturdayBg = "#EFF6FF"
        CalSaturdayText = "#3B82F6"
        CalWeekdayBg = "#F9FAFB"
        CalWeekdayText = "#6B7280"

        # Form Elements
        InputBg = "#FAFAFA"
        InputBorder = "#F3F4F6"
        InputFocusBorder = "#C4B5FD"
        InputText = "#374151"

        # Corner Radius
        CornerRadiusLarge = 24
        CornerRadiusMedium = 20
        CornerRadiusSmall = 14
        CornerRadiusTiny = 12
        CornerRadiusButton = 14
        CornerRadiusInput = 12
        CornerRadiusTab = 12

        # Shadows
        CardShadow = @{ Color = "#000000"; Blur = 20; Depth = 0; Opacity = 0.06 }
        BtnSuccessShadow = @{ Color = "#10B981"; Blur = 16; Depth = 4; Opacity = 0.25 }
        BtnSuccessHoverShadow = @{ Color = "#10B981"; Blur = 24; Depth = 8; Opacity = 0.35 }
        BtnDangerShadow = @{ Color = "#EF4444"; Blur = 16; Depth = 4; Opacity = 0.25 }
        BtnDangerHoverShadow = @{ Color = "#EF4444"; Blur = 24; Depth = 8; Opacity = 0.35 }
        BtnInfoShadow = @{ Color = "#3B82F6"; Blur = 8; Depth = 2; Opacity = 0.25 }
        BtnInfoHoverShadow = @{ Color = "#3B82F6"; Blur = 12; Depth = 4; Opacity = 0.35 }
        BtnSecondaryShadow = @{ Color = "#6B7280"; Blur = 8; Depth = 2; Opacity = 0.25 }
        BtnSecondaryHoverShadow = @{ Color = "#6B7280"; Blur = 12; Depth = 4; Opacity = 0.35 }
        BtnWarningShadow = @{ Color = "#F59E0B"; Blur = 8; Depth = 2; Opacity = 0.25 }
        BtnWarningHoverShadow = @{ Color = "#F59E0B"; Blur = 12; Depth = 4; Opacity = 0.35 }
        NavBtnShadow = @{ Color = "#7C3AED"; Blur = 8; Depth = 2; Opacity = 0.15 }
        NavBtnHoverShadow = @{ Color = "#7C3AED"; Blur = 12; Depth = 4; Opacity = 0.25 }
    }
    dark = @{
        # Window & Background
        WindowBackground = "#1E293B"
        BodyBackground = @("#0F172A", "#0F172A")
        ContentBackground = "#1E293B"
        CardBackground = "#0F172A"
        CardBorder = "#334155"

        # Tab
        TabActive = @("#1E293B", "#1E293B")
        TabActiveText = "#60A5FA"
        TabInactive = "#64748B"
        TabHoverBg = "#334155"
        TabHoverText = "#94A3B8"

        # Time Display
        TimeDisplayBg = @("#1E293B", "#334155")
        TimeDisplayCornerRadius = 16
        TimeDisplayPadding = @(28, 18)
        TimeDisplayBorderColor = "#475569"
        TimeDisplayBorderThickness = 1
        TimeDisplayShadow = @{ Color = "#000000"; Blur = 12; Depth = 3; Opacity = 0.4 }
        TimeDateText = "#94A3B8"
        TimeDateSize = 18
        TimeDateWeight = "SemiBold"
        TimeDateMargin = @(0, 0, 14, 0)
        TimeClockText = @("#CBD5E1")
        TimeClockSize = 18
        TimesheetBtnBg = "#00000000"
        TimesheetBtnText = "#60A5FA"
        TimesheetBtnBorderColor = "#334155"
        TimesheetBtnBorderThickness = 1
        TimesheetBtnCornerRadius = 10
        TimesheetBtnPadding = @(22, 7)
        TimesheetBtnMargin = @(0, 10, 0, 0)
        TimesheetBtnHoverBg = @("#334155")
        TimesheetBtnHoverText = "#60A5FA"

        # Text
        TextPrimary = "#F1F5F9"
        TextSecondary = "#CBD5E1"
        TextDisabled = "#94A3B8"

        # Buttons
        BtnSuccess = @("#10B981", "#059669")
        BtnDanger = @("#EF4444", "#DC2626")
        BtnInfo = @("#3B82F6", "#2563EB")
        BtnWarning = @("#F59E0B", "#D97706")
        BtnSecondary = @("#6B7280", "#4B5563")

        # Calendar
        CalNavBg = @("#334155", "#334155")
        CalNavText = "#60A5FA"
        CalDayBg = "#1E293B"
        CalDayText = "#CBD5E1"
        CalDayHoverBg = @("#334155", "#334155")
        CalDayHoverText = "#F1F5F9"
        CalTodayBg = @("#1E3A8A", "#1E3A8A")
        CalTodayText = "#60A5FA"
        CalSelectedBg = @("#60A5FA", "#3B82F6")
        CalSelectedText = "#0F172A"
        CalHolidayBg = "#1E293B"
        CalHolidayText = "#F87171"
        CalSaturdayBg = "#1E293B"
        CalSaturdayText = "#60A5FA"
        CalWeekdayBg = "#1E293B"
        CalWeekdayText = "#64748B"

        # Form Elements
        InputBg = "#0F172A"
        InputBorder = "#334155"
        InputFocusBorder = "#60A5FA"
        InputText = "#F1F5F9"

        # Corner Radius
        CornerRadiusLarge = 16
        CornerRadiusMedium = 12
        CornerRadiusSmall = 10
        CornerRadiusTiny = 8
        CornerRadiusButton = 10
        CornerRadiusInput = 8
        CornerRadiusTab = 8

        # Shadows
        CardShadow = @{ Color = "#000000"; Blur = 20; Depth = 0; Opacity = 0.15 }
        BtnSuccessShadow = @{ Color = "#10B981"; Blur = 12; Depth = 4; Opacity = 0.30 }
        BtnSuccessHoverShadow = @{ Color = "#10B981"; Blur = 16; Depth = 6; Opacity = 0.40 }
        BtnDangerShadow = @{ Color = "#EF4444"; Blur = 12; Depth = 4; Opacity = 0.30 }
        BtnDangerHoverShadow = @{ Color = "#EF4444"; Blur = 16; Depth = 6; Opacity = 0.40 }
        BtnInfoShadow = @{ Color = "#60A5FA"; Blur = 8; Depth = 2; Opacity = 0.20 }
        BtnInfoHoverShadow = @{ Color = "#60A5FA"; Blur = 12; Depth = 4; Opacity = 0.30 }
        BtnSecondaryShadow = @{ Color = "#64748B"; Blur = 8; Depth = 2; Opacity = 0.20 }
        BtnSecondaryHoverShadow = @{ Color = "#64748B"; Blur = 12; Depth = 4; Opacity = 0.30 }
        BtnWarningShadow = @{ Color = "#FB923C"; Blur = 8; Depth = 2; Opacity = 0.20 }
        BtnWarningHoverShadow = @{ Color = "#FB923C"; Blur = 12; Depth = 4; Opacity = 0.30 }
        NavBtnShadow = @{ Color = "#334155"; Blur = 0; Depth = 0; Opacity = 0 }
        NavBtnHoverShadow = @{ Color = "#60A5FA"; Blur = 8; Depth = 2; Opacity = 0.20 }
    }
}
