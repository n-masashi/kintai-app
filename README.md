# 勤怠打刻アプリ

社内の勤怠管理を効率化するためのPowerShell製デスクトップアプリケーションです。

## 概要

Excelベースのタイムシートへの打刻作業を自動化し、Teams通知との連携により勤怠管理をスムーズにします。

## 主な機能

- **出勤・退勤打刻**: ボタン一つでExcelタイムシートに自動記入
- **シフト管理**: 日勤/早番/遅番/深夜など複数のシフトパターンに対応
- **休暇管理**: 有給休暇、振替休暇、特別休暇などの記録
- **Teams連携**: 出退勤時にMicrosoft Teamsへ自動通知(Workflowはデータに合わせて別途作成の必要があり）
- **一括入力**: カレンダーから複数日を選択して一括記入
- **テーマ切替**: ライト/ダークモード対応

## 技術スタック

- **言語**: PowerShell
- **UI**: WPF (XAML)
- **Excel操作**: COM Object
- **通知**: Microsoft Teams Webhook

## 動作環境

- Windows 10/11
- PowerShell 5.1以上
- Microsoft Excel

## 使用方法

1. `settings.json` を編集してユーザー情報とフォルダパスを設定
2. `AttendanceApp.ps1` を実行
3. カレンダーから日付を選択し、シフトタイプを選んで打刻

## ファイル構成

```
kintai-app/
├── AttendanceApp.ps1           # メインスクリプト
├── MainWindow.xaml             # UI定義
├── settings.json               # 設定ファイル
├── modules/                    # 機能モジュール
│   ├── CalendarLogic.ps1
│   ├── Config.ps1
│   ├── ControlLogic.ps1
│   ├── SettingsLogic.ps1
│   ├── ThemeColors.ps1
│   ├── ThemeEngine.ps1
│   └── TimesheetLogic.ps1
├── attendance_folder/          # 勤怠データ出力先（出勤自動チェックにつかってる）
└── timesheet/                  # タイムシートExcel格納先
```

## 注意事項

このアプリケーションは特定の社内ルールに合わせて作成されたため、そのままでは他の環境で動作しない可能性があります。コードはカスタマイズの参考としてご利用ください。

## ライセンス

MIT License
