# ==============================
# Teams Webhook 投稿
# ==============================
function Get-ProxyCredential {
    param([string]$Path)

    if (Test-Path $Path) {
        try {
            return Import-Clixml -Path $Path
        } catch {
            Write-Host "保存された認証情報の読み込みに失敗しました"
            Remove-Item $Path -ErrorAction SilentlyContinue
        }
    }

    # 新規入力
    $cred = Get-Credential -Message "プロキシ認証情報を入力してください"
    if ($null -eq $cred) {
        throw "認証情報の入力がキャンセルされました"
    }
    $cred | Export-Clixml -Path $Path
    return $cred
}

function Invoke-RestMethodWithAutoProxy {
    param(
        [string]$Uri,
        [string]$Method,
        [byte[]]$Body,
        [string]$ContentType
    )

    # プロキシ検出
    $systemProxy = [System.Net.WebRequest]::GetSystemWebProxy()
    $proxyUri = $systemProxy.GetProxy($Uri)

    # プロキシが必要かチェック（正しい比較方法）
    $needsProxy = ($proxyUri.AbsoluteUri.TrimEnd('/') -ne $Uri.TrimEnd('/'))

    # 基本パラメータ
    $params = @{
        Uri = $Uri
        Method = $Method
        Body = $Body
        ContentType = $ContentType
    }

    # プロキシが必要な場合のみ認証情報を取得して追加
    if ($needsProxy) {
        Write-Host "プロキシ経由で接続します: $($proxyUri.AbsoluteUri)"
        $params['Proxy'] = $proxyUri.AbsoluteUri  # 文字列として渡す
        $params['ProxyCredential'] = Get-ProxyCredential -Path $credPath
    } else {
        Write-Host "プロキシなしで直接接続します"
    }

    # 実行
    try {
        Invoke-RestMethod @params
    } catch {
        # 407エラーかつプロキシ使用時のみリトライ
        if ($needsProxy -and $_.Exception.Response.StatusCode -eq 407) {
            Write-Host "プロキシ認証に失敗しました。認証情報を再入力してください。"
            Remove-Item $credPath -ErrorAction SilentlyContinue

            $params['ProxyCredential'] = Get-ProxyCredential -Path $credPath

            # リトライ
            try {
                Invoke-RestMethod @params
            } catch {
                Write-Host "リトライも失敗しました: $($_.Exception.Message)"
                throw
            }
        } else {
            throw
        }
    }
}

function Send-TeamsPost {
    param(
        [string]$CheckType,
        [string]$WorkMode       = "",
        [string]$NextDateText   = "",
        [string]$NextShift      = "",
        [string]$NextWorkMode   = "",
        [array]$MentionData     = @(),
        [string]$Comment        = ""
    )

    $webhookUrl = $script:settings.teams_workflow.webhook_url
    if ([string]::IsNullOrWhiteSpace($webhookUrl)) {
        throw "WebhookURLが設定されていません。"
    }

    $userName = $script:settings.user_info.full_name
    $userId   = $script:settings.user_info.teams_principal_id

    # column_obj
    $columnObj = @{
        type  = "Column"
        width = "stretch"
        items = @(
            @{
                type    = "TextBlock"
                text    = "${userName}が${CheckType}しました"
                size    = "Medium"
                wrap    = $true
                weight  = "Bolder"
                verticalContentAlignment = "Center"
            }
        )
    }

    # message_obj
    if ($CheckType -eq "出勤") {
        $messageObj = @{
            type    = "TextBlock"
            text    = "業務を開始します(${WorkMode})"
            size    = "Medium"
            wrap    = $true
            spacing = "None"
        }
    } else {
        $messageObj = @{
            type  = "Container"
            spacing = "None"
            items = @(
                @{
                    type    = "TextBlock"
                    text    = "退勤します。次回は${NextDateText} ${NextWorkMode}(${NextShift})です。"
                    size    = "Medium"
                    wrap    = $true
                    spacing = "None"
                },
                @{
                    type    = "TextBlock"
                    text    = "お疲れさまでした。"
                     wrap    = $true
                     spacing = "None"
                }
            )
        }
    }

    # comment_obj
    if ([string]::IsNullOrWhiteSpace($Comment)) {
        $commentObj = @{}
    } elseif ($MentionData -and $MentionData.Count -gt 0) {
        $commentObj = @{
            type    = "TextBlock"
            text    = "コメント: ${Comment}"
            wrap    = $true
            spacing = "None"
        }
    } else {
        $commentObj = @{
            type      = "TextBlock"
            text      = "コメント: ${Comment}"
            wrap      = $true
            spacing   = "Small"
            separator = $true
        }
    }

    # 今のTeamsWorkFlowのアダプティブカード設定に合わせてmessageObjの形式を変える（条件付き）
    if (
        $CheckType -eq "退勤" -and
        $commentObj.Count -gt 0 -and
        (-not $MentionData -or $MentionData.Count -eq 0)
    ) {
        $messageObj.items += $commentObj
    }

    # payload 組み立て
    $mentionArr = @($MentionData | Where-Object { $_ })

    $payload = @{
        mention_data = $mentionArr
        userId       = $userId
        column       = (ConvertTo-Json -InputObject $columnObj -Depth 10 -Compress)
        message      = (ConvertTo-Json -InputObject $messageObj -Depth 10 -Compress)
        comment      = (ConvertTo-Json -InputObject $commentObj -Depth 10 -Compress)
    }

    $body = ConvertTo-Json -InputObject $payload -Depth 10
    # デバッグ用: Postデータをファイル出力
    #$debugPath = Join-Path $script:settings.timesheet_folder "teams_post_debug.json"
    #$body | Out-File -FilePath $debugPath -Encoding utf8 -Force

    #Invoke-RestMethod -Uri $webhookUrl -Method Post -Body ([System.Text.Encoding]::UTF8.GetBytes($body)) -ContentType "application/json; charset=utf-8"
    $credPath = "$env:USERPROFILE\.proxy_cred.xml"
    try {
        Invoke-RestMethodWithAutoProxy `
            -Uri $webhookUrl `
            -Method Post `
            -Body ([System.Text.Encoding]::UTF8.GetBytes($body)) `
            -ContentType "application/json; charset=utf-8"

        Write-Host "送信成功"
    } catch {
        Write-Host "送信失敗: $($_.Exception.Message)"
        throw
    }

}
