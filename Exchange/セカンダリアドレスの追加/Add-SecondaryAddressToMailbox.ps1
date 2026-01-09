param (
    [string]$Mailbox,
    [string]$email
)

# Exchange Onlineに接続済みかどうかを判定する関数
function Test-ExchangeOnlineConnection {
    try {
        # Get-ExoMailboxを使って接続の確認を試みる
        Get-ExoMailbox -ResultSize 1 > $null
        return $true
    } catch {
        return $false
    }
}

# Exchange Onlineに接続されていない場合のみ接続を実行
if (-not (Test-ExchangeOnlineConnection)) {
    Import-Module ExchangeOnlineManagement
    # デバイスコード認証で接続（GUI環境がなくても動作する）
    Connect-ExchangeOnline -Device -ShowProgress $false
    Write-Host "Exchange Onlineに接続しました。"
} else {
    Write-Host "既にExchange Onlineに接続されています。"
}

try {
    # 既存のメールアドレスリストを取得
    $existingEmails = Get-Mailbox -Identity $Mailbox -ErrorAction Stop | Select-Object -ExpandProperty EmailAddresses

    # 既存のメールアドレスリストを表示
    Write-Output "メールボックス '$Mailbox' のメールアドレスリスト:"
    $existingEmails | ForEach-Object { Write-Output "  $_" }

    # 既存のメールアドレスリストを mailaddresses.txt に出力(毎回上書き)
    $existingEmails | Out-File -FilePath mailaddresses.txt -Encoding utf8 -Force

    if ($existingEmails -contains "smtp:$email") {
        Write-Output "エイリアス '$email' は既にメールボックス '$Mailbox' に存在します。"
    } else {
        # 新しいメールアドレスを追加
        $newEmails = $existingEmails + "smtp:$email"

        # メールアドレスリストを更新
        Set-Mailbox -Identity $Mailbox -EmailAddresses $newEmails -ErrorAction Stop

        Write-Output "セカンダリメールアドレス '$email' をメールボックス '$Mailbox' に追加しました。"
    }
}
catch {
    Write-Error "エラーが発生しました: $_"
}
