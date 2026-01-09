param (
    [string]$User
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

# 新たにユーザー情報を取得・表示するコードを追加
$recipientObject = Get-Recipient -Identity $User
Write-Host "姓: $($recipientObject.LastName)"
Write-Host "名: $($recipientObject.FirstName)"
Write-Host "表示名: $($recipientObject.DisplayName)"
Write-Host "メールアドレス: $($recipientObject.PrimarySmtpAddress)"

$groups = Get-DistributionGroup -ResultSize Unlimited | Where-Object {
    (Get-DistributionGroupMember $_.Identity | Where-Object {
        $_.PrimarySmtpAddress -eq $recipientObject.PrimarySmtpAddress
    }) -ne $null
}
Write-Host "所属グループ: $($groups.Name -join ', ')"

$sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited |
    Where-Object {
        (Get-MailboxPermission -Identity $_.Alias -User $recipientObject.PrimarySmtpAddress -ErrorAction SilentlyContinue) -ne $null
    }
Write-Host "権限を持っている共有メールボックス: $($sharedMailboxes.Name -join ', ')"

