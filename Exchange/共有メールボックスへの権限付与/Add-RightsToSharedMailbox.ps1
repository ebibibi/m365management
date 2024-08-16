param (
    [string]$SharedMailbox,
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
    Connect-ExchangeOnline -ShowProgress $false
    Write-Host "Exchange Onlineに接続しました。"
} else {
    Write-Host "既にExchange Onlineに接続されています。"
}

# フルメールボックスアクセス権限を付与します
Add-MailboxPermission -Identity $SharedMailbox -User $User -AccessRights FullAccess -InheritanceType All -Confirm:$false
Write-Host "ユーザー '$User' に '$SharedMailbox' へのフルメールボックスアクセス権限を付与しました。"

# メールボックス所有者として送信する権限を付与します
Add-RecipientPermission -Identity $SharedMailbox -Trustee $User -AccessRights SendAs -Confirm:$false
Write-Host "ユーザー '$User' に '$SharedMailbox' から送信する権限 (SendAs) を付与しました。"

# 終了
Write-Host "権限の付与が完了しました。"
