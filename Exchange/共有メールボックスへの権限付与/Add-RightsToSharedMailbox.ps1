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

# マネージドIDでの接続を試みる関数
function Connect-ExchangeOnlineWithManagedIdentity {
    try {
        # マネージドIDでの接続を試みる
        Import-Module -Name Microsoft.Identity.Client
        Connect-ExchangeOnline -ManagedIdentity -ShowProgress $false
        return $true
    } catch {
        # エラーを無視して、falseを返す
        return $false
    }
}

# Exchange Onlineに接続されていない場合のみ接続を実行
if (-not (Test-ExchangeOnlineConnection)) {
    # まずマネージドIDでの接続を試みる
    $connected = Connect-ExchangeOnlineWithManagedIdentity

    # マネージドIDでの接続が失敗した場合、対話的に接続を試みる
    if (-not $connected) {
        Import-Module ExchangeOnlineManagement
        Connect-ExchangeOnline -ShowProgress $false
        Write-Host "Exchange Onlineに接続しました。"
    }
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
