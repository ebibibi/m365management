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
    try {
        # 3.8.0バージョンを明示的に指定してモジュールをインポート
        $module = Get-Module -Name ExchangeOnlineManagement -ListAvailable | Where-Object {$_.Version -eq "3.8.0"} | Select-Object -First 1
        if ($module) {
            Import-Module -ModuleInfo $module -DisableNameChecking -Force
            Write-Host "Exchange Online Management バージョン 3.8.0 モジュールを読み込みました。"
        } else {
            # 3.8.0が見つからない場合は通常のインポート
            Import-Module ExchangeOnlineManagement -DisableNameChecking -Force
            Write-Host "Exchange Online Management モジュールを読み込みました。"
        }
        
        # 対話的に接続を試みる
        Connect-ExchangeOnline -ShowProgress $false -DisableWAM
        Write-Host "Exchange Onlineに接続しました。"
    } catch {
        Write-Error "Exchange Online Management モジュールの読み込みまたは接続に失敗しました: $($_.Exception.Message)"
        exit 1
    }
} else {
    Write-Host "既にExchange Onlineに接続されています。"
}

try {
    # フルメールボックスアクセス権限を付与します
    Add-MailboxPermission -Identity $SharedMailbox -User $User -AccessRights FullAccess -InheritanceType All -Confirm:$false
    Write-Host "ユーザー '$User' に '$SharedMailbox' へのフルメールボックスアクセス権限を付与しました。"

    # メールボックス所有者として送信する権限を付与します
    Add-RecipientPermission -Identity $SharedMailbox -Trustee $User -AccessRights SendAs -Confirm:$false
    Write-Host "ユーザー '$User' に '$SharedMailbox' から送信する権限 (SendAs) を付与しました。"
} catch {
    Write-Error "権限付与中にエラーが発生しました: $($_.Exception.Message)"
    exit 1
}
