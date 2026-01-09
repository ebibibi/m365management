param (
    [Parameter(Mandatory=$true)]
    [string]$UserId,  # ユーザーのメールIDを受け取る
    [Parameter(Mandatory=$true)]
    [string]$SecurityGroup,  # セキュリティグループの名前を受け取る
    [Parameter(Mandatory=$true)]
    [string]$adminUserId # ExchangeOnlineに接続する管理者ユーザー名
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
    # 現在のユーザー情報を取得
    $currentUser = Get-User -Identity $adminUserId | Where-Object {$_.RecipientType -eq "User"}
    $group = Get-DistributionGroup -Identity $SecurityGroup

    # 現在のユーザーがManagedByに含まれているかを確認
    if ($group.ManagedBy -notcontains $currentUser.Name) {
        # 現在のユーザーをManagedBy属性に追加
        Set-DistributionGroup -Identity $SecurityGroup -ManagedBy @{Add="$($currentUser.Name)"} -BypassSecurityGroupManagerCheck
        Write-Output "現在のユーザーをグループ '$SecurityGroup' のManagedBy属性に追加しました。"
    }

    # ユーザーのオブジェクトを取得
    $user = Get-Recipient -Identity $UserId

    if (-not $user) {
        Write-Error "ユーザー '$UserId' が見つかりません。"
        exit 1
    }

    if (-not $group) {
        Write-Error "セキュリティグループ '$SecurityGroup' が見つかりません。"
        exit 1
    }

    # ユーザーをグループに追加
    Add-DistributionGroupMember -Identity $SecurityGroup -Member $UserId
    Write-Output "ユーザー '$UserId' をセキュリティグループ '$SecurityGroup' に追加しました。"

} catch {
    Write-Error "エラーが発生しました: $_"
    exit 1
}
