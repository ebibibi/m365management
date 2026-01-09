# ExchangeOnlineManagementに接続するためのスクリプト
Write-Host "Exchange Online への接続を開始します..." -ForegroundColor Cyan

# パッケージ管理モジュールの最新化を確認
$packages = @("PackageManagement", "PowerShellGet")
foreach ($package in $packages) {
    $currentModule = Get-Module -Name $package -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
    Write-Host "$($package) の現在のバージョン: $($currentModule.Version)" -ForegroundColor Yellow
}

# 既存のセッションを全て切断
Get-PSSession | Where-Object { $_.ConfigurationName -eq 'Microsoft.Exchange' } | Remove-PSSession
Write-Host "既存のExchange Onlineセッションを切断しました" -ForegroundColor Green

# モジュールをアンロード
Remove-Module -Name ExchangeOnlineManagement -Force -ErrorAction SilentlyContinue
Write-Host "ExchangeOnlineManagement モジュールをアンロードしました" -ForegroundColor Green

# 最新バージョンを使用
$latestModule = Get-Module -Name ExchangeOnlineManagement -ListAvailable | 
                Sort-Object Version -Descending | 
                Select-Object -First 1

Write-Host "使用するバージョン: $($latestModule.Version)" -ForegroundColor Green

# 明示的にそのバージョンを読み込む
Import-Module -ModuleInfo $latestModule -Force -DisableNameChecking -Verbose
Write-Host "バージョン $($latestModule.Version) を読み込みました" -ForegroundColor Green

# デバイスコード認証で接続（GUI環境がなくても動作する）
try {
    Connect-ExchangeOnline -Device
    Write-Host "Exchange Onlineに接続しました" -ForegroundColor Green
    
    # 接続情報を表示
    Get-ConnectionInformation
} catch {
    Write-Error "Exchange Onlineへの接続に失敗しました: $($_.Exception.Message)"
}
