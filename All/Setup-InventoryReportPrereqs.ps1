<#
.SYNOPSIS
    Installs/updates required modules and performs the initial consent flow for Microsoft Graph SDK and Exchange Online PowerShell.

.DESCRIPTION
    • Ensures PowerShell 7.2+  
    • Installs Microsoft.Graph, ImportExcel, ExchangeOnlineManagement (CurrentUser scope)  
    • Prompts the administrator to sign-in and grant consent for Graph scopes:
        Directory.Read.All, User.Read.All, Group.Read.All,
        Team.ReadBasic.All, Channel.ReadBasic.All
    • Connects to Exchange Online (interactive) to cache credentials.

.NOTES
    Run once per workstation (or when modules need updates).
#>

#region ── Helper
function Install-IfMissing {
    param(
        [Parameter(Mandatory)][string]$ModuleName
    )
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "Installing $ModuleName ..." -ForegroundColor Cyan
        Install-Module $ModuleName -Scope CurrentUser -Force -ErrorAction Stop
    }
    else {
        Write-Host "$ModuleName already present." -ForegroundColor Green
    }
}
#endregion

#region ── PowerShell version check
if ($PSVersionTable.PSVersion.Major -lt 7 -or
    ($PSVersionTable.PSVersion.Major -eq 7 -and $PSVersionTable.PSVersion.Minor -lt 2)) {
    Write-Error "PowerShell 7.2 以上が必要です。"
    return
}
#endregion

#region ── Install / update required modules
$required = @('Microsoft.Graph', 'ImportExcel', 'ExchangeOnlineManagement')
foreach ($m in $required) { Install-IfMissing -ModuleName $m }

# Optionally update to latest
foreach ($m in $required) {
    Write-Host "Updating $m to latest version (if needed)..." -ForegroundColor Yellow
    Update-Module $m -ErrorAction SilentlyContinue
}
#endregion

#region ── Microsoft Graph initial consent
Write-Host "`nConnecting to Microsoft Graph to acquire admin consent..." -ForegroundColor Cyan
try {
    Import-Module Microsoft.Graph -ErrorAction Stop
    $Scopes = @(
        "Directory.Read.All",
        "User.Read.All",
        "Group.Read.All",
        "AuditLog.Read.All",
        "Team.ReadBasic.All",
        "Channel.ReadBasic.All"
    )
    Connect-MgGraph -Scopes $Scopes -NoWelcome
    Write-Host "Graph consent completed." -ForegroundColor Green
    Disconnect-MgGraph
} catch {
    Write-Warning "Graph connection failed: $($_.Exception.Message)"
}
#endregion

#region ── Exchange Online initial connection
Write-Host "`nConnecting to Exchange Online..." -ForegroundColor Cyan
try {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    # close old sessions
    Get-PSSession | Where-Object { $_.ConfigurationName -eq 'Microsoft.Exchange' } | Remove-PSSession -ErrorAction SilentlyContinue
    Connect-ExchangeOnline -DisableWAM
    Write-Host "Exchange Online connected. Closing session..." -ForegroundColor Green
    Disconnect-ExchangeOnline -Confirm:$false
} catch {
    Write-Warning "Exchange Online connection failed: $($_.Exception.Message)"
}
#endregion

Write-Host "`nPrerequisite setup finished." -ForegroundColor Green
