<#
.SYNOPSIS
    Generates an inventory report of Microsoft 365 objects (Users, Guests, Groups, Teams, Team Members / Owners,
    Distribution Groups, Shared Mailboxes, etc.) for clean-up review.

.DESCRIPTION
    • Connects silently to Microsoft Graph and Exchange Online (interactive only if session not present)  
    • Collects directory / EXO data with Microsoft.Graph SDK and Exchange Online V3 cmdlets  
    • Exports to a single Excel workbook (requires ImportExcel); if ImportExcel is absent, falls back to multiple CSV files  
    • Creates ./output folder with a timestamped report and execution log

.NOTES
    Run after executing Setup-InventoryReportPrereqs.ps1 at least once to ensure modules / consent.

#>

param(
    [switch]$VerboseOutput
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ($PSBoundParameters.ContainsKey('VerboseOutput')) { $VerbosePreference = 'Continue' }

#region ── Constants
$ReportDate = Get-Date -Format 'yyyyMMdd_HHmm'
$OutputDir  = Join-Path $PSScriptRoot 'output'
$LogPath    = Join-Path $OutputDir  "Run_$ReportDate.log"
$UsingExcel = $false
$ExcelPath  = Join-Path $OutputDir  "M365-Inventory-$ReportDate.xlsx"
#endregion

#region ── Helper
function Log {
    param([string]$Message,[string]$Level='INFO')
    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line  = "[$stamp][$Level] $Message"
    Write-Host $line
    Add-Content -Path $LogPath -Value $line
}
#endregion

#region ── Prepare output folder / log
if (-not (Test-Path $OutputDir)) { $null = New-Item -ItemType Directory -Path $OutputDir }
"Inventory run started $((Get-Date))" | Out-File $LogPath
#endregion

#region ── Module import
foreach ($m in 'Microsoft.Graph','ExchangeOnlineManagement') {
    Import-Module $m -ErrorAction Stop
}
$excelModule = Get-Module -Name ImportExcel -ListAvailable | Select-Object -First 1
if ($excelModule) {
    Import-Module ImportExcel -ErrorAction Stop
    $UsingExcel = $true
    Log "ImportExcel detected. Output will be single workbook."
}
#endregion

#region ── Connect to Microsoft Graph
try {
    if (-not (Get-MgContext)) {
        $Scopes = @(
            "Directory.Read.All",
            "User.Read.All",
            "Group.Read.All",
            "AuditLog.Read.All",
            "Team.ReadBasic.All",
            "Channel.ReadBasic.All"
        )
        Log "Connecting to Microsoft Graph ..."
        # デバイスコード認証で接続（GUI環境がなくても動作する）
        Connect-MgGraph -NoWelcome -Scopes $Scopes -UseDeviceCode
    }
    Log "Connected to Graph as $((Get-MgContext).Account)"
} catch {
    Log "Graph connection failed: $($_.Exception.Message)" 'ERR'
    throw
}
#endregion

#region ── Connect to Exchange Online
try {
    # EXO V3 uses REST API, not PSSession. Check using Get-ConnectionInformation cmdlet.
    $exoConnection = Get-ConnectionInformation -ErrorAction SilentlyContinue
    if (-not $exoConnection) {
        Log "Connecting to Exchange Online ..."
        # デバイスコード認証で接続（GUI環境がなくても動作する）
        Connect-ExchangeOnline -Device
    }
    Log "Connected to Exchange Online."
} catch {
    Log "EXO connection failed: $($_.Exception.Message)" 'ERR'
    throw
}
#endregion

#region ── Collect functions
function Get-Users {
    Log 'Collecting Users...'
    try {
        Get-MgUser -All -Filter "userType eq 'Member'" -Property 'id,displayName,mail,userPrincipalName,accountEnabled,proxyAddresses' |
            Select-Object id,displayName,mail,userPrincipalName,accountEnabled,
                          @{N='ProxyAddresses';E={($_.ProxyAddresses -join ';')}}
    } catch {
        Log "Failed to get users: $($_.Exception.Message)" 'WARN'
        @()
    }
}

function Get-Guests {
    Log 'Collecting Guests...'
    try {
        Get-MgUser -All -Filter "userType eq 'Guest'" -Property 'id,displayName,mail,userPrincipalName,proxyAddresses,externalUserState' |
            Select-Object id,displayName,mail,userPrincipalName,externalUserState,
                          @{N='ProxyAddresses';E={($_.ProxyAddresses -join ';')}}
    } catch {
        Log "Failed to get guests: $($_.Exception.Message)" 'WARN'
        @()
    }
}

function Get-Groups {
    Log 'Collecting Groups (non-Team)...'
    try {
        # Get all groups - resourceProvisioningOptions may not be available in all environments
        $allGroups = Get-MgGroup -All -Property 'id,displayName,mail,mailEnabled,mailNickname,proxyAddresses,groupTypes,createdDateTime'
        # Since we can't reliably filter Teams, return all groups
        $allGroups | Select-Object id,displayName,mail,mailNickname,mailEnabled,groupTypes,
                      @{N='ProxyAddresses';E={($_.ProxyAddresses -join ';')}},
                      createdDateTime
    } catch {
        Log "Failed to get groups: $($_.Exception.Message)" 'WARN'
        @()
    }
}

function Get-Teams {
    Log 'Collecting Teams...'
    Get-MgGroup -All -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" -Property 'id,displayName,mail,visibility,createdDateTime,proxyAddresses' |
        Select-Object id,displayName,mail,visibility,createdDateTime,
                      @{N='ProxyAddresses';E={($_.ProxyAddresses -join ';')}}
}

function Get-TeamMembers {
    param([string[]]$TeamIds)
    $results = @()
    $teamLookup = @{}
    # Create lookup table for team names
    foreach ($team in $Teams) {
        $teamLookup[$team.Id] = $team.DisplayName
    }
    
    foreach ($tid in $TeamIds) {
        $teamName = $teamLookup[$tid]
        Log "   Members for Team $tid"
        try {
            $members = Get-MgGroupMember -GroupId $tid -All
            foreach ($member in $members) {
                # Get full user details
                try {
                    if ($member.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.user') {
                        $userDetails = Get-MgUser -UserId $member.Id -Property 'displayName,userPrincipalName,mail,userType' -ErrorAction SilentlyContinue
                        $results += [PSCustomObject]@{
                            ParentTeamId = $tid
                            ParentTeamName = $teamName
                            MemberId = $member.Id
                            MemberDisplayName = $userDetails.DisplayName
                            MemberUPN = $userDetails.UserPrincipalName
                            MemberEmail = $userDetails.Mail
                            MemberType = $userDetails.UserType
                        }
                    } else {
                        # For non-user members (groups, etc.)
                        $results += [PSCustomObject]@{
                            ParentTeamId = $tid
                            ParentTeamName = $teamName
                            MemberId = $member.Id
                            MemberDisplayName = $member.AdditionalProperties['displayName']
                            MemberUPN = ''
                            MemberEmail = ''
                            MemberType = 'Group/Other'
                        }
                    }
                } catch {
                    Log "Failed to get details for member ${member.Id}: $($_.Exception.Message)" 'WARN'
                }
            }
        } catch {
            Log "Failed to get members for ${tid}: $($_.Exception.Message)" 'WARN'
        }
    }
    $results
}

function Get-TeamOwners {
    param([string[]]$TeamIds)
    $results = @()
    $teamLookup = @{}
    # Create lookup table for team names
    foreach ($team in $Teams) {
        $teamLookup[$team.Id] = $team.DisplayName
    }
    
    foreach ($tid in $TeamIds) {
        $teamName = $teamLookup[$tid]
        Log "   Owners for Team $tid"
        try {
            $owners = Get-MgGroupOwner -GroupId $tid
            foreach ($owner in $owners) {
                # Get full user details
                try {
                    $userDetails = Get-MgUser -UserId $owner.Id -Property 'displayName,userPrincipalName,mail,userType' -ErrorAction SilentlyContinue
                    $results += [PSCustomObject]@{
                        ParentTeamId = $tid
                        ParentTeamName = $teamName
                        OwnerId = $owner.Id
                        OwnerDisplayName = $userDetails.DisplayName
                        OwnerUPN = $userDetails.UserPrincipalName
                        OwnerEmail = $userDetails.Mail
                        OwnerType = $userDetails.UserType
                    }
                } catch {
                    Log "Failed to get details for owner ${owner.Id}: $($_.Exception.Message)" 'WARN'
                }
            }
        } catch {
            Log "Failed to get owners for ${tid}: $($_.Exception.Message)" 'WARN'
        }
    }
    $results
}

function Get-SharedMailboxes {
    Log 'Collecting Shared Mailboxes...'
    Get-ExoMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited |
        Select-Object Guid,DisplayName,PrimarySmtpAddress,UserPrincipalName,
                      @{N='ProxyAddresses';E={($_.EmailAddresses -join ';')}},
                      @{N='LastLogonTime';E={$_.LastLogonTime}},
                      @{N='TotalItemSize';E={$_.TotalItemSize}}
}

function Get-DistributionGroups {
    Log 'Collecting Distribution Groups...'
    Get-ExoRecipient -RecipientTypeDetails MailUniversalDistributionGroup,MailUniversalSecurityGroup -ResultSize Unlimited |
        Select-Object Guid,DisplayName,PrimarySmtpAddress,Alias,
                      @{N='ProxyAddresses';E={($_.EmailAddresses -join ';')}},
                      GroupType
}

function Get-DLMembers {
    param($DistributionGroups)
    $results=@()
    foreach ($dlObj in $DistributionGroups) {
        $dl = $dlObj.PrimarySmtpAddress
        if ([string]::IsNullOrWhiteSpace($dl)) { 
            $dl = $dlObj.Alias
        }
        if ([string]::IsNullOrWhiteSpace($dl)) { continue }
        
        Log "   Members for DL $dl"
        try {
            $members = Get-DistributionGroupMember -Identity $dl -ResultSize Unlimited
            foreach ($member in $members) {
                $results += [PSCustomObject]@{
                    ParentDL = $dl
                    ParentDLName = $dlObj.DisplayName
                    MemberPrimarySMTP = $member.PrimarySmtpAddress
                    MemberDisplayName = $member.DisplayName
                    MemberAlias = $member.Alias
                    MemberType = $member.RecipientType
                    MemberRecipientTypeDetails = $member.RecipientTypeDetails
                }
            }
        } catch {
            Log "Failed to get DL members for ${dl}: $($_.Exception.Message)" 'WARN'
        }
    }
    $results
}

function Get-SharedMailboxPermissions {
    param($SharedMailboxes)
    $results = @()
    Log 'Collecting Shared Mailbox Permissions...'
    
    foreach ($mb in $SharedMailboxes) {
        $mbIdentity = $mb.PrimarySmtpAddress
        if ([string]::IsNullOrWhiteSpace($mbIdentity)) { 
            $mbIdentity = $mb.UserPrincipalName 
        }
        if ([string]::IsNullOrWhiteSpace($mbIdentity)) { continue }
        
        Log "   Permissions for Shared Mailbox $mbIdentity"
        
        # Get FullAccess permissions
        try {
            $fullAccessPerms = Get-ExoMailboxPermission -Identity $mbIdentity | 
                Where-Object { $_.User -notlike 'NT AUTHORITY\*' -and $_.User -notlike 'S-1-5-*' -and $_.IsInherited -eq $false }
            
            foreach ($perm in $fullAccessPerms) {
                $results += [PSCustomObject]@{
                    SharedMailbox = $mbIdentity
                    SharedMailboxName = $mb.DisplayName
                    User = $perm.User
                    AccessRights = ($perm.AccessRights -join ',')
                    PermissionType = 'FullAccess'
                    IsInherited = $perm.IsInherited
                    Deny = $perm.Deny
                }
            }
        } catch {
            Log "Failed to get FullAccess permissions for ${mbIdentity}: $($_.Exception.Message)" 'WARN'
        }
        
        # Get SendAs permissions
        try {
            $sendAsPerms = Get-ExoRecipientPermission -Identity $mbIdentity | 
                Where-Object { $_.Trustee -notlike 'NT AUTHORITY\*' -and $_.Trustee -notlike 'S-1-5-*' }
            
            foreach ($perm in $sendAsPerms) {
                $results += [PSCustomObject]@{
                    SharedMailbox = $mbIdentity
                    SharedMailboxName = $mb.DisplayName
                    User = $perm.Trustee
                    AccessRights = ($perm.AccessRights -join ',')
                    PermissionType = 'SendAs'
                    IsInherited = $perm.IsInherited
                    Deny = $false  # Get-ExoRecipientPermission doesn't have Deny property
                }
            }
        } catch {
            Log "Failed to get SendAs permissions for ${mbIdentity}: $($_.Exception.Message)" 'WARN'
        }
        
        # Get SendOnBehalf permissions
        try {
            $mailbox = Get-ExoMailbox -Identity $mbIdentity -Properties GrantSendOnBehalfTo
            if ($mailbox.GrantSendOnBehalfTo) {
                foreach ($trustee in $mailbox.GrantSendOnBehalfTo) {
                    $results += [PSCustomObject]@{
                        SharedMailbox = $mbIdentity
                        SharedMailboxName = $mb.DisplayName
                        User = $trustee
                        AccessRights = 'SendOnBehalf'
                        PermissionType = 'SendOnBehalf'
                        IsInherited = $false
                        Deny = $false
                    }
                }
            }
        } catch {
            Log "Failed to get SendOnBehalf permissions for ${mbIdentity}: $($_.Exception.Message)" 'WARN'
        }
    }
    
    $results
}
#endregion

#region ── Execute collection
$Users            = Get-Users
$Guests           = Get-Guests
$Groups           = Get-Groups
$Teams            = Get-Teams
$TeamMembers      = if ($Teams) { Get-TeamMembers -TeamIds $Teams.id } else { @() }
$TeamOwners       = if ($Teams) { Get-TeamOwners  -TeamIds $Teams.id } else { @() }
$SharedMailboxes  = Get-SharedMailboxes
$SharedMailboxPermissions = if ($SharedMailboxes) { Get-SharedMailboxPermissions -SharedMailboxes $SharedMailboxes } else { @() }
$DLs              = Get-DistributionGroups
$DLMembers        = if ($DLs) { Get-DLMembers -DistributionGroups $DLs } else { @() }
#endregion

#region ── Export
if ($UsingExcel) {
    Log 'Exporting to Excel...'
    $params = @{ Path = $ExcelPath; AutoSize = $true }
    if ($Users -and @($Users).Count -gt 0)           { $Users           | Export-Excel @params -WorksheetName Users          -FreezeTopRow -TableName 'Users' }
    if ($Guests -and @($Guests).Count -gt 0)          { $Guests          | Export-Excel @params -WorksheetName Guests         -FreezeTopRow -TableName 'Guests' }
    if ($Groups -and @($Groups).Count -gt 0)          { $Groups          | Export-Excel @params -WorksheetName Groups         -FreezeTopRow -TableName 'Groups' }
    if ($Teams -and @($Teams).Count -gt 0)           { $Teams           | Export-Excel @params -WorksheetName Teams          -FreezeTopRow -TableName 'Teams' }
    if ($TeamMembers -and @($TeamMembers).Count -gt 0)     { $TeamMembers     | Export-Excel @params -WorksheetName TeamMembers    -FreezeTopRow -TableName 'TeamMembers' }
    if ($TeamOwners -and @($TeamOwners).Count -gt 0)      { $TeamOwners      | Export-Excel @params -WorksheetName TeamOwners     -FreezeTopRow -TableName 'TeamOwners' }
    if ($SharedMailboxes -and @($SharedMailboxes).Count -gt 0) { $SharedMailboxes | Export-Excel @params -WorksheetName SharedMailboxes -FreezeTopRow -TableName 'SharedMailboxes' }
    if ($SharedMailboxPermissions -and @($SharedMailboxPermissions).Count -gt 0) { $SharedMailboxPermissions | Export-Excel @params -WorksheetName SharedMailboxPerms -FreezeTopRow -TableName 'SharedMailboxPerms' }
    if ($DLs -and @($DLs).Count -gt 0)             { $DLs             | Export-Excel @params -WorksheetName DistributionGroups -FreezeTopRow -TableName 'DistributionGroups' }
    if ($DLMembers -and @($DLMembers).Count -gt 0)       { $DLMembers       | Export-Excel @params -WorksheetName DLMembers      -FreezeTopRow -TableName 'DLMembers' }
    Log "Excel report saved to $ExcelPath"
} else {
    Log 'ImportExcel absent. Exporting to CSV files...'
    $Users           | Export-Csv (Join-Path $OutputDir 'Users.csv')            -NoTypeInfo -Encoding UTF8
    $Guests          | Export-Csv (Join-Path $OutputDir 'Guests.csv')           -NoTypeInfo -Encoding UTF8
    $Groups          | Export-Csv (Join-Path $OutputDir 'Groups.csv')           -NoTypeInfo -Encoding UTF8
    $Teams           | Export-Csv (Join-Path $OutputDir 'Teams.csv')            -NoTypeInfo -Encoding UTF8
    $TeamMembers     | Export-Csv (Join-Path $OutputDir 'TeamMembers.csv')      -NoTypeInfo -Encoding UTF8
    $TeamOwners      | Export-Csv (Join-Path $OutputDir 'TeamOwners.csv')       -NoTypeInfo -Encoding UTF8
    $SharedMailboxes | Export-Csv (Join-Path $OutputDir 'SharedMailboxes.csv')  -NoTypeInfo -Encoding UTF8
    $SharedMailboxPermissions | Export-Csv (Join-Path $OutputDir 'SharedMailboxPermissions.csv') -NoTypeInfo -Encoding UTF8
    $DLs             | Export-Csv (Join-Path $OutputDir 'DistributionGroups.csv') -NoTypeInfo -Encoding UTF8
    $DLMembers       | Export-Csv (Join-Path $OutputDir 'DLMembers.csv')        -NoTypeInfo -Encoding UTF8
    Log "CSV files created in $OutputDir"
}
#endregion

#region ── Cleanup
Disconnect-MgGraph
Disconnect-ExchangeOnline -Confirm:$false
Log 'Inventory collection finished.'
#endregion
