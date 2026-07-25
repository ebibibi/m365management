#requires -Version 5.1
<#
.SYNOPSIS
    Active Directory 健全性診断レポート（HTML）を生成します。

.DESCRIPTION
    Active Directory 環境を読み取り専用で検査し、単一ファイルの HTML レポートを出力します。
    検出結果は「ユーザー影響」を軸に 3 つのフェーズへ自動的に振り分けられ、
    「まず何から手を付けるか」が読み手に分かる形で提示されます。

    このスクリプトは情報の取得のみを行います。Active Directory の構成を変更する処理は
    一切含まれていません（使用するのは Get-* 系コマンドレットのみ）。

.PARAMETER InactiveDays
    非アクティブと判定する日数。既定は 90 日。

.PARAMETER StalePasswordDays
    重要アカウントのパスワードが「長期間未変更」と判定される日数。既定は 180 日。

.PARAMETER OutputPath
    レポートの出力先ディレクトリ。既定はカレントディレクトリ。

.PARAMETER CompanyName
    レポートに表示する組織名。既定は取得したドメイン名。

.PARAMETER Anonymize
    ユーザー名・コンピューター名・識別名を伏せ字にして出力します。
    社外に結果を渡す前提の場合に使用してください。

.PARAMETER DemoData
    Active Directory へ接続せず、サンプル（架空組織）のデータでレポートを生成します。
    レポートの体裁確認や提案時のサンプル提示に使用します。

.EXAMPLE
    .\Invoke-ADHealthReport.ps1 -OutputPath C:\Temp
    既定設定で診断し、C:\Temp に HTML レポートを出力します。

.EXAMPLE
    .\Invoke-ADHealthReport.ps1 -Anonymize -OutputPath C:\Temp
    アカウント名などを伏せた状態でレポートを出力します。

.EXAMPLE
    .\Invoke-ADHealthReport.ps1 -DemoData -OutputPath .
    AD に接続せずサンプルレポートを生成します。

.NOTES
    読み取り専用 / Read-only.
#>
[CmdletBinding()]
param(
    [int]$InactiveDays = 90,
    [int]$StalePasswordDays = 180,
    [string]$OutputPath = '.',
    [string]$CompanyName,
    [switch]$Anonymize,
    [switch]$DemoData
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region ---------- 定数 ----------

# 重要度
$script:SEV_HIGH = 'High'
$script:SEV_MED  = 'Medium'
$script:SEV_LOW  = 'Low'

# 対処フェーズ（NKG案件で有効だった「ユーザー影響で切る」分類）
$script:PHASE1 = 'Phase1'  # ユーザー影響なし・即実施可能
$script:PHASE2 = 'Phase2'  # ユーザー影響の事前調査が必要
$script:PHASE3 = 'Phase3'  # 整理・整頓（急がないが放置しない）

$script:PhaseMeta = [ordered]@{
    Phase1 = @{ Label = 'フェーズ1'; Title = 'ユーザー影響なし・すぐ着手できる'; Desc = '管理者側の作業だけで完了し、エンドユーザーへの影響がありません。ここから着手するのが最も安全で、効果もすぐ出ます。' }
    Phase2 = @{ Label = 'フェーズ2'; Title = '影響範囲の事前調査が必要'; Desc = '対処そのものは難しくありませんが、止まるものがないかの確認が先に要ります。調査してから計画的に進めてください。' }
    Phase3 = @{ Label = 'フェーズ3'; Title = '整理・整頓（急がないが放置しない）'; Desc = '直ちに危険ではないものの、放置すると調査や移行のたびにコストを払い続けることになります。手が空いたときに。' }
}

#endregion

#region ---------- ユーティリティ ----------

function Write-Step {
    param([string]$Message)
    Write-Host "[*] $Message" -ForegroundColor Cyan
}

function Write-Warn {
    param([string]$Message)
    Write-Host "[!] $Message" -ForegroundColor Yellow
}

<#
    検出結果を1件作る。すべてのチェック関数はこの形のオブジェクトを返す。
#>
function New-Finding {
    param(
        [Parameter(Mandatory)][string]$Category,
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][ValidateSet('High','Medium','Low')][string]$Severity,
        [Parameter(Mandatory)][ValidateSet('Phase1','Phase2','Phase3')][string]$Phase,
        [Parameter(Mandatory)][string]$Detail,
        [string]$Impact = 'なし（管理者側の作業のみで完了します）',
        [string]$Action = '',
        [string]$Reference = '',
        [object[]]$Items = @()
    )
    [pscustomobject]@{
        Category  = $Category
        Title     = $Title
        Severity  = $Severity
        Phase     = $Phase
        Detail    = $Detail
        Impact    = $Impact
        Action    = $Action
        Reference = $Reference
        Items     = @($Items)
        Count     = @($Items).Count
    }
}

function Protect-Name {
    param([string]$Value)
    if (-not $Anonymize -or [string]::IsNullOrWhiteSpace($Value)) { return $Value }
    # 同じ値が同じ伏せ字になるよう、ハッシュの先頭8桁を使う（追跡可能・復元不可）
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $bytes = [Text.Encoding]::UTF8.GetBytes($Value.ToLowerInvariant())
    $hash = ($sha.ComputeHash($bytes) | ForEach-Object { $_.ToString('x2') }) -join ''
    $sha.Dispose()
    return ('***-{0}' -f $hash.Substring(0,8))
}

function ConvertTo-HtmlText {
    param([string]$Value)
    if ($null -eq $Value) { return '' }
    return [System.Net.WebUtility]::HtmlEncode($Value)
}

function Get-DaysAgo {
    param([Nullable[datetime]]$When, [datetime]$Now)
    if (-not $When) { return $null }
    return [int]([math]::Floor(($Now - $When).TotalDays))
}

#endregion

#region ---------- データ収集（読み取り専用） ----------

<#
    AD から診断に必要な情報を取得する。
    ここで取得したスナップショットを各チェック関数へ渡す。
#>
function Get-AdSnapshot {
    param([datetime]$Now)

    Import-Module ActiveDirectory -ErrorAction Stop

    $domain = Get-ADDomain
    $forest = Get-ADForest
    $rootDse = Get-ADRootDSE

    Write-Step 'ユーザーを取得しています...'
    $userProps = @(
        'LastLogonTimestamp','PasswordLastSet','PasswordNeverExpires','PasswordNotRequired',
        'Enabled','whenCreated','adminCount','ServicePrincipalName','TrustedForDelegation',
        'DoesNotRequirePreAuth','AllowReversiblePasswordEncryption','AccountNotDelegated','MemberOf'
    )
    $users = Get-ADUser -Filter * -Properties $userProps

    Write-Step 'コンピューターを取得しています...'
    $compProps = @('LastLogonTimestamp','Enabled','whenCreated','OperatingSystem','OperatingSystemVersion','TrustedForDelegation','PasswordLastSet')
    $computers = Get-ADComputer -Filter * -Properties $compProps

    Write-Step 'ドメインコントローラーを取得しています...'
    $dcs = @(Get-ADDomainController -Filter *)

    Write-Step '特権グループを取得しています...'
    $privGroupNames = @(
        'Domain Admins','Enterprise Admins','Schema Admins','Administrators',
        'Account Operators','Backup Operators','Server Operators','Print Operators'
    )
    $privGroups = foreach ($n in $privGroupNames) {
        $g = Get-ADGroup -Filter "Name -eq '$n'" -ErrorAction SilentlyContinue
        if ($g) {
            $members = @()
            try { $members = @(Get-ADGroupMember -Identity $g -Recursive -ErrorAction Stop) } catch { }
            [pscustomobject]@{ Name = $n; Members = $members }
        }
    }

    Write-Step 'パスワードポリシーを取得しています...'
    $pwdPolicy = Get-ADDefaultDomainPasswordPolicy

    Write-Step 'サイト構成を取得しています...'
    $configNC = $rootDse.configurationNamingContext
    $sites   = @(Get-ADObject -SearchBase "CN=Sites,$configNC" -LDAPFilter '(objectClass=site)' -Properties name,distinguishedName -ErrorAction SilentlyContinue)
    $subnets = @(Get-ADObject -SearchBase "CN=Subnets,CN=Sites,$configNC" -LDAPFilter '(objectClass=subnet)' -Properties name,siteObject -ErrorAction SilentlyContinue)

    # AD ごみ箱
    $recycleEnabled = $false
    try {
        $rb = Get-ADOptionalFeature -Filter "Name -like 'Recycle Bin Feature'" -ErrorAction Stop
        $recycleEnabled = @($rb.EnabledScopes).Count -gt 0
    } catch { }

    # krbtgt
    $krbtgt = $null
    try { $krbtgt = Get-ADUser -Identity 'krbtgt' -Properties PasswordLastSet -ErrorAction Stop } catch { }

    # ビルトイン Administrator（RID 500）
    $builtinAdmin = $users | Where-Object { $_.SID.Value -like '*-500' } | Select-Object -First 1

    # LAPS 属性の有無（スキーマ拡張されているか）
    $lapsSchema = $false
    try {
        $schemaNC = $rootDse.schemaNamingContext
        $lapsAttr = Get-ADObject -SearchBase $schemaNC -LDAPFilter '(|(lDAPDisplayName=ms-Mcs-AdmPwd)(lDAPDisplayName=msLAPS-Password))' -ErrorAction SilentlyContinue
        $lapsSchema = $null -ne $lapsAttr
    } catch { }

    [pscustomobject]@{
        IsDemo          = $false
        Now             = $Now
        DomainName      = $domain.DNSRoot
        NetBIOSName     = $domain.NetBIOSName
        DomainMode      = [string]$domain.DomainMode
        ForestMode      = [string]$forest.ForestMode
        Users           = $users
        Computers       = $computers
        DomainControllers = $dcs
        PrivilegedGroups = @($privGroups)
        PasswordPolicy  = $pwdPolicy
        Sites           = $sites
        Subnets         = $subnets
        RecycleBinEnabled = $recycleEnabled
        Krbtgt          = $krbtgt
        BuiltinAdmin    = $builtinAdmin
        LapsSchema      = $lapsSchema
    }
}

#endregion

#region ---------- チェック（各関数は Finding を 0..n 件返す） ----------

function Test-KrbtgtPassword {
    param($Snap)
    if (-not $Snap.Krbtgt) { return }
    $days = Get-DaysAgo -When $Snap.Krbtgt.PasswordLastSet -Now $Snap.Now
    if ($null -eq $days -or $days -lt $StalePasswordDays) { return }
    New-Finding -Category '特権・認証' -Title 'krbtgt アカウントのパスワードが長期間変更されていません' `
        -Severity $SEV_HIGH -Phase $PHASE1 `
        -Detail "krbtgt のパスワード最終変更から $days 日が経過しています（基準: $StalePasswordDays 日）。krbtgt は Kerberos チケットの署名に使われるアカウントで、古いままだと Golden Ticket 攻撃の成立余地が残ります。" `
        -Impact 'なし（正しい手順で実施すれば、利用者への影響はありません）' `
        -Action 'krbtgt のパスワードを変更します。2 回変更が必要ですが、1 回目と 2 回目の間は最低でも TGT の最大有効期間（既定 10 時間）以上あけてください。連続して 2 回リセットすると Kerberos 認証が壊れます。' `
        -Reference 'https://learn.microsoft.com/ja-jp/defender-for-identity/security-posture-assessments/accounts' `
        -Items @([pscustomobject]@{ 対象='krbtgt'; 最終変更=$Snap.Krbtgt.PasswordLastSet; 経過日数=$days })
}

function Test-BuiltinAdminPassword {
    param($Snap)
    if (-not $Snap.BuiltinAdmin) { return }
    $days = Get-DaysAgo -When $Snap.BuiltinAdmin.PasswordLastSet -Now $Snap.Now
    if ($null -eq $days -or $days -lt $StalePasswordDays) { return }
    New-Finding -Category '特権・認証' -Title 'ビルトイン Administrator のパスワードが長期間変更されていません' `
        -Severity $SEV_HIGH -Phase $PHASE2 `
        -Detail "ビルトイン Administrator（RID 500）のパスワード最終変更から $days 日が経過しています。攻撃時に最初に狙われるアカウントです。" `
        -Impact '要確認（タスクスケジューラーやサービスがこのアカウントを使っていると停止します）' `
        -Action '使用箇所を洗い出したうえでパスワードを変更します。通常運用でビルトイン Administrator を使わない設計にしておくことも併せて検討してください。' `
        -Reference 'https://learn.microsoft.com/ja-jp/defender-for-identity/security-posture-assessments/accounts' `
        -Items @([pscustomobject]@{ 対象=(Protect-Name $Snap.BuiltinAdmin.SamAccountName); 最終変更=$Snap.BuiltinAdmin.PasswordLastSet; 経過日数=$days })
}

function Test-PrivilegedGroupSize {
    param($Snap)
    $rows = foreach ($g in $Snap.PrivilegedGroups) {
        [pscustomobject]@{ グループ=$g.Name; メンバー数=@($g.Members).Count }
    }
    $big = @($rows | Where-Object { $_.グループ -in @('Domain Admins','Enterprise Admins','Schema Admins') -and $_.メンバー数 -gt 5 })
    if ($big.Count -eq 0) { return }
    New-Finding -Category '特権・認証' -Title '特権グループのメンバーが多すぎます' `
        -Severity $SEV_HIGH -Phase $PHASE2 `
        -Detail ('Domain Admins / Enterprise Admins / Schema Admins のいずれかで、メンバー数が 5 を超えています。特権アカウントは 1 つ侵害されれば全体が侵害されるため、数そのものがリスクになります。') `
        -Impact '要確認（実際に必要としている担当者を外すと運用が止まります）' `
        -Action '棚卸しを行い、常時必要な人だけを残します。日常業務用アカウントと管理用アカウントを分け、必要なときだけ昇格する運用へ移行してください。' `
        -Reference 'https://learn.microsoft.com/ja-jp/windows-server/identity/ad-ds/plan/security-best-practices/appendix-b--privileged-accounts-and-groups-in-active-directory' `
        -Items $rows
}

function Test-UnconstrainedDelegation {
    param($Snap)
    $items = @()
    $items += $Snap.Computers | Where-Object { $_.TrustedForDelegation -and $_.Enabled } |
        ForEach-Object { [pscustomobject]@{ 種別='コンピューター'; 名前=(Protect-Name $_.Name); OS=$_.OperatingSystem } }
    $items += $Snap.Users | Where-Object { $_.TrustedForDelegation -and $_.Enabled } |
        ForEach-Object { [pscustomobject]@{ 種別='ユーザー'; 名前=(Protect-Name $_.SamAccountName); OS='' } }
    # ドメインコントローラーは既定で制約なし委任が設定されるため除外
    $dcNames = @($Snap.DomainControllers | ForEach-Object { Protect-Name $_.Name })
    $items = @($items | Where-Object { $_.名前 -notin $dcNames })
    if ($items.Count -eq 0) { return }
    New-Finding -Category '特権・認証' -Title '制約なし委任（Unconstrained Delegation）が設定されています' `
        -Severity $SEV_HIGH -Phase $PHASE2 `
        -Detail ('DC 以外で制約なし委任が設定されたオブジェクトが {0} 件あります。このホストが侵害されると、接続してきた利用者のチケットが取得され、権限が横取りされます。' -f $items.Count) `
        -Impact '要確認（該当システムの動作要件を確認してから変更してください）' `
        -Action '制約付き委任（Constrained Delegation）またはリソースベースの制約付き委任へ移行します。不要であれば設定を解除してください。' `
        -Reference 'https://learn.microsoft.com/ja-jp/defender-for-identity/security-posture-assessments/unconstrained-kerberos' `
        -Items $items
}

function Test-KerberoastableAccounts {
    param($Snap)
    $items = @($Snap.Users | Where-Object {
            $_.Enabled -and $_.ServicePrincipalName -and @($_.ServicePrincipalName).Count -gt 0 -and $_.SamAccountName -ne 'krbtgt'
        } | ForEach-Object {
            $days = Get-DaysAgo -When $_.PasswordLastSet -Now $Snap.Now
            [pscustomobject]@{ アカウント=(Protect-Name $_.SamAccountName); SPN数=@($_.ServicePrincipalName).Count; パスワード経過日数=$days; 特権=([bool]$_.adminCount) }
        })
    if ($items.Count -eq 0) { return }
    $priv = @($items | Where-Object { $_.特権 }).Count
    $sev = if ($priv -gt 0) { $SEV_HIGH } else { $SEV_MED }
    New-Finding -Category '特権・認証' -Title 'SPN が設定されたユーザーアカウント（Kerberoasting の標的）があります' `
        -Severity $sev -Phase $PHASE3 `
        -Detail ('SPN を持つユーザーアカウントが {0} 件あります（うち特権アカウント {1} 件）。これらはオフラインでのパスワード解析を試みられる対象になります。' -f $items.Count, $priv) `
        -Impact '要確認（サービスの認証に使われているため、変更時は停止を伴う可能性があります）' `
        -Action 'グループ管理サービスアカウント（gMSA）への移行を検討します。難しい場合は、25 文字以上のランダムなパスワードを設定し、特権グループから外してください。' `
        -Reference 'https://learn.microsoft.com/ja-jp/defender-for-identity/security-posture-assessments/riskiest-lateral-movement-paths' `
        -Items $items
}

function Test-AsRepRoastable {
    param($Snap)
    $items = @($Snap.Users | Where-Object { $_.Enabled -and $_.DoesNotRequirePreAuth } |
        ForEach-Object { [pscustomobject]@{ アカウント=(Protect-Name $_.SamAccountName); 作成日=$_.whenCreated } })
    if ($items.Count -eq 0) { return }
    New-Finding -Category '特権・認証' -Title 'Kerberos 事前認証が無効なアカウントがあります' `
        -Severity $SEV_HIGH -Phase $PHASE2 `
        -Detail ('事前認証が不要な設定のアカウントが {0} 件あります。パスワードハッシュを外部から取得され、オフラインで解析される恐れがあります（AS-REP Roasting）。' -f $items.Count) `
        -Impact '要確認（古いシステムの都合で設定されている場合があります）' `
        -Action '設定が必要な理由を確認し、不要であれば事前認証を有効に戻してください。' `
        -Reference 'https://learn.microsoft.com/ja-jp/defender-for-identity/security-posture-assessments/preauthentication-not-required' `
        -Items $items
}

function Test-PasswordHygiene {
    param($Snap)
    $out = @()

    $notReq = @($Snap.Users | Where-Object { $_.Enabled -and $_.PasswordNotRequired } |
        ForEach-Object { [pscustomobject]@{ アカウント=(Protect-Name $_.SamAccountName); 作成日=$_.whenCreated } })
    if ($notReq.Count -gt 0) {
        $out += New-Finding -Category 'アカウント衛生' -Title 'パスワード不要の設定になっているアカウントがあります' `
            -Severity $SEV_HIGH -Phase $PHASE2 `
            -Detail ('「パスワードを必要としない」が有効なアカウントが {0} 件あります。空パスワードでログオンできる状態になり得ます。' -f $notReq.Count) `
            -Impact '要確認（該当アカウントの利用状況を確認してください）' `
            -Action '設定を解除し、パスワードを設定します。使われていないアカウントであれば無効化してください。' `
            -Reference 'https://learn.microsoft.com/ja-jp/windows-server/identity/ad-ds/manage/understand-security-identifiers' `
            -Items $notReq
    }

    $rev = @($Snap.Users | Where-Object { $_.Enabled -and $_.AllowReversiblePasswordEncryption } |
        ForEach-Object { [pscustomobject]@{ アカウント=(Protect-Name $_.SamAccountName) } })
    if ($rev.Count -gt 0) {
        $out += New-Finding -Category 'アカウント衛生' -Title 'パスワードが可逆暗号化で保存されています' `
            -Severity $SEV_HIGH -Phase $PHASE2 `
            -Detail ('可逆暗号化が有効なアカウントが {0} 件あります。実質的に平文と同じ扱いで、AD が侵害された際にそのまま読み取られます。' -f $rev.Count) `
            -Impact '要確認（この設定を必要とする古い認証方式が使われている可能性があります）' `
            -Action '依存しているアプリケーションを確認し、可能であれば設定を解除してパスワードを再設定します。' `
            -Reference 'https://learn.microsoft.com/ja-jp/windows-server/security/kerberos/store-passwords-using-reversible-encryption' `
            -Items $rev
    }

    $neverExp = @($Snap.Users | Where-Object { $_.Enabled -and $_.PasswordNeverExpires } |
        ForEach-Object {
            $days = Get-DaysAgo -When $_.PasswordLastSet -Now $Snap.Now
            [pscustomobject]@{ アカウント=(Protect-Name $_.SamAccountName); パスワード経過日数=$days; 特権=([bool]$_.adminCount) }
        })
    if ($neverExp.Count -gt 0) {
        $privCount = @($neverExp | Where-Object { $_.特権 }).Count
        $out += New-Finding -Category 'アカウント衛生' -Title 'パスワード無期限のアカウントがあります' `
            -Severity ($(if ($privCount -gt 0) { $SEV_HIGH } else { $SEV_MED })) -Phase $PHASE3 `
            -Detail ('パスワード無期限のアカウントが {0} 件あります（うち特権アカウント {1} 件）。多くはサービスアカウントですが、棚卸しされないまま残っていることがあります。' -f $neverExp.Count, $privCount) `
            -Impact '要確認（サービスアカウントの場合、変更時に停止を伴います）' `
            -Action '用途を台帳化し、可能なものは gMSA へ移行します。人が使うアカウントであれば無期限設定を解除してください。' `
            -Items $neverExp
    }

    return $out
}

function Test-StaleObjects {
    param($Snap)
    $out = @()
    $threshold = $Snap.Now.AddDays(-1 * $InactiveDays)

    $staleUsers = @($Snap.Users | Where-Object {
            $_.Enabled -and $_.LastLogonTimestamp -and
            ([datetime]::FromFileTime($_.LastLogonTimestamp) -lt $threshold)
        } | ForEach-Object {
            $last = [datetime]::FromFileTime($_.LastLogonTimestamp)
            [pscustomobject]@{ アカウント=(Protect-Name $_.SamAccountName); 最終ログオン=$last; 経過日数=(Get-DaysAgo -When $last -Now $Snap.Now) }
        })
    if ($staleUsers.Count -gt 0) {
        $out += New-Finding -Category '棚卸し' -Title '長期間使われていないユーザーアカウントがあります' `
            -Severity $SEV_MED -Phase $PHASE3 `
            -Detail ('{0} 日以上ログオンしていない有効なユーザーが {1} 件あります。退職者が残っている場合、そのまま侵入口になります。' -f $InactiveDays, $staleUsers.Count) `
            -Impact '要確認（人事情報との突き合わせが必要です）' `
            -Action '人事部門と照合し、不要なものは無効化したうえで一定期間後に削除します。まず無効化して様子を見る運用が安全です。' `
            -Items $staleUsers
    }

    $staleComps = @($Snap.Computers | Where-Object {
            $_.Enabled -and $_.LastLogonTimestamp -and
            ([datetime]::FromFileTime($_.LastLogonTimestamp) -lt $threshold)
        } | ForEach-Object {
            $last = [datetime]::FromFileTime($_.LastLogonTimestamp)
            [pscustomobject]@{ コンピューター=(Protect-Name $_.Name); OS=$_.OperatingSystem; 最終ログオン=$last; 経過日数=(Get-DaysAgo -When $last -Now $Snap.Now) }
        })
    if ($staleComps.Count -gt 0) {
        $out += New-Finding -Category '棚卸し' -Title '長期間使われていないコンピューターアカウントがあります' `
            -Severity $SEV_LOW -Phase $PHASE3 `
            -Detail ('{0} 日以上ログオンしていない有効なコンピューターが {1} 件あります。廃棄済み端末の残骸である可能性があります。' -f $InactiveDays, $staleComps.Count) `
            -Impact 'なし（無効化から始めれば影響を確認しながら進められます）' `
            -Action '無効化して専用 OU へ移動し、一定期間問題がなければ削除します。' `
            -Items $staleComps
    }

    return $out
}

function Test-LegacyOperatingSystem {
    param($Snap)
    $legacyPattern = 'Windows (2000|XP|Vista|7|8|Server 2003|Server 2008|Server 2012)'
    $items = @($Snap.Computers | Where-Object {
            $_.Enabled -and $_.OperatingSystem -and ($_.OperatingSystem -match $legacyPattern)
        } | ForEach-Object { [pscustomobject]@{ コンピューター=(Protect-Name $_.Name); OS=$_.OperatingSystem } })
    if ($items.Count -eq 0) { return }
    New-Finding -Category '基盤' -Title 'サポートが終了した OS が残っています' `
        -Severity $SEV_HIGH -Phase $PHASE2 `
        -Detail ('サポート終了済みと思われる OS の有効なコンピューターが {0} 件あります。更新プログラムが提供されないため、既知の脆弱性が残り続けます。' -f $items.Count) `
        -Impact '要確認（業務システムが載っている場合は移行計画が必要です）' `
        -Action '用途を確認し、更新・移行・隔離のいずれかを計画します。どうしても残す場合はネットワークを分離してください。' `
        -Items $items
}

function Test-DomainConfiguration {
    param($Snap)
    $out = @()

    if (-not $Snap.RecycleBinEnabled) {
        $out += New-Finding -Category '基盤' -Title 'AD ごみ箱が有効になっていません' `
            -Severity $SEV_MED -Phase $PHASE1 `
            -Detail 'Active Directory ごみ箱が無効です。オブジェクトを誤って削除した場合、復元にバックアップからのリストアが必要になります。' `
            -Impact 'なし（有効化しても既存の動作に影響しません）' `
            -Action '有効化を検討してください。ただし一度有効にすると無効化できない点に注意が必要です。' `
            -Reference 'https://learn.microsoft.com/ja-jp/windows-server/identity/ad-ds/get-started/adac/introduction-to-active-directory-administrative-center-enhancements--level-100-#enable-active-directory-recycle-bin' `
            -Items @([pscustomobject]@{ 項目='AD ごみ箱'; 状態='無効' })
    }

    $pp = $Snap.PasswordPolicy
    $issues = @()
    if ($pp.MinPasswordLength -lt 14) { $issues += [pscustomobject]@{ 項目='最小パスワード長'; 現在値=$pp.MinPasswordLength; 推奨='14 文字以上' } }
    if ($pp.LockoutThreshold -eq 0)   { $issues += [pscustomobject]@{ 項目='ロックアウトしきい値'; 現在値='未設定'; 推奨='10 回以下' } }
    if (-not $pp.ComplexityEnabled)   { $issues += [pscustomobject]@{ 項目='複雑さの要件'; 現在値='無効'; 推奨='有効' } }
    if ($issues.Count -gt 0) {
        $out += New-Finding -Category '基盤' -Title 'パスワードポリシーが推奨値を下回っています' `
            -Severity $SEV_MED -Phase $PHASE2 `
            -Detail ('既定のドメインパスワードポリシーに、推奨値を下回る項目が {0} 件あります。' -f $issues.Count) `
            -Impact '要確認（次回のパスワード変更時から利用者に影響します。事前告知を推奨します）' `
            -Action '段階的に引き上げます。長さを優先し、定期変更の強制よりも長さと漏えい対策を重視する方針が現在は推奨されています。' `
            -Reference 'https://learn.microsoft.com/ja-jp/windows-server/identity/ad-ds/get-started/adac/introduction-to-active-directory-administrative-center-enhancements--level-100-' `
            -Items $issues
    }

    if (-not $Snap.LapsSchema) {
        $out += New-Finding -Category '基盤' -Title 'LAPS（ローカル管理者パスワードの自動管理）が導入されていません' `
            -Severity $SEV_MED -Phase $PHASE2 `
            -Detail 'LAPS 用のスキーマ属性が見つかりません。端末のローカル管理者パスワードが共通になっている場合、1 台の侵害が全台に波及します。' `
            -Impact '要確認（導入時に端末側の設定変更が必要です）' `
            -Action 'Windows LAPS の導入を検討してください。端末ごとに異なるパスワードが自動管理されるようになります。' `
            -Reference 'https://learn.microsoft.com/ja-jp/windows-server/identity/laps/laps-overview' `
            -Items @([pscustomobject]@{ 項目='Windows LAPS'; 状態='未導入と推定' })
    }

    return $out
}

function Test-SiteTopology {
    param($Snap)
    $out = @()

    $siteNames = @($Snap.Sites | ForEach-Object { $_.Name })
    $orphanSubnets = @($Snap.Subnets | Where-Object { -not $_.siteObject } |
        ForEach-Object { [pscustomobject]@{ サブネット=$_.Name; 割当サイト='(なし)' } })
    if ($orphanSubnets.Count -gt 0) {
        $out += New-Finding -Category 'サイト構成' -Title 'どのサイトにも割り当てられていないサブネットがあります' `
            -Severity $SEV_LOW -Phase $PHASE3 `
            -Detail ('サイト未割当のサブネットが {0} 件あります。該当ネットワークの端末が、最寄りではない DC に接続してしまう可能性があります。' -f $orphanSubnets.Count) `
            -Impact 'なし（割り当ての追加は影響が限定的です）' `
            -Action '適切なサイトへ割り当てるか、使われていないサブネット定義であれば削除します。' `
            -Items $orphanSubnets
    }

    $usedSites = @($Snap.DomainControllers | ForEach-Object { $_.Site } | Sort-Object -Unique)
    $emptySites = @($siteNames | Where-Object { $_ -notin $usedSites } | ForEach-Object { [pscustomobject]@{ サイト=$_; 状態='DC が存在しません' } })
    if ($emptySites.Count -gt 0) {
        $out += New-Finding -Category 'サイト構成' -Title 'ドメインコントローラーが存在しないサイトがあります' `
            -Severity $SEV_LOW -Phase $PHASE3 `
            -Detail ('DC が配置されていないサイトが {0} 件あります。拠点閉鎖後に定義だけが残っているケースがよくあります。' -f $emptySites.Count) `
            -Impact 'なし' `
            -Action '意図した構成でなければ、サイト定義を整理してください。' `
            -Items $emptySites
    }

    return $out
}

$script:CheckRegistry = @(
    'Test-KrbtgtPassword'
    'Test-BuiltinAdminPassword'
    'Test-PrivilegedGroupSize'
    'Test-UnconstrainedDelegation'
    'Test-KerberoastableAccounts'
    'Test-AsRepRoastable'
    'Test-PasswordHygiene'
    'Test-StaleObjects'
    'Test-LegacyOperatingSystem'
    'Test-DomainConfiguration'
    'Test-SiteTopology'
)

#endregion

#region ---------- スコア ----------

<#
    100 点からの減点方式。High=12点 / Medium=5点 / Low=2点。
    ただし重要度ごとに減点の上限を設ける（High 48 / Medium 25 / Low 12）。
    上限がないと、指摘が多い環境がすべて 0 点になってしまい、
    改善しても点が動かず「比較のための目安」として機能しないため。
#>
function Get-HealthScore {
    param([object[]]$Findings)

    $weights = @{ High = 12; Medium = 5;  Low = 2  }
    $caps    = @{ High = 48; Medium = 25; Low = 12 }

    $deduction = 0
    foreach ($sev in @('High','Medium','Low')) {
        $n = @($Findings | Where-Object Severity -eq $sev).Count
        $deduction += [math]::Min($caps[$sev], $n * $weights[$sev])
    }
    return [math]::Max(0, 100 - $deduction)
}

function Get-ScoreRank {
    param([int]$Score)
    if ($Score -ge 85) { return @{ Rank='A'; Text='おおむね良好です'; Color='#1f7a70' } }
    if ($Score -ge 65) { return @{ Rank='B'; Text='改善の余地があります'; Color='#c78a1e' } }
    if ($Score -ge 40) { return @{ Rank='C'; Text='早めの対処を推奨します'; Color='#e2621f' } }
    return @{ Rank='D'; Text='優先的な対処が必要です'; Color='#c0392b' }
}

#endregion

#region ---------- HTML 生成 ----------

function ConvertTo-ItemTable {
    param([object[]]$Items)
    if (-not $Items -or $Items.Count -eq 0) { return '' }
    $shown = $Items | Select-Object -First 50
    $cols = @($shown[0].PSObject.Properties.Name)

    $sb = [Text.StringBuilder]::new()
    [void]$sb.Append('<table class="items"><thead><tr>')
    foreach ($c in $cols) { [void]$sb.Append('<th>' + (ConvertTo-HtmlText $c) + '</th>') }
    [void]$sb.Append('</tr></thead><tbody>')
    foreach ($row in $shown) {
        [void]$sb.Append('<tr>')
        foreach ($c in $cols) {
            $v = $row.$c
            if ($v -is [datetime]) { $v = $v.ToString('yyyy-MM-dd') }
            [void]$sb.Append('<td>' + (ConvertTo-HtmlText ([string]$v)) + '</td>')
        }
        [void]$sb.Append('</tr>')
    }
    [void]$sb.Append('</tbody></table>')
    if ($Items.Count -gt 50) {
        [void]$sb.Append('<p class="more">ほか ' + ($Items.Count - 50) + ' 件（全件は CSV をご確認ください）</p>')
    }
    return $sb.ToString()
}

function New-HtmlReport {
    param(
        [Parameter(Mandatory)]$Snap,
        [Parameter(Mandatory)][object[]]$Findings,
        [Parameter(Mandatory)][string]$Company
    )

    $score = Get-HealthScore -Findings $Findings
    $rank  = Get-ScoreRank -Score $score
    $high  = @($Findings | Where-Object Severity -eq 'High').Count
    $med   = @($Findings | Where-Object Severity -eq 'Medium').Count
    $low   = @($Findings | Where-Object Severity -eq 'Low').Count
    $generated = $Snap.Now.ToString('yyyy年MM月dd日 HH:mm')

    $sb = [Text.StringBuilder]::new()

    # --- 各フェーズの本文 ---
    $phaseHtml = [Text.StringBuilder]::new()
    foreach ($pkey in $script:PhaseMeta.Keys) {
        $meta = $script:PhaseMeta[$pkey]
        $inPhase = @($Findings | Where-Object Phase -eq $pkey |
            Sort-Object @{ Expression = { switch ($_.Severity) { 'High' {0} 'Medium' {1} default {2} } } })
        [void]$phaseHtml.Append('<section class="phase"><div class="phase-head"><span class="phase-badge ' + $pkey.ToLower() + '">' + $meta.Label + '</span><div><h2>' + (ConvertTo-HtmlText $meta.Title) + '</h2><p>' + (ConvertTo-HtmlText $meta.Desc) + '</p></div><span class="phase-count">' + $inPhase.Count + ' 件</span></div>')
        if ($inPhase.Count -eq 0) {
            [void]$phaseHtml.Append('<p class="empty">このフェーズで検出された項目はありません。</p>')
        }
        foreach ($f in $inPhase) {
            $sevClass = $f.Severity.ToLower()
            [void]$phaseHtml.Append('<article class="finding">')
            [void]$phaseHtml.Append('<header><span class="sev ' + $sevClass + '">' + $f.Severity + '</span><h3>' + (ConvertTo-HtmlText $f.Title) + '</h3><span class="cat">' + (ConvertTo-HtmlText $f.Category) + '</span></header>')
            [void]$phaseHtml.Append('<p class="detail">' + (ConvertTo-HtmlText $f.Detail) + '</p>')
            [void]$phaseHtml.Append('<dl><dt>利用者への影響</dt><dd>' + (ConvertTo-HtmlText $f.Impact) + '</dd>')
            if ($f.Action) { [void]$phaseHtml.Append('<dt>推奨する対処</dt><dd>' + (ConvertTo-HtmlText $f.Action) + '</dd>') }
            [void]$phaseHtml.Append('</dl>')
            if ($f.Items -and $f.Items.Count -gt 0) {
                [void]$phaseHtml.Append('<details><summary>該当 ' + $f.Items.Count + ' 件を表示</summary>' + (ConvertTo-ItemTable -Items $f.Items) + '</details>')
            }
            if ($f.Reference) {
                [void]$phaseHtml.Append('<p class="ref">参考: <a href="' + (ConvertTo-HtmlText $f.Reference) + '" target="_blank" rel="noopener">' + (ConvertTo-HtmlText $f.Reference) + '</a></p>')
            }
            [void]$phaseHtml.Append('</article>')
        }
        [void]$phaseHtml.Append('</section>')
    }

    # --- カテゴリ別集計 ---
    $catRows = $Findings | Group-Object Category | Sort-Object Count -Descending | ForEach-Object {
        '<tr><td>' + (ConvertTo-HtmlText $_.Name) + '</td><td class="num">' + $_.Count + '</td></tr>'
    }

    $demoBanner = if ($Snap.IsDemo) {
        '<div class="demo-banner">これは <strong>サンプルレポート</strong> です。架空の組織のデータを用いて、実際の出力形式をお見せしています。</div>'
    } else { '' }

    $anonNote = if ($Anonymize) {
        '<p class="anon-note">🔒 このレポートは<strong>匿名化モード</strong>で生成されています。アカウント名・コンピューター名は伏せ字（同一の対象は同一の伏せ字）に置き換えられています。</p>'
    } else { '' }

    [void]$sb.Append(@"
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Active Directory 健全性診断レポート － $(ConvertTo-HtmlText $Company)</title>
<style>
  :root{--ink:#161d2b;--soft:#3d4759;--paper:#faf7f2;--paper2:#f2ede4;--card:#fff;--line:#e3dccf;--accent:#e2621f;--teal:#1f7a70;--high:#c0392b;--med:#c78a1e;--low:#7a8290}
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:"Segoe UI","Yu Gothic UI","Noto Sans JP",sans-serif;color:var(--ink);background:var(--paper);font-size:15px;line-height:1.85;-webkit-font-smoothing:antialiased}
  .wrap{max-width:1000px;margin:0 auto;padding:0 24px}
  h1,h2,h3{line-height:1.4}
  .demo-banner{background:#161d2b;color:#ffc79a;text-align:center;padding:10px 16px;font-size:13.5px;font-weight:600}
  header.top{background:linear-gradient(160deg,#1c2436,#161d2b);color:#f4ede2;padding:38px 0 34px}
  header.top .kicker{font-size:12.5px;letter-spacing:.16em;color:#ffb07a;font-weight:700;margin-bottom:10px}
  header.top h1{font-size:30px;font-weight:800;margin-bottom:8px}
  header.top .meta{font-size:13.5px;color:#c1b8ac}
  .score-wrap{display:grid;grid-template-columns:230px 1fr;gap:22px;margin:-26px auto 0;position:relative}
  .score{background:var(--card);border:1px solid var(--line);border-radius:16px;padding:24px;text-align:center;box-shadow:0 14px 34px -22px rgba(22,29,43,.5)}
  .score .val{font-size:56px;font-weight:800;line-height:1}
  .score .rank{font-size:13px;font-weight:800;margin-top:6px}
  .score .lbl{font-size:12px;color:var(--soft);margin-top:10px;border-top:1px solid var(--line);padding-top:10px}
  .sum{background:var(--card);border:1px solid var(--line);border-radius:16px;padding:24px 26px;box-shadow:0 14px 34px -22px rgba(22,29,43,.5)}
  .sum h2{font-size:17px;margin-bottom:12px}
  .counts{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:14px}
  .counts span{font-size:13px;font-weight:700;padding:6px 13px;border-radius:8px;border:1px solid var(--line)}
  .counts .h{background:#fdeceb;color:var(--high);border-color:#f3cfcc}
  .counts .m{background:#fdf5e6;color:#9a6a12;border-color:#eeddb8}
  .counts .l{background:#f1f2f4;color:var(--low);border-color:#dfe1e5}
  .sum p{font-size:14px;color:var(--soft)}
  table.kv{width:100%;border-collapse:collapse;margin-top:14px;font-size:13.5px}
  table.kv td{padding:5px 0;border-bottom:1px dashed var(--line)}
  table.kv td:last-child{text-align:right;font-weight:700}
  section.block{margin-top:44px}
  section.block>h2{font-size:20px;margin-bottom:6px}
  section.block>p.lead{font-size:14.5px;color:var(--soft);margin-bottom:18px}
  .phase{margin-top:34px}
  .phase-head{display:flex;align-items:flex-start;gap:16px;background:var(--paper2);border:1px solid var(--line);border-radius:14px;padding:18px 20px}
  .phase-badge{flex-shrink:0;font-size:12.5px;font-weight:800;color:#fff;padding:6px 13px;border-radius:999px}
  .phase-badge.phase1{background:var(--teal)}
  .phase-badge.phase2{background:var(--accent)}
  .phase-badge.phase3{background:#6b7280}
  .phase-head h2{font-size:17.5px;margin-bottom:4px}
  .phase-head p{font-size:13.5px;color:var(--soft)}
  .phase-count{margin-left:auto;flex-shrink:0;font-size:13px;font-weight:800;color:var(--soft)}
  .finding{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:22px 24px;margin-top:14px}
  .finding header{display:flex;align-items:center;gap:11px;flex-wrap:wrap;margin-bottom:10px}
  .finding h3{font-size:16.5px;font-weight:800}
  .sev{font-size:11.5px;font-weight:800;color:#fff;padding:3px 9px;border-radius:6px}
  .sev.high{background:var(--high)}.sev.medium{background:var(--med)}.sev.low{background:var(--low)}
  .cat{margin-left:auto;font-size:12px;font-weight:700;color:var(--soft);background:var(--paper2);padding:3px 10px;border-radius:6px}
  .detail{font-size:14.5px;color:var(--soft);margin-bottom:12px}
  .finding dl{display:grid;grid-template-columns:130px 1fr;gap:6px 14px;font-size:14px;background:var(--paper);border-radius:10px;padding:14px 16px}
  .finding dt{font-weight:800;color:var(--ink)}
  .finding dd{color:var(--soft)}
  details{margin-top:12px}
  summary{cursor:pointer;font-size:13.5px;font-weight:700;color:var(--accent)}
  table.items{width:100%;border-collapse:collapse;margin-top:10px;font-size:13px}
  table.items th{text-align:left;background:var(--paper2);padding:7px 9px;border:1px solid var(--line);font-weight:700}
  table.items td{padding:6px 9px;border:1px solid var(--line)}
  .more{font-size:12.5px;color:var(--soft);margin-top:6px}
  .ref{margin-top:10px;font-size:12.5px;word-break:break-all}
  .ref a{color:var(--teal)}
  .empty{font-size:14px;color:var(--soft);padding:14px 4px}
  .anon-note{font-size:13px;color:var(--soft);background:#eef4f3;border:1px solid #cfe3e0;border-radius:10px;padding:12px 16px;margin-top:16px}
  .readonly{margin-top:20px;background:#eef4f3;border:1px solid #cfe3e0;border-radius:12px;padding:16px 18px;font-size:13.5px;color:#1f4a45}
  footer{margin-top:52px;background:#161d2b;color:#9d958a;font-size:12.5px;padding:26px 0}
  footer strong{color:#e6ded3}
  @media print{body{background:#fff}.finding,.sum,.score{box-shadow:none}details{display:none}}
  @media(max-width:760px){.score-wrap{grid-template-columns:1fr;margin-top:20px}.finding dl{grid-template-columns:1fr}}
</style>
</head>
<body>
$demoBanner
<header class="top">
  <div class="wrap">
    <div class="kicker">ACTIVE DIRECTORY HEALTH REPORT</div>
    <h1>Active Directory 健全性診断レポート</h1>
    <div class="meta">$(ConvertTo-HtmlText $Company) ／ 生成日時: $generated</div>
  </div>
</header>

<div class="wrap">
  <div class="score-wrap">
    <div class="score">
      <div class="val" style="color:$($rank.Color)">$score</div>
      <div class="rank" style="color:$($rank.Color)">評価 $($rank.Rank) ／ $($rank.Text)</div>
      <div class="lbl">重要度に応じた減点方式（重要度ごとに上限あり）。絶対的な指標ではなく、<strong>対処前後の比較</strong>にお使いください。</div>
    </div>
    <div class="sum">
      <h2>検出のサマリー</h2>
      <div class="counts">
        <span class="h">重要度 High: $high 件</span>
        <span class="m">Medium: $med 件</span>
        <span class="l">Low: $low 件</span>
      </div>
      <p>本レポートは <strong>読み取り専用</strong> の検査により作成しています。以下では検出結果を「<strong>利用者への影響があるかどうか</strong>」で 3 つのフェーズに分け、着手しやすい順に並べています。</p>
      <table class="kv">
        <tr><td>ドメイン</td><td>$(ConvertTo-HtmlText $Snap.DomainName)</td></tr>
        <tr><td>ドメインの機能レベル</td><td>$(ConvertTo-HtmlText $Snap.DomainMode)</td></tr>
        <tr><td>フォレストの機能レベル</td><td>$(ConvertTo-HtmlText $Snap.ForestMode)</td></tr>
        <tr><td>ドメインコントローラー</td><td>$(@($Snap.DomainControllers).Count) 台</td></tr>
        <tr><td>ユーザー / コンピューター</td><td>$(@($Snap.Users).Count) / $(@($Snap.Computers).Count)</td></tr>
      </table>
    </div>
  </div>

  $anonNote

  <div class="readonly">
    🔒 <strong>この診断は読み取り専用です。</strong>使用したのは情報を取得するコマンドのみで、Active Directory の構成を変更する処理は含まれていません。
  </div>

  <section class="block">
    <h2>まず、何から手を付けるか</h2>
    <p class="lead">検出件数の多さより、着手の順序が大切です。フェーズ 1 から順に進めることをお勧めします。</p>
    $($phaseHtml.ToString())
  </section>

  <section class="block">
    <h2>カテゴリ別の内訳</h2>
    <table class="items" style="max-width:460px">
      <thead><tr><th>カテゴリ</th><th>件数</th></tr></thead>
      <tbody>$($catRows -join '')</tbody>
    </table>
  </section>

  <section class="block">
    <h2>このレポートの読み方と注意点</h2>
    <div class="finding">
      <p class="detail">
        検出された項目は「必ず直すべきもの」ではありません。<strong>意図してその設定になっている場合もあります。</strong>
        とくに「消していいもの」と「消すと壊れるもの」の判別には、その環境が作られた経緯を知る必要があります。
        判断に迷う項目があれば、実際の構成や運用と突き合わせたうえでご相談ください。
      </p>
      <p class="detail" style="margin-bottom:0">
        スコアは環境間の優劣を示すものではなく、<strong>同じ環境の改善前後を比べるための目安</strong>としてご利用ください。
      </p>
    </div>
  </section>
</div>

<footer>
  <div class="wrap">
    <strong>Active Directory 健全性診断レポート</strong>　生成: $generated<br>
    読み取り専用の検査により作成。検出内容の解釈と対処の判断については、環境の経緯を踏まえた確認をお勧めします。
  </div>
</footer>
</body>
</html>
"@)

    return $sb.ToString()
}

#endregion

#region ---------- サンプルデータ ----------

<#
    AD へ接続せずにレポートの体裁を確認するための架空データ。
    実在の組織とは関係ありません。
#>
function Get-DemoSnapshot {
    param([datetime]$Now)

    $mkUser = {
        param($name,$daysLogon,$daysPwd,$enabled,$admin,$spn,$flags)
        [pscustomobject]@{
            SamAccountName = $name
            Enabled = $enabled
            LastLogonTimestamp = if ($null -ne $daysLogon) { $Now.AddDays(-1*$daysLogon).ToFileTime() } else { $null }
            PasswordLastSet = $Now.AddDays(-1*$daysPwd)
            PasswordNeverExpires = $flags.Contains('neverexp')
            PasswordNotRequired  = $flags.Contains('nopwd')
            DoesNotRequirePreAuth = $flags.Contains('nopreauth')
            AllowReversiblePasswordEncryption = $flags.Contains('reversible')
            TrustedForDelegation = $flags.Contains('deleg')
            AccountNotDelegated = $false
            adminCount = $(if ($admin) { 1 } else { $null })
            ServicePrincipalName = $spn
            whenCreated = $Now.AddDays(-1*($daysPwd+400))
            SID = [pscustomobject]@{ Value = 'S-1-5-21-1111111111-2222222222-3333333333-' + (Get-Random -Minimum 1100 -Maximum 9999) }
            MemberOf = @()
        }
    }

    $users = @()
    $users += & $mkUser 'krbtgt' $null 412 $false $false @() @()
    $admin = & $mkUser 'Administrator' 12 730 $true $true @() @()
    $admin.SID = [pscustomobject]@{ Value = 'S-1-5-21-1111111111-2222222222-3333333333-500' }
    $users += $admin
    $users += & $mkUser 'svc-sql01'     3   980 $true $true  @('MSSQLSvc/sql01.contoso.local:1433') @('neverexp')
    $users += & $mkUser 'svc-backup'    1   1120 $true $false @('BackupSvc/bk01.contoso.local')      @('neverexp')
    $users += & $mkUser 'svc-scan'      40  1450 $true $false @('HTTP/scan.contoso.local')           @('neverexp','nopreauth')
    $users += & $mkUser 'legacy-app'    260 1600 $true $false @()                                    @('reversible','nopwd')
    $users += & $mkUser 'taro.yamada'   1   45  $true $false @() @()
    $users += & $mkUser 'hanako.suzuki' 2   62  $true $false @() @()
    $users += & $mkUser 'ex.tanaka'     420 900 $true $false @() @()
    $users += & $mkUser 'ex.watanabe'   380 870 $true $false @() @()
    $users += & $mkUser 'ex.kobayashi'  310 810 $true $false @() @()
    1..14 | ForEach-Object { $users += & $mkUser ("user{0:D3}" -f $_) (Get-Random -Minimum 1 -Maximum 20) 60 $true $false @() @() }

    $mkComp = {
        param($name,$os,$daysLogon,$enabled,$deleg)
        [pscustomobject]@{
            Name = $name
            Enabled = $enabled
            OperatingSystem = $os
            OperatingSystemVersion = ''
            LastLogonTimestamp = $Now.AddDays(-1*$daysLogon).ToFileTime()
            TrustedForDelegation = $deleg
            whenCreated = $Now.AddDays(-1200)
            PasswordLastSet = $Now.AddDays(-1*$daysLogon)
        }
    }
    $computers = @()
    $computers += & $mkComp 'DC01' 'Windows Server 2019 Standard' 0 $true $true
    $computers += & $mkComp 'DC02' 'Windows Server 2019 Standard' 0 $true $true
    $computers += & $mkComp 'FS01' 'Windows Server 2012 R2 Standard' 2 $true $false
    $computers += & $mkComp 'APP01' 'Windows Server 2016 Standard' 1 $true $true
    $computers += & $mkComp 'PRINT01' 'Windows Server 2008 R2 Standard' 5 $true $false
    1..22 | ForEach-Object { $computers += & $mkComp ("PC-{0:D3}" -f $_) 'Windows 10 Pro' (Get-Random -Minimum 0 -Maximum 8) $true $false }
    1..6  | ForEach-Object { $computers += & $mkComp ("OLD-PC-{0:D2}" -f $_) 'Windows 7 Professional' (Get-Random -Minimum 200 -Maximum 700) $true $false }

    $dcs = @(
        [pscustomobject]@{ Name='DC01'; HostName='dc01.contoso.local'; Site='Tokyo' }
        [pscustomobject]@{ Name='DC02'; HostName='dc02.contoso.local'; Site='Tokyo' }
    )

    $privGroups = @(
        [pscustomobject]@{ Name='Domain Admins';     Members=@(1..9)  }
        [pscustomobject]@{ Name='Enterprise Admins'; Members=@(1..3)  }
        [pscustomobject]@{ Name='Schema Admins';     Members=@(1..1)  }
        [pscustomobject]@{ Name='Administrators';    Members=@(1..11) }
        [pscustomobject]@{ Name='Backup Operators';  Members=@(1..2)  }
    )

    [pscustomobject]@{
        IsDemo            = $true
        Now               = $Now
        DomainName        = 'contoso.local'
        NetBIOSName       = 'CONTOSO'
        DomainMode        = 'Windows2016Domain'
        ForestMode        = 'Windows2016Forest'
        Users             = $users
        Computers         = $computers
        DomainControllers = $dcs
        PrivilegedGroups  = $privGroups
        PasswordPolicy    = [pscustomobject]@{ MinPasswordLength=8; LockoutThreshold=0; ComplexityEnabled=$true }
        Sites             = @(
            [pscustomobject]@{ Name='Tokyo' }
            [pscustomobject]@{ Name='Osaka' }
            [pscustomobject]@{ Name='OldBranch' }
        )
        Subnets           = @(
            [pscustomobject]@{ Name='10.0.0.0/24'; siteObject='CN=Tokyo,...' }
            [pscustomobject]@{ Name='10.1.0.0/24'; siteObject=$null }
            [pscustomobject]@{ Name='192.168.50.0/24'; siteObject=$null }
        )
        RecycleBinEnabled = $false
        Krbtgt            = $users | Where-Object SamAccountName -eq 'krbtgt' | Select-Object -First 1
        BuiltinAdmin      = $admin
        LapsSchema        = $false
    }
}

#endregion

#region ---------- メイン ----------

$now = Get-Date

Write-Host ''
Write-Host 'Active Directory 健全性診断' -ForegroundColor White
Write-Host '----------------------------------------' -ForegroundColor DarkGray
Write-Host '読み取り専用で実行します（構成を変更する処理は含まれません）' -ForegroundColor DarkGray
Write-Host ''

if ($DemoData) {
    Write-Warn 'サンプルモードです。Active Directory へは接続しません。'
    $snapshot = Get-DemoSnapshot -Now $now
} else {
    $snapshot = Get-AdSnapshot -Now $now
}

if (-not $CompanyName) {
    $CompanyName = $snapshot.DomainName
}

Write-Step '検査を実行しています...'
$findings = @()
foreach ($check in $script:CheckRegistry) {
    try {
        $result = & $check -Snap $snapshot
        if ($result) { $findings += @($result) }
    } catch {
        Write-Warn ("{0} の実行中にエラーが発生したためスキップしました: {1}" -f $check, $_.Exception.Message)
    }
}

$score = Get-HealthScore -Findings $findings
$rank  = Get-ScoreRank -Score $score

Write-Host ''
Write-Host ('  スコア: {0} 点（評価 {1} — {2}）' -f $score, $rank.Rank, $rank.Text) -ForegroundColor White
Write-Host ('  検出: {0} 件（High {1} / Medium {2} / Low {3}）' -f `
    $findings.Count,
    @($findings | Where-Object Severity -eq 'High').Count,
    @($findings | Where-Object Severity -eq 'Medium').Count,
    @($findings | Where-Object Severity -eq 'Low').Count) -ForegroundColor White
foreach ($pkey in $script:PhaseMeta.Keys) {
    $n = @($findings | Where-Object Phase -eq $pkey).Count
    Write-Host ('    {0} {1}: {2} 件' -f $script:PhaseMeta[$pkey].Label, $script:PhaseMeta[$pkey].Title, $n) -ForegroundColor Gray
}
Write-Host ''

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}
$stamp = $now.ToString('yyyyMMdd-HHmmss')
$htmlPath = Join-Path $OutputPath ("AD-HealthReport-{0}.html" -f $stamp)
$csvPath  = Join-Path $OutputPath ("AD-HealthReport-{0}.csv"  -f $stamp)

Write-Step 'HTML レポートを生成しています...'
$html = New-HtmlReport -Snap $snapshot -Findings $findings -Company $CompanyName
$html | Out-File -LiteralPath $htmlPath -Encoding utf8

# 明細 CSV（全件）
$csvRows = foreach ($f in $findings) {
    if ($f.Items.Count -gt 0) {
        foreach ($it in $f.Items) {
            $flat = ($it.PSObject.Properties | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join '; '
            [pscustomobject]@{ Phase=$f.Phase; Severity=$f.Severity; Category=$f.Category; Title=$f.Title; Item=$flat; Impact=$f.Impact; Action=$f.Action; Reference=$f.Reference }
        }
    } else {
        [pscustomobject]@{ Phase=$f.Phase; Severity=$f.Severity; Category=$f.Category; Title=$f.Title; Item=''; Impact=$f.Impact; Action=$f.Action; Reference=$f.Reference }
    }
}
$csvRows | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8

Write-Host ''
Write-Host ('  HTML: {0}' -f (Resolve-Path -LiteralPath $htmlPath)) -ForegroundColor Green
Write-Host ('  CSV : {0}' -f (Resolve-Path -LiteralPath $csvPath))  -ForegroundColor Green
Write-Host ''

# 結果をパイプラインにも返す（$result = .\Invoke-ADHealthReport.ps1 のように使えます）
$findings

#endregion
