function Invoke-ADStaleInventory {
    [CmdletBinding()]
    param(
        [int]$InactiveDays = 90,
        [string]$OutputPath,
        [string]$DnsServer,
        [switch]$CheckSites = $true  # ★ 追加：サイト/サブネット/サイトリンクの健全性チェック
    )

    # ===== 共通：準備 =====
    $now = Get-Date
    $inactiveThresholdDate = $now.AddDays(-1 * $InactiveDays)

    $rootDse = Get-ADRootDSE
    $configNC = $rootDse.configurationNamingContext
    $domainNC = $rootDse.defaultNamingContext

    $results = New-Object System.Collections.Generic.List[object]

    # ===== 1. DCメタデータ（前回と同等） =====
    $liveDCs = Get-ADDomainController -Filter * -ErrorAction SilentlyContinue |
        Select-Object HostName,Name,IPv4Address,Site,ComputerObjectDN

    $ntdsObjects = Get-ADObject -SearchBase $configNC `
        -LDAPFilter '(objectClass=nTDSDSA)' -Properties objectClass,distinguishedName,objectGUID,invocationId,msDS-Behavior-Version `
        -ErrorAction SilentlyContinue

    $dcMeta = foreach ($ntds in $ntdsObjects) {
        $serverDN = ($ntds.DistinguishedName -split '(?i),CN=NTDS Settings,')[0]
        $serverObj = Get-ADObject -Identity $serverDN -Properties dNSHostName,siteObject -ErrorAction SilentlyContinue
        if (-not $serverObj) { continue }
        $comp = $null
        if ($serverObj.dNSHostName) {
            $comp = Get-ADComputer -Filter "dnsHostName -eq '$($serverObj.dNSHostName)'" -Properties lastLogonTimestamp,Enabled,whenCreated `
                -ErrorAction SilentlyContinue
        }
        [pscustomobject]@{
            ServerObjectDN        = $serverObj.DistinguishedName
            NTDSDN                = $ntds.DistinguishedName
            DNSHostName           = $serverObj.dNSHostName
            ADSiteObject          = $serverObj.siteObject
            ComputerDN            = $comp.DistinguishedName
            ComputerEnabled       = $comp.Enabled
            ComputerWhenCreated   = $comp.whenCreated
            ComputerLastLogonTS   = $comp.lastLogonTimestamp
        }
    }

    $liveNames = $liveDCs.HostName | ForEach-Object { $_.ToLower() }

    foreach ($row in $dcMeta) {
        $isReachable = $false
        if ($row.DNSHostName) {
            $isReachable = $liveNames -contains $row.DNSHostName.ToLower()
        }
        if (-not $isReachable) {
            $lastSeen = $null
            if ($row.ComputerLastLogonTS) { $lastSeen = [DateTime]::FromFileTime([int64]$row.ComputerLastLogonTS) }
            $results.Add([pscustomobject]@{
                Category           = 'DC-Metadata'
                Name               = $row.DNSHostName
                DN                 = $row.ServerObjectDN
                Detail             = 'NTDS/Serverは存在するが稼働DCとして見えない。降格後メタデータ残骸の可能性。'
                LastSeen           = $lastSeen
                Confidence         = 'High'
                RecommendedAction  = 'ntdsutil等でメタデータクリーンアップ検討 (要慎重)'
            })
        } elseif (-not $row.ComputerDN) {
            $results.Add([pscustomobject]@{
                Category           = 'DC-Metadata'
                Name               = $row.DNSHostName
                DN                 = $row.ServerObjectDN
                Detail             = 'DCは稼働中だが対応Computerオブジェクトが見つからない。'
                LastSeen           = $null
                Confidence         = 'Medium'
                RecommendedAction  = 'Computerオブジェクトの存在/OUを確認'
            })
        }
    }

    # ===== 2. 使われていないコンピュータ =====
    $computers = Get-ADComputer -Filter * -Properties lastLogonTimestamp,whenCreated,Enabled,OperatingSystem,PasswordLastSet -ErrorAction SilentlyContinue
    foreach ($c in $computers) {
        $llts = $null
        if ($c.lastLogonTimestamp) { $llts = [DateTime]::FromFileTime([int64]$c.lastLogonTimestamp) }
        $isDisabled = ($c.Enabled -eq $false)
        $isStaleByLogon = ($llts -and $llts -lt $inactiveThresholdDate)
        $isStaleByCreateOnly = (-not $llts -and $c.whenCreated -lt $inactiveThresholdDate.AddDays(-$InactiveDays))
        if ($isDisabled -or $isStaleByLogon -or $isStaleByCreateOnly) {
            $detailBits = @()
            if ($isDisabled) { $detailBits += "アカウントが無効(Disabled)" }
            if ($isStaleByLogon) { $detailBits += "最終ログオンが $InactiveDays 日より古い ($llts)" }
            if ($isStaleByCreateOnly) { $detailBits += "作成から長期未使用 (lastLogonTimestampなし)" }
            $confidence = if ($isDisabled) { 'High' } elseif ($isStaleByLogon) { 'Medium' } else { 'Low' }
            $results.Add([pscustomobject]@{
                Category='Stale-Computer'; Name=$c.Name; DN=$c.DistinguishedName;
                Detail = ($detailBits -join '; '); LastSeen=$llts; Confidence=$confidence;
                RecommendedAction='担当部署に確認後、無効化/削除やOU隔離を検討'
            })
        }
    }

    # ===== 3. 使われていないユーザー =====
    $users = Get-ADUser -Filter * -Properties lastLogonTimestamp,Enabled,whenCreated,PasswordNeverExpires,PasswordLastSet -ErrorAction SilentlyContinue
    foreach ($u in $users) {
        $ullts = $null
        if ($u.lastLogonTimestamp) { $ullts = [DateTime]::FromFileTime([int64]$u.lastLogonTimestamp) }
        $isDisabledUser = ($u.Enabled -eq $false)
        $isInactiveUser = ($ullts -and $ullts -lt $inactiveThresholdDate)
        $isOldNoLogon   = (-not $ullts -and $u.whenCreated -lt $inactiveThresholdDate.AddDays(-$InactiveDays))
        if ($isDisabledUser -or $isInactiveUser -or $isOldNoLogon) {
            $detailBits = @()
            if ($isDisabledUser) { $detailBits += "アカウントが無効(Disabled)" }
            if ($isInactiveUser) { $detailBits += "最終ログオンが $InactiveDays 日より古い ($ullts)" }
            if ($isOldNoLogon)   { $detailBits += "作成後ほぼ未使用(lastLogonTimestampなし)" }
            if ($u.PasswordNeverExpires) { $detailBits += "PasswordNeverExpires=True(サービス/共有アカウントの可能性)" }
            $confidence = if ($isDisabledUser) { 'High' } elseif ($isInactiveUser) { 'Medium' } else { 'Low' }
            $results.Add([pscustomobject]@{
                Category='Stale-User'; Name=$u.SamAccountName; DN=$u.DistinguishedName;
                Detail = ($detailBits -join '; '); LastSeen=$ullts; Confidence=$confidence;
                RecommendedAction='人事・主管部署と突合。無効化済みなら削除候補/アーカイブOUへ'
            })
        }
    }

    # ===== 4. (任意) 降格DC名のDNS残骸 =====
    if ($DnsServer) {
        $suspectDCs = $results | Where-Object { $_.Category -eq 'DC-Metadata' -and $_.Confidence -eq 'High' -and $_.Name }
        foreach ($dc in $suspectDCs) {
            $hostFqdn = $dc.Name
            if (-not $hostFqdn) { continue }
            $dnsRecords = @()
            try {
                $dnsRecords += Get-DnsServerResourceRecord -ComputerName $DnsServer -ZoneName ($hostFqdn.Split('.',2)[1]) -Name ($hostFqdn.Split('.',2)[0]) -ErrorAction SilentlyContinue
            } catch {}
            $forestRoot = (Get-ADForest).RootDomain
            $msdcsZone = "_msdcs.$forestRoot"
            try {
                $dnsRecords += Get-DnsServerResourceRecord -ComputerName $DnsServer -ZoneName $msdcsZone -ErrorAction SilentlyContinue |
                    Where-Object { $_.RecordData -match [regex]::Escape($hostFqdn) }
            } catch {}
            if ($dnsRecords -and $dnsRecords.Count -gt 0) {
                $results.Add([pscustomobject]@{
                    Category='Stale-DC-DNS'; Name=$hostFqdn; DN="DNS:$DnsServer";
                    Detail="降格済み疑いのDC [$hostFqdn] のレコードがDNSに残っている可能性";
                    LastSeen=$null; Confidence='High';
                    RecommendedAction='DNSのA/SRV/CNAMEを手動で確認・掃除'
                })
            }
        }
    }

    # ===== 5. ★ サイト/サブネット/サイトリンクの健全性チェック =====
    if ($CheckSites) {

        # 5-0. 収集
        $sitesBase = "CN=Sites,$configNC"
        $allSites = Get-ADObject -SearchBase $sitesBase -LDAPFilter '(objectClass=site)' -Properties distinguishedName,name -ErrorAction SilentlyContinue

        $siteMap = @{}
        foreach ($s in $allSites) { $siteMap[$s.DistinguishedName] = $s }

        $allSubnets = Get-ADObject -SearchBase "CN=Subnets,$sitesBase" -LDAPFilter '(objectClass=subnet)' -Properties name,siteObject,distinguishedName -ErrorAction SilentlyContinue

        $siteLinks = @()
        foreach ($t in @("IP","SMTP")) {
            $path = "CN=$t,CN=Inter-Site Transports,$sitesBase"
            $siteLinks += Get-ADObject -SearchBase $path -LDAPFilter '(objectClass=siteLink)' -Properties siteList,name,distinguishedName -ErrorAction SilentlyContinue
        }

        # 5-1. 各サイトのサーバー数・サブネット数・リンク参照数
        $siteReferencedSet = [System.Collections.Generic.HashSet[string]]::new()
        foreach ($lnk in $siteLinks) {
            if ($lnk.siteList) { foreach ($dn in @($lnk.siteList)) { [void]$siteReferencedSet.Add($dn) } }
        }

        foreach ($site in $allSites) {
            $servers = @()
            try {
                $servers = Get-ADObject -SearchBase ("CN=Servers," + $site.DistinguishedName) -LDAPFilter '(objectClass=server)' -ErrorAction SilentlyContinue
            } catch {}
            $serverCount = ($servers | Measure-Object).Count

            $subnetCount = ($allSubnets | Where-Object { $_.siteObject -eq $site.DistinguishedName }).Count
            $linkCount = if ($siteReferencedSet.Contains($site.DistinguishedName)) { 1 } else { 0 }

            # パターン判定
            if ($serverCount -eq 0 -and $subnetCount -eq 0 -and $linkCount -eq 0) {
                $results.Add([pscustomobject]@{
                    Category='AD-Site'; Name=$site.Name; DN=$site.DistinguishedName;
                    Detail='サーバー/サブネット/サイトリンクいずれからも参照されない空サイト';
                    LastSeen=$null; Confidence='High';
                    RecommendedAction='依存を再確認のうえサイト削除候補'
                })
            }
            elseif ($serverCount -eq 0 -and $subnetCount -gt 0) {
                $results.Add([pscustomobject]@{
                    Category='AD-Site'; Name=$site.Name; DN=$site.DistinguishedName;
                    Detail="サブネットは割当済みだがサーバー無し（支店サイトでDCなし運用なら正常の可能性）";
                    LastSeen=$null; Confidence='Medium';
                    RecommendedAction='意図通りなら維持、不要ならサブネット統合/サイト削除を検討'
                })
            }
            elseif ($serverCount -gt 0 -and $subnetCount -eq 0) {
                $results.Add([pscustomobject]@{
                    Category='AD-Site'; Name=$site.Name; DN=$site.DistinguishedName;
                    Detail="サーバー有りだがサブネット未割当（クライアントのサイト判定が不安定）";
                    LastSeen=$null; Confidence='Medium';
                    RecommendedAction='該当拠点のサブネットを作成しサイトに割当'
                })
            }

            if ($serverCount -gt 0 -and $linkCount -eq 0 -and $allSites.Count -gt 1) {
                $results.Add([pscustomobject]@{
                    Category='AD-Site'; Name=$site.Name; DN=$site.DistinguishedName;
                    Detail="複数サイト構成なのにサイトリンクに未参加（レプリケーション経路が無い可能性）";
                    LastSeen=$null; Confidence='High';
                    RecommendedAction='siteLinkへ当該サイトを追加（既存リンクに追加 or 新規作成）'
                })
            }
        }

        # 5-2. サブネットの不整合・重複/オーバーラップ
        function ConvertTo-IPv4Int([string]$ip) {
            $p = $ip.Split('.'); if ($p.Count -ne 4) { return $null }
            try { [uint32]($p[0] -shl 24 -bor ($p[1] -shl 16) -bor ($p[2] -shl 8) -bor $p[3]) } catch { return $null }
        }
        function Parse-CIDR([string]$cidr) {
            $parts = $cidr -split '/'
            if ($parts.Count -ne 2) { return $null }
            $ipInt = ConvertTo-IPv4Int $parts[0]; if ($null -eq $ipInt) { return $null }
            $prefix = [int]$parts[1]; if ($prefix -lt 0 -or $prefix -gt 32) { return $null }
            $mask = if ($prefix -eq 0) { [uint32]0 } else { [uint32]0xFFFFFFFF -shl (32 - $prefix) }
            $network = $ipInt -band $mask
            $lower = $network
            $upper = [uint32]($network -bor ([uint32]0xFFFFFFFF -bxor $mask))
            [pscustomobject]@{ CIDR=$cidr; Prefix=$prefix; Mask=$mask; Net=$network; Lower=$lower; Upper=$upper }
        }

        # サブネット整合性
        foreach ($sn in $allSubnets) {
            $snSite = $sn.siteObject
            if (-not $snSite) {
                $results.Add([pscustomobject]@{
                    Category='AD-Subnet'; Name=$sn.Name; DN=$sn.DistinguishedName;
                    Detail='どのサイトにも割り当てられていないサブネット';
                    LastSeen=$null; Confidence='High';
                    RecommendedAction='適切なサイトに割当 or 不要なら削除'
                })
                continue
            }
            if (-not $siteMap.ContainsKey($snSite)) {
                $results.Add([pscustomobject]@{
                    Category='AD-Subnet'; Name=$sn.Name; DN=$sn.DistinguishedName;
                    Detail='存在しないサイトを参照しているサブネット';
                    LastSeen=$null; Confidence='High';
                    RecommendedAction='正しいサイトへ割当し直す or サブネット削除'
                })
            }
        }

        # 重複/オーバーラップ検出（IPv4のみ）
        $rangeItems = @()
        foreach ($sn in $allSubnets) {
            $cidr = $sn.Name
            $r = Parse-CIDR $cidr
            if ($null -ne $r) {
                $rangeItems += [pscustomobject]@{
                    SubnetDN=$sn.DistinguishedName; SiteDN=$sn.siteObject; CIDR=$r.CIDR;
                    Lower=$r.Lower; Upper=$r.Upper; Prefix=$r.Prefix
                }
            } else {
                # フォーマット不正/IPv6などは軽めに注意
                $results.Add([pscustomobject]@{
                    Category='AD-Subnet'; Name=$sn.Name; DN=$sn.DistinguishedName;
                    Detail='CIDR表記の解析に失敗（IPv6または表記ゆれの可能性）';
                    LastSeen=$null; Confidence='Low';
                    RecommendedAction='表記の統一 or IPv6は別途レビュー'
                })
            }
        }

        # O(n^2)簡易チェック（数百個規模なら許容）
        for ($i=0; $i -lt $rangeItems.Count; $i++) {
            for ($j=$i+1; $j -lt $rangeItems.Count; $j++) {
                $a = $rangeItems[$i]; $b = $rangeItems[$j]
                # 完全重複
                if ($a.CIDR -eq $b.CIDR) {
                    $results.Add([pscustomobject]@{
                        Category='AD-Subnet';
                        Name="$($a.CIDR) (duplicate)";
                        DN="$($a.SubnetDN) | $($b.SubnetDN)";
                        Detail='同一CIDRのサブネットが複数存在';
                        LastSeen=$null; Confidence='High';
                        RecommendedAction='どちらか片方を削除（割当サイトを確認のうえ統合）'
                    })
                    continue
                }
                # オーバーラップ（範囲が交差）
                $overlap = -not( ($a.Upper -lt $b.Lower) -or ($b.Upper -lt $a.Lower) )
                if ($overlap) {
                    $results.Add([pscustomobject]@{
                        Category='AD-Subnet';
                        Name="$($a.CIDR) ? $($b.CIDR)";
                        DN="$($a.SubnetDN) | $($b.SubnetDN)";
                        Detail='サブネット範囲がオーバーラップ（サイト判定が不安定になります）';
                        LastSeen=$null; Confidence='High';
                        RecommendedAction='CIDRを見直し、重複/包含関係を解消'
                    })
                }
            }
        }

        # 5-3. サイトリンクの健全性
        foreach ($lnk in $siteLinks) {
            $sitesInLink = @()
            if ($lnk.siteList) { $sitesInLink = @($lnk.siteList) }
            $count = $sitesInLink.Count
            if ($count -lt 2) {
                $results.Add([pscustomobject]@{
                    Category='AD-SiteLink'; Name=$lnk.Name; DN=$lnk.DistinguishedName;
                    Detail='リンクに含まれるサイトが2つ未満（レプリケーション経路として無効）';
                    LastSeen=$null; Confidence='High';
                    RecommendedAction='対象サイトを2つ以上設定 or 不要ならリンク削除'
                })
            }
            foreach ($sdn in $sitesInLink) {
                if (-not $siteMap.ContainsKey($sdn)) {
                    $results.Add([pscustomobject]@{
                        Category='AD-SiteLink'; Name=$lnk.Name; DN=$lnk.DistinguishedName;
                        Detail="存在しないサイト [$sdn] を参照しているサイトリンク";
                        LastSeen=$null; Confidence='High';
                        RecommendedAction='siteList から不正参照を除去'
                    })
                }
            }
        }
    }

    # ===== 出力整形・保存 =====
    $order = @{ 'High' = 0; 'Medium' = 1; 'Low' = 2 }
    $final = $results | Sort-Object {
        if ($order.ContainsKey($_.Confidence)) { $order[$_.Confidence] } else { 9 }
    }, Category, Name

    if ($OutputPath) {
        if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath | Out-Null }
        $timestamp = $now.ToString('yyyyMMdd-HHmmss')
        $final | Export-Csv (Join-Path $OutputPath "AD-StaleInventory-$timestamp.csv") -NoTypeInformation -Encoding UTF8

        $byCats = @('DC-Metadata','Stale-Computer','Stale-User','Stale-DC-DNS','AD-Site','AD-Subnet','AD-SiteLink')
        foreach ($cat in $byCats) {
            $final | Where-Object { $_.Category -eq $cat } |
                Export-Csv (Join-Path $OutputPath "AD-StaleInventory-$cat-$timestamp.csv") -NoTypeInformation -Encoding UTF8
        }
        Write-Host "[INFO] CSVを書き出しました: $OutputPath" -ForegroundColor Green
    }

    return $final
}

<# 実行例
# サイト健全性チェック込み（既定で有効）
Invoke-ADStaleInventory -OutputPath "C:\Temp\AD-StaleReport"

# サブネット/サイトリンク含むがDNSは不要
Invoke-ADStaleInventory -CheckSites -InactiveDays 180

# 降格DCのDNS残骸まで見る
Invoke-ADStaleInventory -DnsServer "dc01.contoso.local" -OutputPath "C:\Temp\AD-StaleReport"
#>
