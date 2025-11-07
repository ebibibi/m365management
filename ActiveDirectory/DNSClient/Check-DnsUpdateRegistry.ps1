#Requires -Version 3.0
<# 
  Check-DnsUpdateRegistry.ps1
  DNS 動的更新に関わる主要レジストリ値を収集してレポートを出力します。
  - 実行：スイッチ不要（カレント配下に出力ディレクトリを作成）
  - 収集対象：
      [Global] HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters
        - DisableDynamicUpdate (DWORD, default 0)
        - DisableReverseAddressRegistrations (DWORD, default 0)
        - DefaultRegistrationRefreshInterval (DWORD, default 86400)
        - DefaultRegistrationTTL (DWORD, default 1200)
        - UpdateSecurityLevel (DWORD, default 0)  ; 0 / 0x10 / 0x100
        - DisableReplaceAddressesInConflicts (DWORD, default 0)

      [Per-Interface] HKLM\...\Parameters\Interfaces\<GUID>\DisableDynamicUpdate (DWORD, default 0)

      [Per-Adapter] HKLM\...\Parameters\Adapters\<InterfaceName>\MaxNumberOfAddressesToRegister (DWORD, default 1)

  注意：
    - 値が存在しない場合は「未設定（= 既定値が有効）」として扱います。
    - 実効判定（DNS更新 有効/無効）は
        Global DisableDynamicUpdate == 0 かつ Per-Interface DisableDynamicUpdate == 0 → 有効
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# --- 出力先ディレクトリ作成 ---
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$OutDir = Join-Path -Path (Get-Location) -ChildPath "DNSUpdateRegistryReport_$stamp"
New-Item -Path $OutDir -ItemType Directory -Force | Out-Null

# --- ヘルパー：レジストリ値取得（存在しない→$null） ---
function Get-RegValue {
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [string] $Name
    )
    try {
        if (Test-Path -LiteralPath $Path) {
            $item = Get-ItemProperty -LiteralPath $Path -ErrorAction Stop
            if ($null -ne ($item.PSObject.Properties[$Name])) {
                return $item.$Name
            }
        }
        return $null
    }
    catch {
        return $null
    }
}

# --- ヘルパー：意味付け（簡易） ---
function Meaning-DisableDynamicUpdate([Nullable[int]]$value) {
    if ($value -eq 1) { 'DNS更新の登録を無効' }
    elseif ($value -eq 0) { 'DNS更新の登録を有効' }
    else { '未設定（既定=有効）' }
}
function Meaning-DisableReverse([Nullable[int]]$value) {
    if ($value -eq 1) { 'PTR登録しない' }
    elseif ($value -eq 0) { 'PTR登録する' }
    else { '未設定（既定=PTR登録する）' }
}
function Meaning-UpdateSecurityLevel([Nullable[int]]$value) {
    switch ($value) {
        0 { '未保護で拒否された場合のみ保護付きを送信（既定）' }
        16 { '未保護のみ送信 (0x10)' }
        256 { '保護付きのみ送信 (0x100)' }
        default { '未設定（既定=0）/不明' }
    }
}
function Meaning-DisableReplace([Nullable[int]]$value) {
    if ($value -eq 1) { '競合時は上書きせずバックアウト（ログにもエラー出さない）' }
    elseif ($value -eq 0) { '競合時は既存Aを自IPのAで上書き（既定）' }
    else { '未設定（既定=上書き）' }
}

# --- 既定値辞書 ---
$Defaults = @{
    DisableDynamicUpdate                  = 0
    DisableReverseAddressRegistrations    = 0
    DefaultRegistrationRefreshInterval    = 86400
    DefaultRegistrationTTL                = 1200
    UpdateSecurityLevel                   = 0
    DisableReplaceAddressesInConflicts    = 0
    PerInterface_DisableDynamicUpdate     = 0
    MaxNumberOfAddressesToRegister        = 1
}

$BasePath = 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters'

# --- Global 値収集 ---
$globalRows = @()

$globalKeys = @(
    @{ Name='DisableDynamicUpdate'; Path=$BasePath; Default=$Defaults.DisableDynamicUpdate; MeaningFunc='Meaning-DisableDynamicUpdate' },
    @{ Name='DisableReverseAddressRegistrations'; Path=$BasePath; Default=$Defaults.DisableReverseAddressRegistrations; MeaningFunc='Meaning-DisableReverse' },
    @{ Name='DefaultRegistrationRefreshInterval'; Path=$BasePath; Default=$Defaults.DefaultRegistrationRefreshInterval; MeaningFunc=$null },
    @{ Name='DefaultRegistrationTTL'; Path=$BasePath; Default=$Defaults.DefaultRegistrationTTL; MeaningFunc=$null },
    @{ Name='UpdateSecurityLevel'; Path=$BasePath; Default=$Defaults.UpdateSecurityLevel; MeaningFunc='Meaning-UpdateSecurityLevel' },
    @{ Name='DisableReplaceAddressesInConflicts'; Path=$BasePath; Default=$Defaults.DisableReplaceAddressesInConflicts; MeaningFunc='Meaning-DisableReplace' }
)

foreach ($g in $globalKeys) {
    $val = Get-RegValue -Path $g.Path -Name $g.Name
    $present = ($null -ne $val)
    $effective = if ($present) { [int64]$val } else { [int64]$g.Default }
    $meaning = if ($g.MeaningFunc) { & $g.MeaningFunc $val } else { '' }

    $globalRows += [pscustomobject]@{
        SettingName = $g.Name
        RegistryPath = $g.Path
        Present      = if($present){'Yes'}else{'No'}
        Value        = if($present){$val}else{'(not set)'}
        Default      = $g.Default
        Effective    = $effective
        Meaning      = $meaning
    }
}

$globalCsv = Join-Path $OutDir 'global.csv'
$globalRows | Export-Csv -NoTypeInformation -Encoding UTF8 $globalCsv

# --- インターフェイス別 DisableDynamicUpdate 収集 ---
$ifBase = Join-Path $BasePath 'Interfaces'
$ifRows = @()

# ネットワークアダプタの GUID → フレンドリ名解決
# Get-NetAdapter が使えない環境もあるため両対応
$guidToName = @{}
try {
    $adapters = Get-NetAdapter -ErrorAction Stop | Where-Object { $_.InterfaceGuid -ne $null }
    foreach ($a in $adapters) {
        $guidToName[$a.InterfaceGuid.Guid.ToUpper()] = $a.Name
    }
}
catch {
    # WMI フォールバック
    $nics = Get-WmiObject Win32_NetworkAdapter -ErrorAction SilentlyContinue | Where-Object { $_.GUID }
    foreach ($n in $nics) {
        $guidToName[$n.GUID.ToUpper()] = $n.NetConnectionID
    }
}

$globalDisable = ($globalRows | Where-Object SettingName -eq 'DisableDynamicUpdate').Effective

if (Test-Path -LiteralPath $ifBase) {
    Get-ChildItem -LiteralPath $ifBase | ForEach-Object {
        $guid = $_.PSChildName.ToUpper()
        $ifPath = $_.PSPath
        $ifName = if ($guidToName.ContainsKey($guid)) { $guidToName[$guid] } else { $guid }

        $ifVal = Get-RegValue -Path $ifPath -Name 'DisableDynamicUpdate'
        $ifPresent = ($null -ne $ifVal)
        $ifEffective = if ($ifPresent) { [int64]$ifVal } else { [int64]$Defaults.PerInterface_DisableDynamicUpdate }

        # 実効判定：両方0なら「有効」
        $dnsUpdateEnabled = if ( ($globalDisable -eq 0) -and ($ifEffective -eq 0) ) { 'Enabled' } else { 'Disabled' }

        $ifRows += [pscustomobject]@{
            InterfaceGuid  = $guid
            InterfaceName  = $ifName
            RegistryPath   = $ifPath
            Present        = if($ifPresent){'Yes'}else{'No'}
            Value          = if($ifPresent){$ifVal}else{'(not set)'}
            Effective      = $ifEffective
            Global_DisableDynamicUpdate = $globalDisable
            DNS_Update_Effective        = $dnsUpdateEnabled
            Meaning       = Meaning-DisableDynamicUpdate $ifVal
        }
    }
}

$ifCsv = Join-Path $OutDir 'interfaces.csv'
$ifRows | Sort-Object InterfaceName | Export-Csv -NoTypeInformation -Encoding UTF8 $ifCsv

# --- アダプター別 MaxNumberOfAddressesToRegister 収集 ---
$adaptersBase = Join-Path $BasePath 'Adapters'
$adRows = @()
if (Test-Path -LiteralPath $adaptersBase) {
    Get-ChildItem -LiteralPath $adaptersBase | ForEach-Object {
        $nameKey = $_.PSChildName
        $path = $_.PSPath
        $val = Get-RegValue -Path $path -Name 'MaxNumberOfAddressesToRegister'
        $present = ($null -ne $val)
        $effective = if ($present) { [int64]$val } else { [int64]$Defaults.MaxNumberOfAddressesToRegister }

        $note = if ($effective -eq 0) { 'このアダプターはIPをDNS登録しない（Max=0）' }
                elseif ($effective -eq 1) { '既定：最初の1つのIPのみ登録' }
                else { '複数IPを最大 値 分まで登録' }

        $adRows += [pscustomobject]@{
            AdapterKeyName = $nameKey
            RegistryPath   = $path
            Present        = if($present){'Yes'}else{'No'}
            Value          = if($present){$val}else{'(not set)'}
            Default        = $Defaults.MaxNumberOfAddressesToRegister
            Effective      = $effective
            Note           = $note
        }
    }
}

$adCsv = Join-Path $OutDir 'adapters.csv'
$adRows | Sort-Object AdapterKeyName | Export-Csv -NoTypeInformation -Encoding UTF8 $adCsv

# --- README（要約） ---
$readme = @()
$readme += "DNS Update Registry Report"
$readme += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
$readme += ""
$readme += "Global settings path:"
$readme += "  HKLM:\\SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters"
$readme += ""
$readme += "Checked values and defaults:"
$readme += "  DisableDynamicUpdate (DWORD, default 0)"
$readme += "  DisableReverseAddressRegistrations (DWORD, default 0)"
$readme += "  DefaultRegistrationRefreshInterval (DWORD, default 86400 seconds)"
$readme += "  DefaultRegistrationTTL (DWORD, default 1200 seconds)"
$readme += "  UpdateSecurityLevel (DWORD, default 0; 0|0x10|0x100)"
$readme += "  DisableReplaceAddressesInConflicts (DWORD, default 0)"
$readme += ""
$readme += "Per-interface:"
$readme += "  HKLM:\\...\\Parameters\\Interfaces\\<GUID>\\DisableDynamicUpdate (DWORD, default 0)"
$readme += ""
$readme += "Per-adapter:"
$readme += "  HKLM:\\...\\Parameters\\Adapters\\<InterfaceName>\\MaxNumberOfAddressesToRegister (DWORD, default 1)"
$readme += ""
$readme += "Effective DNS update = (Global DisableDynamicUpdate == 0) AND (Per-Interface DisableDynamicUpdate == 0)"
$readme | Out-File -FilePath (Join-Path $OutDir 'README.txt') -Encoding UTF8

Write-Host "✅ Completed. Output -> $OutDir"
