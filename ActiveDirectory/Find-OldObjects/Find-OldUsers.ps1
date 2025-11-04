# 1年前を基準に設定
$oneYearAgo = (Get-Date).AddYears(-1)

# Active Directoryからユーザーを検索
Get-ADUser -Filter {Enabled -eq $true -and lastLogonTimestamp -lt $oneYearAgo} -Properties lastLogonTimestamp |
    Select-Object Name,
                  SamAccountName,
                  @{Name='LastLogonDate';Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}} |
    Sort-Object LastLogonDate
