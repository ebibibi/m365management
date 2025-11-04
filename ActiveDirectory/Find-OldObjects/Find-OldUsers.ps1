$oneYearAgo = (Get-Date).AddYears(-1)

Get-ADUser -Filter {Enabled -eq $true -and lastLogonTimestamp -lt $oneYearAgo} -Properties lastLogonTimestamp |
    Select-Object Name,
                  SamAccountName,
                  @{Name='LastLogonDate';Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}} |
    Sort-Object LastLogonDate
