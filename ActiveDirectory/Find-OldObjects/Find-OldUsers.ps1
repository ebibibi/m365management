$oneYearAgo = (Get-Date).AddYears(-1)

Get-ADUser -Filter {
    Enabled -eq $true -and
    lastLogonTimestamp -lt $oneYearAgo
} -Properties lastLogonTimestamp |
Where-Object {
    $_.SamAccountName -notmatch '^(IUSR_|IWAM_|ASPNET|Guest|krbtgt|DefaultAccount|WDAGUtilityAccount)'
} |
Select-Object Name,
              SamAccountName,
              @{Name='LastLogonDate';Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}} |
Sort-Object LastLogonDate
