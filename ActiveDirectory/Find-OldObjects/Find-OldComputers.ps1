# Find computers that have not logged on in over one year
$oneYearAgo = (Get-Date).AddYears(-1)
Get-ADComputer -Filter {lastLogonTimestamp -lt $oneYearAgo} -Properties lastLogonTimestamp |
    Select-Object Name, @{Name='LastLogonDate';Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}} |
    Sort-Object LastLogonDate

