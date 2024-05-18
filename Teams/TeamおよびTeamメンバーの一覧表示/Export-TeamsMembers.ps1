# Install the Microsoft Teams PowerShell module
# Install-Module -Name MicrosoftTeams -Force -AllowClobber

# Import the Microsoft Teams module
Import-Module MicrosoftTeams

# Connect to Microsoft Teams
Connect-MicrosoftTeams

# Get the list of Teams
$teams = Get-Team

# Prepare the CSV file
$output = @()

foreach ($team in $teams) {
    $teamId = $team.GroupId
    $teamName = $team.DisplayName

    # Get the team owners
    $owners = Get-TeamUser -GroupId $teamId -Role Owner
    $ownerNames = $owners | ForEach-Object { $_.User }
    $ownerNames = $ownerNames -join ", "

    # Get the team members
    $members = Get-TeamUser -GroupId $teamId -Role Member
    $memberNames = $members | ForEach-Object { $_.User }
    $memberNames = $memberNames -join ", "

    # Add to the output array
    $output += [PSCustomObject]@{
        TeamName  = $teamName
        Owners    = $ownerNames
        Members   = $memberNames
    }
}

# Export to CSV
$tempPath = "TempTeamsReport.csv"
$output | Export-Csv -Path $tempPath -NoTypeInformation

# Convert to UTF-8 with BOM
$finalPath = "TeamsReport.csv"
Get-Content -Path $tempPath | Out-File -FilePath $finalPath -Encoding utf8BOM

# Clean up the temporary file
Remove-Item $tempPath

Write-Host "Teams report exported to $finalPath"