# Install the Microsoft Teams PowerShell module
# Install-Module -Name MicrosoftTeams -Force -AllowClobber

# Import the Microsoft Teams module
Import-Module MicrosoftTeams

# デバイスコード認証で接続（GUI環境がなくても動作する）
Connect-MicrosoftTeams -UseDeviceAuthentication

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

    # Add team info to the output array
    $output += [PSCustomObject]@{
        TeamName  = $teamName
        ChannelName  = "N/A"
        Owners    = $ownerNames
        Members   = $memberNames
    }

    # Get the list of channels for the team
    $channels = Get-TeamChannel -GroupId $teamId

    foreach ($channel in $channels) {
        $channelName = $channel.DisplayName

        # Get the channel members
        $channelMembers = Get-TeamChannelUser -GroupId $teamId -DisplayName $channelName -Role Member
        $channelMemberNames = $channelMembers | ForEach-Object { $_.User }
        $channelMemberNames = $channelMemberNames -join ", "

        # Get the channel owners
        $channelOwners = Get-TeamChannelUser -GroupId $teamId -DisplayName $channelName -Role Owner
        $channelOwnerNames = $channelOwners | ForEach-Object { $_.User }
        $channelOwnerNames = $channelOwnerNames -join ", "

        # Add channel info to the output array
        $output += [PSCustomObject]@{
            TeamName     = $teamName
            ChannelName  = $channelName
            Owners = $channelOwnerNames
            Members = $channelMemberNames
        }
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
