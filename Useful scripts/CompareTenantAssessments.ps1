# Compare tenant results
Clear-Host
$importSourcePath = (Read-Host "Enter the path for the source tenant export (e.g., C:\Users\CalebBaker\OneDrive - PAX8\Desktop\source.xlsx)").Trim('"')
$importDestinationPath = (Read-Host "Enter the path for the destination tenant export (e.g., C:\Users\CalebBaker\OneDrive - PAX8\Desktop\destination.xlsx)").Trim('"')
$resultsPath = (Read-Host "Enter the path for the comparison results export (e.g., C:\Users\CalebBaker\OneDrive - PAX8\Desktop\Comparison_Results.xlsx)").Trim('"')

# Find duplicate users based on DisplayName
Write-Progress -Activity "Comparing tenants" -Status "Comparing users..." -PercentComplete 25
$sourceFullUsers = Import-Excel -Path $importSourcePath -WorksheetName "Users"
$destinationFullUsers = Import-Excel -Path $importDestinationPath -WorksheetName "Users"

$sDisplayNames = $sourceFullUsers | Select-Object -ExpandProperty DisplayName -Unique
$dDisplayNames = $destinationFullUsers | Select-Object -ExpandProperty DisplayName

$duplicateUsersInDestination = $sDisplayNames | Where-Object { $_ -in $dDisplayNames }

$resultsUsers = foreach ($duplicateUser in $duplicateUsersInDestination) {
    $matchesUsers = $destinationFullUsers | Where-Object { $_.DisplayName -eq $duplicateUser }
    foreach ($match in $matchesUsers) {
        [PSCustomObject]@{
            DisplayName = $match.DisplayName
            DestinationUserPrincipalName = $match.UserPrincipalName
            UserType = $match.UserType
            AccountEnabled = $match.AccountEnabled
            Licenses = $match.Licenses
        }
    }
}

# Find duplicate mailboxes based on DisplayName
Write-Progress -Activity "Comparing tenants" -Status "Comparing mailboxes..." -PercentComplete 50
$sourceFullMailboxes = Import-Excel -Path $importSourcePath -WorksheetName "Mailboxes"
$destinationFullMailboxes = Import-Excel -Path $importDestinationPath -WorksheetName "Mailboxes"

$sDisplayNames = $sourceFullMailboxes | Select-Object -ExpandProperty DisplayName -Unique
$dDisplayNames = $destinationFullMailboxes | Select-Object -ExpandProperty DisplayName

$duplicateMbxsInDestination = $sDisplayNames | Where-Object { $_ -in $dDisplayNames }

$resultsMailboxes = foreach ($duplicateMailbox in $duplicateMbxsInDestination) {
    $matchesMailboxes = $destinationFullMailboxes | Where-Object { $_.DisplayName -eq $duplicateMailbox }
    foreach ($match in $matchesMailboxes) {
        [PSCustomObject]@{
            DisplayName = $match.DisplayName
            DestinationPrimarySmtpAddress = $match.PrimaryEmail
            MailboxType = $match.RecipientType
            Size = $match."Size(GB)"
            LitigationHoldEnabled = $match.LitigationHold
            ArchiveEnabled = $match.ArchiveEnabled
        }
    }
}

# Compare Teams based on TeamName
Write-Progress -Activity "Comparing tenants" -Status "Comparing Teams..." -PercentComplete 75
$sourceFullTeams = Import-Excel -Path $importSourcePath -WorksheetName "Teams"
$destinationFullTeams = Import-Excel -Path $importDestinationPath -WorksheetName "Teams"

$sTeamNames = $sourceFullTeams | Select-Object -ExpandProperty TeamName -Unique
$dTeamNames = $destinationFullTeams | Select-Object -ExpandProperty TeamName

$duplicateTeamsInDestination = $sTeamNames | Where-Object { $_ -in $dTeamNames }

$resultsTeams = foreach ($duplicateTeam in $duplicateTeamsInDestination) {
    $matchesTeams = $destinationFullTeams | Where-Object { $_.TeamName -eq $duplicateTeam }
    foreach ($match in $matchesTeams) {
        [PSCustomObject]@{
            TeamName = $match.TeamName
            DestinationEmail = $match.Email
            ArchiveEnabled = $match.IsArchived
        }
    }
}

# Compare SharePoint sites and M365 groups based on Title
Write-Progress -Activity "Comparing tenants" -Status "Comparing SharePoint sites..." -PercentComplete 90
$sourceFullSharePoint = Import-Excel -Path $importSourcePath -WorksheetName "SharePointSites" | Where-Object { $_.IsTeamsConnected -eq $false }
$destinationFullSharePoint = Import-Excel -Path $importDestinationPath -WorksheetName "SharePointSites" | Where-Object { $_.IsTeamsConnected -eq $false }

$sTitle = $sourceFullSharePoint | Select-Object -ExpandProperty Title -Unique
$dTitle = $destinationFullSharePoint | Select-Object -ExpandProperty Title

$duplicateSharePointInDestination = $sTitle | Where-Object { $_ -in $dTitle }

$resultsSharePoint = foreach ($duplicateSharePoint in $duplicateSharePointInDestination) {
    $matchesSharePoint = $destinationFullSharePoint | Where-Object { $_.Title -eq $duplicateSharePoint }
    foreach ($match in $matchesSharePoint) {
        [PSCustomObject]@{
            Title = $match.Title
            DestinationURL = $match.URL
            TeamsConnected = $match.IsTeamsConnected
        }
    }
}

# Export results to XLSX
Write-Progress -Activity "Comparing tenants" -Status "Exporting results..." -PercentComplete 100
$resultsUsers | Export-Excel -Path $resultsPath -WorksheetName "Users" -TableStyle Light1 -AutoSize
$resultsMailboxes | Export-Excel -Path $resultsPath -WorksheetName "Mailboxes" -TableStyle Light1 -AutoSize -Append
$resultsTeams | Export-Excel -Path $resultsPath -WorksheetName "Teams" -TableStyle Light1 -AutoSize -Append
$resultsSharePoint | Export-Excel -Path $resultsPath -WorksheetName "SharePointSites" -TableStyle Light1 -AutoSize -Append -Show
Write-Progress -Activity "Comparing tenants" -Completed
Write-Host "Comparison complete! Results exported to $resultsPath" -ForegroundColor Green