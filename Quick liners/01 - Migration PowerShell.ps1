#-------------------------------------------------------------#
# Update primary SMTP for a specific group
Get-UnifiedGroup -Identity "3PL CONSOLIDATION PROJECT - LUX (Axis Global)"
Set-UnifiedGroup -Identity "3PL CONSOLIDATION PROJECT - LUX (Axis Global)" -PrimarySmtpAddress "3PLCONSOLIDATIONPROJECT-LUX@axisg.com"
#-------------------------------------------------------------#
# Export of primary SMTPs of all onmicrosoft.com accounts
Get-ExoMailbox -ResultSize Unlimited |
    Where-Object { $_.PrimarySmtpAddress -like "*.onmicrosoft*" } |
    Select-Object DisplayName, UserPrincipalName, PrimarySmtpAddress, RecipientTypeDetails |
    Sort-Object DisplayName |
    Export-Csv -Path "C:\Users\CalebBaker\OneDrive - PAX8\Desktop\primarysmtp.csv" -NoTypeInformation
#-------------------------------------------------------------#
# CSV with SamAccountName, PrimarySMTP, Alias
Import-CSV "C:\Users\Administrator\Desktop\smtp.csv" | % { Get-AdUser -Identity $_.SamAccountName -Properties proxyAddresses | Select-Object Name, @{L = "ProxyAddresses"; E = { ($_.ProxyAddresses -like 'smtp:*') -join ";"}} } | Export-CSV "[path]\proxyaddresses.csv" -Notypeinformation
Import-CSV "C:\Users\Administrator\Desktop\smtp.csv" | % { Set-AdUser -Identity $_.SamAccountName -Replace @{proxyAddresses="SMTP:$($_.PrimarySMTP)", "smtp:$($_.Alias)"} }
#-------------------------------------------------------------#
# Update UPNs and Primary SMTP via Microsoft Graph
Connect-MgGraph -Scopes user.readwrite.all

Import-Csv "C:\Users\CalebBaker\OneDrive - PAX8\Documents\01 - Migrations\Planters\UPNChange.csv" | ForEach-Object {
    $body = @{
        mail   = $_.NewUPN
        userPrincipalName = $_.NewUPN
    } | ConvertTo-Json

    Write-Host "Updating user: $($_.UserID) to new UPN: $($_.NewUPN)" -ForegroundColor Green
    Try {
        Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/users/$($_.UserID)" -Body $body | Out-Null
        Write-Host "Successfully updated user: $($_.UserID)" -ForegroundColor Green
    }
    Catch {
        Write-Host "Failed to update user: $($_.UserID). Error: $_" -ForegroundColor Red
    }
}

#-------------------------------------------------------------#
# Update aliases
Connect-ExchangeOnline

Import-CSV "C:\Users\CalebBaker\OneDrive - PAX8\Documents\1_Migrations\Sun Theory\aliases_sunhouse.csv"|%{Set-Mailbox -Identity $_.UPN -EmailAddresses @{Add=$_.Alias}}

#-------------------------------------------------------------#
# Assign licenses
Connect-MgGraph -Scopes user.readwrite.all
Import-Csv "C:\Users\CalebBaker\OneDrive - PAX8\Documents\02 - Teams Phone\UV&S Technology - TEC\LicenseAssignments-TEC.csv" | ForEach-Object {
    $body = @{
        addLicenses = @(
            @{
                skuId = $_.SKU
            }
        )
        removeLicenses = @()
    } | ConvertTo-Json

    Write-Host "Assigning license to user: $($_.UserPrincipalName) with SKU: $($_.SKU)" -ForegroundColor Cyan
    Try {
        Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$($_.UserPrincipalName)/assignLicense" -Body $body -ErrorAction Stop | Out-Null
        Write-Host "Successfully assigned license to user: $($_.UserPrincipalName)`n" -ForegroundColor Green
    }
    Catch {
        Write-Host "Failed to assign license to user: $($_.UserPrincipalName). Error: $_" -ForegroundColor Red
    }
}
#-------------------------------------------------------------#
# Enable OneDrive for users (Use Windows PowerShell)
Connect-SPOService -Url https://planterscoop607-admin.sharepoint.com

$users = Get-Content -Path "C:\Users\CalebBaker\OneDrive - PAX8\Documents\01 - Migrations\Planters\users.txt"
Request-SPOPersonalSite -UserEmails $users

# Confirm OneDrive URLs
Get-SPOSite -IncludePersonalSite $true -Limit all -Filter "Url -like '-my.sharepoint.com/personal/'" | Select-Object -ExpandProperty Url
#-------------------------------------------------------------#
# Manage Exchange Online protocols for users
Connect-ExchangeOnline
$users = Import-Csv "C:\Users\CalebBaker\OneDrive - PAX8\Documents\01 - Migrations\Planters\Users.csv"
#$allUsers = Get-Mailbox

foreach ($user in $users) {
    Write-Host "Configuring protocols for user: $($user.UPN)" -ForegroundColor Cyan
    Try {
        # Disable all protocols except OWA
        #Set-CASMailbox -Identity $user.UPN -PopEnabled $False -ImapEnabled $False -MAPIEnabled $False -ActiveSyncEnabled $False -EwsEnabled $False -OWAEnabled $True -ErrorAction Stop
        # Enable all protocols
        #Set-CASMailbox -Identity $user.UPN -PopEnabled $True -ImapEnabled $True -MAPIEnabled $True -ActiveSyncEnabled $True -EwsEnabled $True -OWAEnabled $True -ErrorAction Stop
        Write-Host "Successfully configured protocols for user: $($user.UPN)`n" -ForegroundColor Green
    }
    Catch {
        Write-Host "Failed to configure protocols for user: $($user.UPN). Error: $_`n" -ForegroundColor Red
    }
}
#-------------------------------------------------------------#
# Bulk reset passwords for users and forces password change at next sign-in
# Install-Module -Name ImportExcel -Force
Connect-MgGraph -Scopes UserAuthenticationMethod.ReadWrite.All, Organization.Read.All

$tenantInfo = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/organization"
$tenantName = $tenantInfo.value[0].displayName

$users = Import-CSV -Path "C:\Users\CalebBaker\OneDrive - PAX8\Documents\01 - Migrations\Adkins\passwordResetTemplate.csv"

# Example for a generic password for all users
$genericPassword = "Adkins293482!"
$body = @{ newPassword = $genericPassword } | ConvertTo-Json

# Empty body for system-generated password
#$body = @{} | ConvertTo-Json

# Create array to store results
$passwordResetResults = @()

foreach ($user in $users) {
    Write-Host "Resetting password for user: $($user.UPN)" -ForegroundColor Cyan
    try {
        $resetPassword = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$($user.UPN)/authentication/methods/28c10230-6103-485e-b985-444c60001490/resetPassword" -Body $body -ErrorAction Stop
        
        # Add result to array
        $passwordResetResults += [PSCustomObject]@{
            UPN = $user.UPN
            #NewPassword = $resetPassword.newPassword
            GenericPassword = $genericPassword
            Status = "Success"
        }
        
        Write-Host "Password reset successfully for user: $($user.UPN)`n" -ForegroundColor Green
    }
    catch {
        # Add failed result to array
        $passwordResetResults += [PSCustomObject]@{
            UPN = $user.UPN
            NewPassword = $null
            Status = "Failed: $_"
        }
        
        Write-Host "Failed to reset password for user: $($user.UPN). Error: $_`n" -ForegroundColor Red
    }
}

# Export results to XLSX
$passwordResetResults | Export-Excel -Path "C:\Users\CalebBaker\OneDrive - PAX8\Desktop\$($tenantName)_passwordreset_results.xlsx" -AutoSize -WorksheetName "PasswordResetResults"
Write-Host "Results exported to $($tenantName)_passwordreset_results.xlsx" -ForegroundColor Green

#-------------------------------------------------------------#