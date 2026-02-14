Get-NetFirewallRule -PolicyStore ActiveStore | Where-Object { $_.DisplayName -like "*partofrulename*" }

#-------------------------------------------------------------#

$profile = Get-NetConnectionProfile
Set-NetConnectionProfile -InterfaceIndex $profile.InterfaceIndex -NetworkCategory Private

#-------------------------------------------------------------#

$serviceName = "LTService"
$service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue

if ($service) {
    Write-Host "$serviceName found."
    exit 0
} else {
    Write-Host "$serviceName not found."
    exit 1
}

#-------------------------------------------------------------#
<#
Simple script: finds firewall rules with Name or DisplayName starting with 'RC_' and sets their Profile to Domain,Private,Public.
Run this script from an elevated (Administrator) PowerShell session.
#>

$pattern = 'RC_*'
$rules = Get-NetFirewallRule -ErrorAction SilentlyContinue | Where-Object { $_.Name -like $pattern -or $_.DisplayName -like $pattern }
if (-not $rules) {
    Write-Output "No firewall rules matching '$pattern' were found."
    return
}
foreach ($r in $rules) {
    try {
        Set-NetFirewallRule -Name $r.Name -Profile Domain,Private,Public -ErrorAction Stop
        Write-Output "Updated rule: $($r.Name)"
    } catch {
        Write-Error "Failed to update rule $($r.Name): $_"
    }
}