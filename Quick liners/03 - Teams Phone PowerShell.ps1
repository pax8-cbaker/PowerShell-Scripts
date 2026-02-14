#-------------------------------------------------------------#
# Assign phone numbers to users
Connect-MicrosoftTeams
$numbers = Import-CSV -Path "C:\Users\CalebBaker\OneDrive - PAX8\Documents\02 - Teams Phone\UV&S Technology - TEC\AssignNumbers-TEC.csv"

foreach ($number in $numbers) {
    Write-Host "Assigning phone number $($number.DID) to user $($number.UPN) with location ID $($number.LocaID)" -ForegroundColor Cyan
    Try {
        Set-CsPhoneNumberAssignment -Identity $number.UPN -PhoneNumber $number.DID -PhoneNumberType CallingPlan -LocationID $number.LocaID
        Write-Host "Successfully assigned phone number $($number.DID) to user $($number.UPN)`n" -ForegroundColor Green
    }
    Catch {
        Write-Host "Failed to assign phone number $($number.DID) to user $($number.UPN). Error: $_`n" -ForegroundColor Red
    }
}

#-------------------------------------------------------------#
# Assign Caller ID policy to users
Get-CsCallingLineIdentity
$policyAssignments = Import-Csv -Path "C:\Users\CalebBaker\OneDrive - PAX8\Documents\02 - Teams Phone\CallerID.csv"

foreach ($assignment in $policyAssignments) {
    Write-Host "Assigning Caller ID policy $($assignment.CallerIDPolicy) to user $($assignment.UserPrincipalName)" -ForegroundColor Cyan
    Try {
        Grant-CsCallingLineIdentity -Identity $assignment.UserPrincipalName -PolicyName $assignment.CallerIDPolicy
        Write-Host "Successfully assigned Caller ID policy $($assignment.CallerIDPolicy) to user $($assignment.UserPrincipalName)`n" -ForegroundColor Green
    }
    Catch {
        Write-Host "Failed to assign Caller ID policy $($assignment.CallerIDPolicy) to user $($assignment.UserPrincipalName). Error: $_`n" -ForegroundColor Red
    }
}

#-------------------------------------------------------------#
# Unassign phone numbers from users
$usersToUnassign = Import-CSV -Path "C:\Users\CalebBaker\OneDrive - PAX8\Documents\02 - Teams Phone\UnassignNumbers-TEC.csv"

foreach ($user in $usersToUnassign) {
    Write-Host "Unassigning phone number from user $($user.UPN)" -ForegroundColor Cyan
    Try {
        Remove-CsPhoneNumberAssignment -Identity $user.UPN -RemoveAll -ErrorAction Stop | Out-Null
        Write-Host "Successfully unassigned phone number from user $($user.UPN)`n" -ForegroundColor Green
    }
    Catch {
        Write-Host "Failed to unassign phone number from user $($user.UPN). Error: $_`n" -ForegroundColor Red
    }
}