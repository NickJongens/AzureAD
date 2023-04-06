# Connect to Azure AD and Exchange Online using modern authentication
Connect-AzureAD
Connect-MsolService
Connect-ExchangeOnline

# Get the user account to be archived
$username = Read-Host "Enter the UserPrincipalName of the user to be archived"
$user = Get-AzureADUser -ObjectId $username

# Disable the user account
Set-AzureADUser -ObjectId $user.ObjectId -AccountEnabled $false

# Store the user's group memberships
$groups = Get-AzureADUserMembership -ObjectId $user.ObjectId

# Remove the user from all groups
foreach ($group in $groups) {
    Remove-AzureADGroupMember -ObjectId $group.ObjectId -MemberObjectId $user.ObjectId
}

# Format the user's group memberships in a table
$table = $groups | Select-Object DisplayName, ObjectId | Format-Table -AutoSize | Out-String

# Rename the user's display name
$newDisplayName = "ARCHIVED - " + $user.DisplayName
Set-AzureADUser -ObjectId $user.ObjectId -DisplayName $newDisplayName

# Convert the user's mailbox to a shared mailbox
Enable-Mailbox -Identity $user.UserPrincipalName -Shared -Confirm:$false

# Remove the user's Office 365 licenses
$licenses = Get-MsolUserLicense -UserPrincipalName $user.UserPrincipalName
foreach ($license in $licenses) {
    $licenseName = $license.AccountSkuId.Replace(":",".")
    Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -RemoveLicenses $licenseName
}

# Remove the user account
Remove-AzureADUser -ObjectId $user.ObjectId

# Disconnect from Azure AD and Exchange Online
Disconnect-AzureAD
Disconnect-MsolService
Remove-PSSession $ExchangeSession

# Output the user's group memberships in a table
Write-Host "The user was a member of the following groups:" -ForegroundColor Yellow
Write-Host $table
