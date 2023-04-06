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

# Get the user's directory role memberships
if ($user.ObjectId) {
    $roles = Get-AzureADDirectoryRoleMember -ObjectId $user.ObjectId | ForEach-Object {
        Get-AzureADDirectoryRole -ObjectId $_.DirectoryRole.ObjectId
    }
}

# Format the user's roles in a table
$roleTable = $roles | Select-Object DisplayName, ObjectId | Format-Table -AutoSize | Out-String

# Remove the user from all active directory roles
foreach ($role in $roles) {
    Remove-AzureADDirectoryRoleMember -ObjectId $role.ObjectId -MemberObjectId $user.ObjectId
}

# Rename the user's display name
$newDisplayName = "ARCHIVED - " + $user.DisplayName
Set-AzureADUser -ObjectId $user.ObjectId -DisplayName $newDisplayName

# Convert the user's mailbox to a shared mailbox
$mailboxtoconvert = Get-Mailbox $user.UserPrincipalName
Set-Mailbox $mailboxtoconvert.Identity -Type Shared

# Remove the user's Office 365 licenses
$licenses = Get-MsolUserLicense -UserPrincipalName $user.UserPrincipalName
foreach ($license in $licenses) {
    $licenseName = $license.AccountSkuId.Replace(":",".")
    Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -RemoveLicenses $licenseName
}

# Disconnect from Azure AD and Exchange Online
Disconnect-AzureAD
Disconnect-MsolService
Remove-PSSession $ExchangeSession

# Output the user's group memberships in a table
Write-Host "The user was a member of the following groups:" -ForegroundColor Yellow
Write-Host $table
# Output the user's roles in a table
Write-Host "The user had the following directory roles:" -ForegroundColor Yellow
Write-Host $roleTable
