# Connect to Azure AD and Exchange Online using modern authentication
Connect-AzureAD
Connect-MsolService
Connect-ExchangeOnline

# Get the user account to be archived
$username = Read-Host "Enter the UserPrincipalName of the user to be archived"
$user = Get-AzureADUser -ObjectId $username

$groups | ForEach-Object {
    Write-Host "Removing user '$($user.UserPrincipalName)' from group '$($_.DisplayName)'"
    Remove-AzureADGroupMember -ObjectId $_.ObjectId -MemberId $user.ObjectId -Confirm:$false
}

# Format the user's group memberships in a table
$table = $groups | Select-Object DisplayName, ObjectId | Format-Table -AutoSize | Out-String

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

# Disable the user account
Set-AzureADUser -ObjectId $user.ObjectId -AccountEnabled $false

# Output the user's group memberships in a table
Write-Host "The user was a member of the following groups:" -ForegroundColor Yellow
Write-Host $table
