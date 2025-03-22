$creds = Get-StoredCredential -Target "Batcher"
# Define the output path
$outputPath = "C:\ExchangeMigration"
#Start-Transcript -OutputDirectory $outputPath
#Install-Module -Name SharePointPnPPowerShellOnline -Scope AllUsers
#Import-Module ExchangeOnlineManagement
Import-Module SharePointPnPPowerShellOnline -WarningAction SilentlyContinue
Import-Module CredentialManager
$certThumbprint = "5E79521C7F31B61A5870F20BE06CF78E3947655C"
$appid = '0609911f-6d2a-47a7-a49c-d41535b2d4a3'
$Tenantdomain = 'polsinelli.onmicrosoft.com'
$tenantid = 'c824b048-0922-4f96-87fc-bf920e1a756d'
Connect-ExchangeOnline -AppId $appid -Organization $Tenantdomain -CertificateThumbprint $certThumbprint

# Email address for the Distribution Group
$groupEmail = "migrationteam@polsinelli.com"
$groupName = "Migration Team"

# Check if the email address is already in use
$existingGroup = Get-DistributionGroup -Identity $groupEmail -ErrorAction SilentlyContinue
# List of user emails to add to the group
    $userEmails = @("gregory.hall@polsinelli.com")
if ($null -eq $existingGroup) {
    # Create the Distribution Group
    $newGroup = New-DistributionGroup -Name $groupName -Alias ( $groupEmail -replace "@.*", "" ) -PrimarySmtpAddress $groupEmail -MemberJoinRestriction Closed -MemberDepartRestriction Closed

    # Optionally, enable the group to receive emails from external senders
    Set-DistributionGroup -Identity $newGroup.Identity -RequireSenderAuthenticationEnabled $false -HiddenFromAddressListsEnabled $true
    # Add users to the Distribution Group
    foreach ($userEmail in $userEmails) {
        Add-DistributionGroupMember -Identity $newGroup.Identity -Member $userEmail
    }

    Write-Host "Distribution Group '$groupName' created and members added."
} else {
    Write-Host "The email address $groupEmail is already in use. Please choose a unique email address."
}

# Disconnect from the session
Disconnect-ExchangeOnline -Confirm:$false

