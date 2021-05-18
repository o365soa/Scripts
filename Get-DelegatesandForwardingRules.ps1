############################################################################
#This sample script is not supported under any Microsoft standard support program or service.
#This sample script is provided AS IS without warranty of any kind.
#Microsoft further disclaims all implied warranties including, without limitation, any implied
#warranties of merchantability or of fitness for a particular purpose. The entire risk arising
#out of the use or performance of the sample script and documentation remains with you. In no
#event shall Microsoft, its authors, or anyone else involved in the creation, production, or
#delivery of the scripts be liable for any damages whatsoever (including, without limitation,
#damages for loss of business profits, business interruption, loss of business information,
#or other pecuniary loss) arising out of the use of or inability to use the sample script or
#documentation, even if Microsoft has been advised of the possibility of such damages.
############################################################################

<#
	.SYNOPSIS
		Get mailbox delegates and forwarding rules for one or all mailboxes in an organization.

	.DESCRIPTION

        For one or all mailboxes, get a list of who has full access permission, delegate (folder) access,
        and inbox rules that forward email.  You can choose to skip getting any of these phases,
        usually for time constraints.

        To get mailbox delegates, the primary calendar folder's display name is needed, which the script
        will determine.  This allows for support of non-English mailboxes.  If all mailboxes use
        English display names, you can make the script complete in less time by using the
        NoLocalization switch parameter.

        Output is saved to files appended with either the mailbox name or the current date/time,
        accordingly.

    .PARAMETER Identity
        Valid identity of a mailbox.  Absence of this parameter will cause the script to
        process all mailboxes.
    
    .PARAMETER DoNotGetFullAccessPermissions
        Switch to not get security principals that have been granted full mailbox access.

    .PARAMETER DoNotGetFolderDelegates
        Switch to not get mailbox delegates (those granted delegate access to the calendar and, optionally,
        other folders).

    .PARAMETER DoNotGetInboxRules
        Switch to not get inbox rules that forward emails.

    .PARAMETER NoLocalization
        Switch to skip getting the localized display name of the calendar folder.  If all mailboxes
        use an English locale, the script will complete in less time by using this parameter because
        it will always use "Calendar" as the folder display name instead of looking up the display name.
    
    .NOTES
        Version 2.0
        May 18, 2021
#>

#Requires -Modules ExchangeOnlineManagement
param (
    [string]$Identity,
    [switch]$DoNotGetFullAccessPermissions,
    [switch]$DoNotGetFolderDelegates,
    [switch]$DoNotGetInboxRules,
    [switch]$NoLocalization
)

# Connect to Exchange Online if necessary
if (-not(Get-Command -Name Get-OrganizationConfig -ErrorAction SilentlyContinue)){
    Connect-ExchangeOnline
}

$allMailboxes = @()
if ($Identity) {
    $allMailboxes += Get-ExoMailbox -Identity $Identity -ErrorVariable InvalidMailbox
    $fileSuffix = $allMailboxes.DisplayName
}
else {
    $allMailboxes += Get-ExoMailbox -ResultSize unlimited
    $fileSuffix = Get-Date -Format yyyyMMdd
}
    
$userInboxRules = @()
$userFullAccessDelegates = @()
$userMailboxFolderDelegates = @()

# Determine tasks to perform for progress bar
$tasks = @()
if ($DoNotGetFullAccessPermissions -eq $false) {$tasks += "full mailbox access"}
if ($DoNotGetFolderDelegates -eq $false) {$tasks += "calendar delegates"}
if ($DoNotGetInboxRules -eq $false) {$tasks += "Inbox rules"}

$i = 1
foreach ($mailbox in $allMailboxes)
{
    Write-Progress -Activity "Checking mailboxes for $($tasks -join ', ').  Now checking $($mailbox.DisplayName)" `
        -Status "Overall progress" -PercentComplete ($i/$allMailboxes.Count*100)
    
    if ($DoNotGetInboxRules -eq $false) {
        # Get inbox rules that forward or redirect
        $userInboxRules += Get-InboxRule -Mailbox $mailbox.DistinguishedName | 
            Select-Object -Property @{n='Mailbox';e={$mailbox.PrimarySmtpAddress}}, Name, Enabled, ForwardTo, ForwardAsAttachmentTo, RedirectTo, DeleteMessage | 
            Where-Object {($null -ne $_.ForwardTo) -or ($null -ne $_.ForwardAsAttachmentTo) -or ($null -ne $_.RedirectsTo)}
    }

    if ($DoNotGetFullAccessPermissions -eq $false) {
        # Get security principals with Full Access permission to mailbox        
        $userFullAccessDelegates += Get-MailboxPermission -Identity $mailbox.DistinguishedName | `
            Where-Object {($_.IsInherited -ne "True") -and ($_.User -notlike "*SELF*")} | `
            Select-Object -Property @{n="Mailbox";e={$_.Identity}},@{n="Assignee";e={$_.User}},AccessRights
    }

    if ($DoNotGetFolderDelegates -eq $false) {
        # Get delegate access (calendar) to mailbox
        if ($NoLocalization) {
            $primaryCalendarPath = '/Calendar'
        }
        else {
            # Get localized calendar folder name
            $primaryCalendarPath = Get-MailboxFolderStatistics -Identity $mailbox.DistinguishedName -FolderScope Calendar | 
                Where-Object {$_.FolderType -eq 'Calendar'} | Select-Object -ExpandProperty FolderPath
            $userMailboxFolderDelegates += Get-MailboxFolderPermission -Identity "$($mailbox.DistinguishedName):\$($primaryCalendarPath.Substring(1))" | 
                Where-Object {$_.SharingPermissionFlags -like "*Delegate*"} | 
                Select-Object -Property @{n="Mailbox";e={$mailbox.UserPrincipalName}},@{n="Delegate";e={$_.User}}
        }
    }

    $i++
}

if (-not($InvalidMailbox)){
    if ($DoNotGetInboxRules -eq $false) {$userInboxRules | Export-Csv -Path MailForwardingRules-$fileSuffix.csv -NoTypeInformation}
    if ($DoNotGetFullAccessPermissions -eq $false) {$userFullAccessDelegates | Export-Csv -Path FullAccessPermissions-$fileSuffix.csv -NoTypeInformation}
    if ($DoNotGetFolderDelegates -eq $false) {$userMailboxFolderDelegates | Export-Csv -Path DelegateAccessPermissions-$fileSuffix.csv -NoTypeInformation}
}