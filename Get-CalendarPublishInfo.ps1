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
 
#Requires -Modules ExchangeOnlineManagement

<#
	.SYNOPSIS
		This script exports mailbox calendar publishing settings

	.DESCRIPTION
	    Iterates through mailboxes and outputs all with (anonymous) calendar publishing enabled, including sharing level.

    .PARAMETER  OutCsv
        Will export the results to a CSV File in the script root called Get-CalendarPublishInfo.csv

	.EXAMPLE
		.\Get-CalendarPublishInfo.ps1

	.EXAMPLE
		.\Get-CalendarPublishInfo.ps1 -OutCSV

	.NOTES
		For updates, and more scripts, visit https://github.com/O365AES/Scripts
        Version 1.1
        January 6, 2025
#>

param(
     [switch]$OutCSV
)

$cfs = @()
$i = 1

$mailboxes = Get-ExoMailbox -ResultSize Unlimited | Select-Object -Property Identity,UserPrincipalName,DisplayName
foreach ($mailbox in $mailboxes) {
    Write-Progress -Activity "Getting calendar publishing details" -Status "Processing mailbox $($mailbox.DisplayName)" -PercentComplete ($i/$mailboxes.Count*100)

    # Get users calendar folder settings for their default Calendar folder
    # Get localized folder name
	$primaryCalendarPath = Get-MailboxFolderStatistics -Identity $mailbox.UserPrincipalName -FolderScope Calendar | Where-Object {$_.FolderType -eq 'Calendar'} | Select-Object -ExpandProperty FolderPath
    $cf = Get-MailboxCalendarFolder -Identity "$($mailbox.UserPrincipalName):\$($primaryCalendarPath.Substring(1))"
    
    # If publishing is turned on, add to the result set
    if($cf.PublishEnabled -eq $true) {
        $cfs += New-Object -TypeName psobject -Property @{
            UserPrincipalName=$mailbox.UserPrincipalName
            PublishEnabled=$cf.PublishEnabled
            DetailLevel=$cf.DetailLevel
            PublishedCalendarUrl=$cf.PublishedCalendarUrl
            PublishedICalUrl=$cf.PublishedICalUrl
        }
    }
    $i++
}
Write-Progress -Activity "Getting calendar publishing details" -Status " " -Completed

# Export to CSV if specified
if($cfs.Count -gt 0) {
    if($OutCsv) {
        $cfs | Export-Csv -Path "$PSScriptRoot\Get-CalendarPublishInfo.csv" -NoTypeInformation -NoClobber
    } else {
        return $cfs
    }
} else {
    Write-Host "No users have calendar publishing enabled"
}
