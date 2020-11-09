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
 
#Requires -Version 4

<#
	.SYNOPSIS
		This script exports mailbox calendar publishing settings

	.DESCRIPTION
	    Iterates through mailboxes and dumps calender folder permissions, shows publishing permissions.

        This script is useful for highlighting users that have anonymous calendar sharing turned on.

    .PARAMETER  OutCSV
        Will export the results to a CSV File in the script root called Get-CalendarPublishInfo.csv

	.EXAMPLE
		PS C:\> .\Get-CalendarPublishInfo.ps1

	.EXAMPLE
		PS C:\> .\Get-CalendarPublishInfo.ps1 -OutCSV

	.NOTES
		Cam Murray
		Field Engineer - Microsoft
		cam.murray@microsoft.com
		
		For updates, and more scripts, visit https://github.com/O365AES/Scripts
		
		Last update: 29 March 2017

	.LINK
		about_functions_advanced

#>

Param(
     [switch]$OutCSV
)

$cfs = @()

ForEach($mailbox in (Get-Mailbox -ResultSize Unlimited | select Identity,UserPrincipalName)) {
    Write-Verbose "Checking $($mailbox.Identity)"

    # Get users calendar folder settings for their default Calendar folder
    $cf=Get-MailboxCalendarFolder -Identity "$($mailbox.Identity):\Calendar" 
    
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
}

# Export to csv if required
if($cfs.Count -gt 0) {
    if($OutCSV) {
        $cfs | export-csv -Path "$PSScriptRoot\Get-CalendarPublishInfo.csv" -NoTypeInformation -NoClobber
    } else {
        return $cfs
    }
} else {
    Write-Host "No users with publishing enabled"
}
