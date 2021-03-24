##############################################################################################
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
##############################################################################################

<#
	.SYNOPSIS
		Configure mailbox audit logging for the Unified Audit Log and enable the recommended
        actions to log

	.DESCRIPTION

        This script has two phases:
        1. Enable mailbox audit logging for all mailboxes that do not have logging enabled
        by default, i.e., resource mailboxes, and explicitly for those that are enabled 
        by default but do not have their mailbox audit events being sent to the Unified Audit Log.
        
        2. For all mailboxes that do not have the recommended actions being logged, enable the
        the recommended actions (overwriting whatever is currently set). The recommended actions
        are those enabled by default. (This phase can be skipped with the DoNotSetAuditActions parameter.)

        You need to be connected to Exchange Online using the v2 management
        module in order for this script to run.  If the module is not installed,
        run Install-Module ExchangeOnlineManagement.

    .PARAMETER  DoNotSetAuditActions
        Switch to skip the second phase of the script, which sets the recommended actions to be logged.
    
    .NOTES
        Version 2.1
        March 24, 2021
#>

#Requires -Module ExchangeOnlineManagement
[CmdletBinding()]
Param (
    [Switch]$DoNotSetAuditActions
)

if (-not(Get-Command -Name Get-Mailbox -ErrorAction SilentlyContinue))
    {
    throw "Please connect to Exchange Online before running this script."
    }

$today = Get-Date -Format yyyyMMdd

#Function that determines if a mailbox's actions that are being logged are at least the
#ones that are recommended
function Compare-AuditActions ($mailbox,$logonType)
    {
    #Actions enabled by default for logon types (plus MailboxLogin for owner) and licensing
    [System.Collections.ArrayList]$NonE5OwnerActions = 'Update','MoveToDeletedItems','SoftDelete','HardDelete','UpdateFolderPermissions','UpdateInboxRules','UpdateCalendarDelegation','ApplyRecord'
    [System.Collections.ArrayList]$NonE5DelegateActions = 'Update','MoveToDeletedItems','SoftDelete','HardDelete','SendAs','SendOnBehalf','Create','UpdateFolderPermissions','UpdateInboxRules','ApplyRecord'
    [System.Collections.ArrayList]$NonE5AdminActions = 'Update','MoveToDeletedItems','SoftDelete','HardDelete','SendAs','SendOnBehalf','Create','UpdateFolderPermissions','UpdateInboxRules','UpdateCalendarDelegation','ApplyRecord'
    [System.Collections.ArrayList]$E5OwnerActions = $NonE5OwnerActions + 'MailItemsAccessed','Send'
    [System.Collections.ArrayList]$E5DelegateActions = $NonE5DelegateActions + 'MailItemsAccessed'
    [System.Collections.ArrayList]$E5AdminActions = $NonE5AdminActions + 'MailItemsAccessed','Send'
    
    if ($mailbox.DefaultAuditSet -contains $logonType)
        {
        #Enabled actions are default
        return $false
        }
    else 
        {
        #Determine if mailbox that is not using default audit set is logging at least the default actions
        switch ($logonType)
            {
            'Admin' {$diffObject = $mailbox.AuditAdmin}
            'Delegate' {$diffObject = $mailbox.AuditDelegate}
            'Owner' {$diffObject = $mailbox.AuditOwner}
            }
        
        #Set reference actions based on mailbox licensing
        if ($mailbox.PersistedCapabilities -contains 'BPOS_S_EquivioAnalytics' `
            -or $mailbox.PersistedCapabilities -contains 'M365Auditing')
            {
            #Mailbox has E5 license or Compliance add-on license
            switch ($logonType)
                {
                'Admin' {$refObject = $E5AdminActions}
                'Delegate' {$refObject = $E5DelegateActions}
                'Owner' {$refObject = $E5OwnerActions}
                }
            }
        else
            {
            switch ($logonType)
                {
                'Admin' {$refObject = $NonE5AdminActions}
                'Delegate' {$refObject = $NonE5DelegateActions}
                'Owner' {$refObject = $NonE5OwnerActions}
                }
            }

        if (Compare-Object -ReferenceObject $refObject -DifferenceObject $diffObject |
            Where-Object {$_.SideIndicator -eq '<='})
            {
            #Mailbox is missing recommended actions
            return $true
            }
        else
            {
            return $false
            }
        }
    }

Write-Host "$(Get-Date) Beginning phase 1..." -ForegroundColor Green

<#
    Get relevant mailboxes:

    1.Those that do not have audit logging enabled, which applies only to resource mailboxes.
    2.Those that are enabled by global auditing (aka on-by-default) but whose events
      are not being sent to the Unified Audit Log, which are shared mailboxes and non-E5 user
      mailboxes that have not been explicitly enabled for audit logging.

      Important note: Only a server-side filter can be used to determine #2.  The
      AuditEnabled property will always have a value of True, so it requires Exchange to determine
      mailboxes that are implicitly True because of global auditing or explicitly True via
      the Set-Mailbox cmdlet.
#>
Write-Host "$(Get-Date) Getting mailboxes whose events are not being sent to the Unified Audit Log..." `
    -ForegroundColor Green
[array]$nonUALMailboxes = Get-EXOMailbox -ResultSize:Unlimited -Filter `
    'AuditEnabled -ne $true -and PersistedCapabilities -ne "BPOS_S_EquivioAnalytics" -and PersistedCapabilities -ne "M365Auditing"'

Write-Host "$(Get-Date) $($nonUALMailboxes.Count) mailboxes were returned." `
    -ForegroundColor Green
if ($nonUALMailboxes.Count -gt 0)
    {
    Write-Host "$(Get-Date) Configuring the mailboxes so audit logging events from this point forward are sent to the Unified Audit Log..." `
        -ForegroundColor Green

    [System.Collections.ArrayList]$UALMailboxLog = @()
    $UALMailboxLogFile = "MailboxAuditLoggingEnabled-$today.log"
    $i = 1
    foreach ($mb in $nonUALMailboxes)
        {
        Write-Progress -Activity "Enabling audit logging" -Status "Processing mailbox for $($mb.DisplayName)" `
            -PercentComplete ($i/$nonUALMailboxes.Count*100)
        Write-Verbose -Message "Enabling audit logging for $($mb.DisplayName)."
        #Manually/explicitly enable audit logging for the mailbox
        Set-Mailbox -Identity $mb.DistinguishedName -AuditEnabled $true
        $UALMailboxLog.Add($mb.UserPrincipalName) | Out-Null
        $i++
        }
    Write-Progress -Activity "Enabling audit logging" -Status " " -Completed
    $UALMailboxLog | Out-File -FilePath $UALMailboxLogFile -Append
    Write-Host "$(Get-Date) Mailboxes that have been modified are logged in $UALMailboxLogFile in the current directory." `
        -ForegroundColor Green
    }
Write-Host "$(Get-Date) Phase 1 is complete." -ForegroundColor Green

if ($DoNotSetAuditActions -eq $false)
    {
    Write-Host "$(Get-Date) Beginning phase 2..." -ForegroundColor Green
    
    #Get all mailboxes because client-side processing is necessary
    Write-Host "$(Get-Date) Getting all non-Group mailboxes. This may take some time..." `
        -ForegroundColor Green
    [array]$allMailboxes = Get-EXOMailbox -ResultSize Unlimited `
        -Properties DefaultAuditSet,AuditAdmin,AuditDelegate,AuditOwner,PersistedCapabilities
    
    Write-Host "$(Get-Date) $($allMailboxes.Count) mailboxes were returned." -ForegroundColor Green
    
    #Filter results to include those that don't have the recommended actions enabled for any of
    #Admin, Delegate, and Owner
    Write-Host "$(Get-Date) Filtering mailboxes that do not have the recommended actions enabled..." `
        -ForegroundColor Green
    [System.Collections.ArrayList]$nonDefaultActionsMB =@()
    $j = 1
    foreach ($mb in $allMailboxes){
        Write-Progress -Activity "Determining mailboxes that do not have the recommended actions enabled" `
            -Status "Processing mailbox for $($mb.DisplayName)" -PercentComplete ($j/$allMailboxes.Count*100)

        if ($mb.DefaultAuditSet -join '' -ne 'AdminDelegateOwner') {
            if ((Compare-AuditActions -mailbox $mb -logonType Owner) -or
                (Compare-AuditActions -mailbox $mb -logonType Delegate) -or
                (Compare-AuditActions -mailbox $mb -logonType Admin)){
                    
                    $nonDefaultActionsMB.Add($mb) | Out-Null
            }
        }
        $j++
    }
    Write-Progress -Activity "Determining mailboxes that do not have the recommended actions enabled" -Status " " -Completed

    Write-Host "$(Get-Date) $($nonDefaultActionsMB.Count) mailboxes do not have at least the recommended actions enabled." `
        -ForegroundColor Green
    
    if ($nonDefaultActionsMB.Count -gt 0)
        {
        Write-Host "$(Get-Date) Resetting audit logging so the default actions are enabled..." `
            -ForegroundColor Green
        
        [System.Collections.ArrayList]$RecommendedActionsLog = @()
        $RecommendedActionsLogFile = "MailboxAuditActionsUpdated-$today.log"
        $k = 1
        foreach ($mb in $nonDefaultActionsMB)
            {
            Write-Progress -Activity "Restting default actions" -Status "Processing mailbox for $($mb.DisplayName)" `
                -PercentComplete ($k/$nonDefaultActionsMB.Count*100)
            Write-Verbose -Message "Resetting default actions for $($mb.DisplayName)."

            #Configure the mailbox to use the default audit set for all logon types
            Set-Mailbox -Identity $mb.DistinguishedName -DefaultAuditSet Admin,Delegate,Owner
            $RecommendedActionsLog.Add($mb.UserPrincipalName) | Out-Null
            $k++
            }
        Write-Progress -Activity "Resetting default actions" -Status " " -Completed
        $RecommendedActionsLog | Out-File -FilePath $RecommendedActionsLogFile -Append
        Write-Host "$(Get-Date) Mailboxes that have been modified are logged in $RecommendedActionsLogFile in the current directory." `
            -ForegroundColor Green
        }
    Write-Host "$(Get-Date) Phase 2 is complete." -ForegroundColor Green
    }
else
    {
    Write-Host "$(Get-Date) Phase 2 has been skipped because the DoNotSetAuditActions switch was used." `
        -ForegroundColor Green
    }
