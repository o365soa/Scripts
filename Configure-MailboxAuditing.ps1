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
        the recommended actions (preserving any additional actions that are enabled). The recommended actions
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
function Compare-AuditActions ($mailbox)
    {
    #Actions enabled by default for logon types and licensing
    [System.Collections.ArrayList]$NonE5OwnerActions = 'Update','MoveToDeletedItems','SoftDelete','HardDelete','UpdateFolderPermissions','UpdateInboxRules','UpdateCalendarDelegation','ApplyRecord'
    [System.Collections.ArrayList]$NonE5DelegateActions = 'Update','MoveToDeletedItems','SoftDelete','HardDelete','SendAs','SendOnBehalf','Create','UpdateFolderPermissions','UpdateInboxRules','ApplyRecord'
    [System.Collections.ArrayList]$NonE5AdminActions = 'Update','MoveToDeletedItems','SoftDelete','HardDelete','SendAs','SendOnBehalf','Create','UpdateFolderPermissions','UpdateInboxRules','UpdateCalendarDelegation','ApplyRecord'
    [System.Collections.ArrayList]$E5OwnerActions = $NonE5OwnerActions + 'MailItemsAccessed','Send'
    [System.Collections.ArrayList]$E5DelegateActions = $NonE5DelegateActions + 'MailItemsAccessed'
    [System.Collections.ArrayList]$E5AdminActions = $NonE5AdminActions + 'MailItemsAccessed','Send'
    
    $logonTypes = @('AuditOwner','AuditDelegate','AuditAdmin')

    #Set reference actions based on mailbox licensing
    if ($mailbox.PersistedCapabilities -contains 'BPOS_S_EquivioAnalytics' `
    -or $mailbox.PersistedCapabilities -contains 'M365Auditing') {
        $AuditAdminActions = $E5AdminActions
        $AuditDelegateActions = $E5DelegateActions
        $AuditOwnerActions = $E5OwnerActions
    }
    else {
        $AuditAdminActions = $NonE5AdminActions
        $AuditDelegateActions = $NonE5DelegateActions
        $AuditOwnerActions = $NonE5OwnerActions
    }

    foreach ($logonType in $logonTypes) {

        #Actions being logged for the logon type to compare against
        $diffObject = $mailbox.$logonType

        #Set baseline actions for the logon type to use as a reference
        switch ($logonType) {
            'AuditAdmin' {$refObject = $AuditAdminActions}
            'AuditDelegate' {$refObject = $AuditDelegateActions}
            'AuditOwner' {$refObject = $AuditOwnerActions}
        }

        $actionComparison = Compare-Object -ReferenceObject $refObject -DifferenceObject $diffObject
        
        if ($actionComparison | Where-Object {$_.SideIndicator -eq '<='}) {
            #Mailbox is missing recommended actions
            $actionsMissing = $true
        }
        if ($actionComparison | Where-Object {$_.SideIndicator -eq '=>'}) {              
            #Non-default actions are being logged
            New-Variable -Name $('nonDefault'+$logonType) -Value $true
            $nonDefault = $true
        }
    }

    if ($actionsMissing) {
        return New-Object -TypeName psobject -Property @{
            MissingDefaultActions = $true
            HasCustomActions = $(if ($nonDefault) {$true} else {$false})
            HasCustomActionsAdmin = $nonDefaultAuditAdmin
            HasCustomActionsDelegate = $nonDefaultAuditDelegate
            HasCustomActionsOwner = $nonDefaultAuditOwner
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

        #Check only mailboxes not using default audit set for any logon type
        if ($mb.DefaultAuditSet -join '' -ne 'AdminDelegateOwner') {
            
            $mbActionsCheck = Compare-AuditActions -mailbox $mb
            
            #Add mailbox to collection of those that need to be updated and whether extra actions need to be re-added
            if ($mbActionsCheck.MissingDefaultActions -eq $true) {
                $mbToAdd = New-Object -TypeName psobject -Property @{
                    DisplayName = $mb.DisplayName
                    DistinguishedName = $mb.DistinguishedName
                    UserPrincipalName = $mb.UserPrincipalName
                    HasCustomActions = $mbActionsCheck.HasCustomActions
                    HasCustomActionsAdmin = $mbActionsCheck.HasCustomActionsAdmin
                    HasCustomActionsDelegate = $mbActionsCheck.HasCustomActionsDelegate
                    HasCustomActionsOwner = $mbActionsCheck.HasCustomActionsOwner
                    AuditAdmin = $mb.AuditAdmin
                    AuditDelegate = $mb.AuditDelegate
                    AuditOwner = $mb.AuditOwner
                }

            $nonDefaultActionsMB.Add($mbToAdd) | Out-Null
            }
        }
        $j++
    }
    Write-Progress -Activity "Determining mailboxes that do not have the recommended actions enabled" -Status " " -Completed

    Write-Host "$(Get-Date) $($nonDefaultActionsMB.Count) mailboxes do not have at least the recommended actions enabled." `
        -ForegroundColor Green
    
    if ($nonDefaultActionsMB.Count -gt 0)
        {
        Write-Host "$(Get-Date) Adding recommended actions to be logged..." `
            -ForegroundColor Green
        
        [System.Collections.ArrayList]$RecommendedActionsLog = @()
        $RecommendedActionsLogFile = "MailboxAuditActionsUpdated-$today.log"
        $k = 1
        foreach ($mb in $nonDefaultActionsMB)
            {
            Write-Progress -Activity "Adding recommended actions" -Status "Processing mailbox for $($mb.DisplayName)" `
                -PercentComplete ($k/$nonDefaultActionsMB.Count*100)

            #Configure the mailbox to use the default audit set for all logon types
            Write-Verbose -Message "Adding recommended actions for $($mb.DisplayName)."
            Set-Mailbox -Identity $mb.DistinguishedName -DefaultAuditSet Admin,Delegate,Owner
            
            #If additional actions were being logged, add them back
            #Actions being added that are already logged do not result in error
            if ($mb.HasCustomActions -eq $true) {
                Write-Verbose -Message "Re-adding custom actions for $($mb.DisplayName)."
                #Build command based on logon types that need actions re-added
                $command = {Set-Mailbox -Identity $mb.DistinguishedName}
                if ($mb.HasCustomActionsAdmin -eq $true) {
                    $command = [ScriptBlock]::Create($command.ToString() + ' -AuditAdmin @{add=$mb.AuditAdmin}')
                }
                if ($mb.HasCustomActionsDelegate -eq $true) {
                    $command = [ScriptBlock]::Create($command.ToString() + ' -AuditDelegate @{add=$mb.AuditDelegate}')
                }
                if ($mb.HasCustomActionsOwner -eq $true) {
                    $command = [ScriptBlock]::Create($command.ToString() + ' -AuditOwner @{add=$mb.AuditOwner}')
                }
                #Write-Verbose -Message "Command that will be executed to re-add custom actions: $command"
                #Execute built command
                &$command
            }

            $RecommendedActionsLog.Add($mb.UserPrincipalName) | Out-Null
            $k++
            }
        Write-Progress -Activity "Adding recommended actions" -Status " " -Completed
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
