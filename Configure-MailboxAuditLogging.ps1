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
        are those enabled by default. (This phase can be included with the ConfigureAuditActions parameter.)

        You need to be connected with the Exchange Online module in order for this script to run.
        If the module is not installed, run Install-Module ExchangeOnlineManagement, then Connect-ExchangeOnline.

    .PARAMETER ConfigureAuditActions
        Switch to also perform the second phase of the script, which sets the recommended actions to be logged.
    .PARAMETER DoNotRetainCustomActions
        Switch to not keep any non-default actions that were enabled for a mailbox. In other words,
        only reset affected mailboxes to the default audit set. Note: Mailboxes that are not missing any default
        actions will not have any non-default actions disabled.
    .PARAMETER WhatIf
        Switch that causes the script to not make any changes, but still writes to the log file(s) the mailboxes
        that would be updated.
    
    .NOTES
        Version 2.4.0
        December 30, 2024
#>

#Requires -Module ExchangeOnlineManagement
[CmdletBinding(SupportsShouldProcess)]
Param (
    [switch]$ConfigureAuditActions,
    [switch]$DoNotRetainCustomActions
)

if (-not(Get-Command -Name Get-Mailbox -ErrorAction SilentlyContinue))
    {
    throw "Please connect to Exchange Online before running this script."
    }

$WhatIfPreference = $false
$today = Get-Date -Format yyyyMMdd

# Actions enabled by default
[System.Collections.ArrayList]$AuditOwnerActions = @('ApplyRecord','HardDelete','MailItemsAccessed','MoveToDeletedItems','Send','SoftDelete','Update','UpdateCalendarDelegation','UpdateFolderPermissions','UpdateInboxRules')
[System.Collections.ArrayList]$AuditDelegateActions = @('ApplyRecord','Create','HardDelete','MailItemsAccessed','MoveToDeletedItems','SendAs','SendOnBehalf','SoftDelete','Update','UpdateFolderPermissions','UpdateInboxRules')
[System.Collections.ArrayList]$AuditAdminActions = @('ApplyRecord','Create','HardDelete','MailItemsAccessed','MoveToDeletedItems','Send','SendAs','SendOnBehalf','SoftDelete','Update','UpdateCalendarDelegation','UpdateFolderPermissions','UpdateInboxRules')

# Function that determines if a mailbox's actions that are enabled are at least the
# ones that are recommended
function Compare-AuditActions ($mailbox) {
    $logonTypes = @('AuditOwner','AuditDelegate','AuditAdmin')

    foreach ($logonType in $logonTypes) {

        # Actions being logged for the logon type to compare against
        $diffObject = $mailbox.$logonType

        # Set baseline actions for the logon type to use as a reference
        switch ($logonType) {
            'AuditAdmin' {$refObject = $AuditAdminActions}
            'AuditDelegate' {$refObject = $AuditDelegateActions}
            'AuditOwner' {$refObject = $AuditOwnerActions}
        }

        $actionComparison = Compare-Object -ReferenceObject $refObject -DifferenceObject $diffObject
        $missingActions = $actionComparison | Where-Object {$_.SideIndicator -eq '<='}
        $extraActions = $actionComparison | Where-Object {$_.SideIndicator -eq '=>'}
        if ($missingActions) {
            # Mailbox is missing recommended actions
              New-Variable -Name $('missing'+$logonType) -Value ($missingActions | ForEach-Object {$_.InputObject})
            if ($extraActions) {              
                # Non-default actions are enabled
                New-Variable -Name $('nonDefault'+$logonType) -Value $true
                $nonDefault = $true
            }
        }
    }

    if ($missingActions) {
        return New-Object -TypeName psobject -Property @{
            IsMissingDefaultActions = $true
            MissingActionsAdmin = $missingAuditAdmin
            MissingActionsDelegate = $missingAuditDelegate
            MissingActionsOwner = $missingAuditOwner
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

    1. Those that do not have audit logging enabled, which applies only to resource mailboxes.
    2. Those that are enabled by global audit logging (aka on-by-default) but whose events
       are not being sent to the Unified Audit Log, which are the following:
       * Shared mailboxes
       * User mailboxes that do not have M365 Advanced Auditing service plan assigned and
         have not been explicitly enabled for audit logging

    Important note: Only a server-side filter can be used to determine #2.  The
    AuditEnabled property will always have a value of True, so it requires Exchange to determine
    mailboxes that are implicitly True because of global audit logging or explicitly True via
    the Set-Mailbox cmdlet.
#>
Write-Host "$(Get-Date) Getting mailboxes whose events are not being sent to the Unified Audit Log..." -ForegroundColor Green
[array]$nonUALMailboxes = Get-EXOMailbox -ResultSize:Unlimited -Filter `
    'AuditEnabled -ne $true -and ((RecipientTypeDetails -eq "UserMailbox" -and PersistedCapabilities -ne "M365Auditing") -or (RecipientTypeDetails -ne "UserMailbox"))'

Write-Host "$(Get-Date) $($nonUALMailboxes.Count) mailboxes were returned." -ForegroundColor Green
if ($nonUALMailboxes.Count -gt 0) {
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
        # Manually/explicitly enable audit logging for the mailbox
        if ($PSCmdlet.ShouldProcess($mb.UserPrincipalName,"Set-Mailbox -AuditEnabled `$true")) {
            Set-Mailbox -Identity $mb.DistinguishedName -AuditEnabled $true
        }
        $UALMailboxLog.Add($mb.UserPrincipalName) | Out-Null
        $i++
        }
    Write-Progress -Activity "Enabling audit logging" -Status " " -Completed
    $UALMailboxLog | Out-File -FilePath $UALMailboxLogFile -Append
    Write-Host "$(Get-Date) Mailboxes that have been modified are logged in $UALMailboxLogFile in the current directory." `
        -ForegroundColor Green
    }
Write-Host "$(Get-Date) Phase 1 is complete." -ForegroundColor Green

if ($ConfigureAuditActions) {
    Write-Host "$(Get-Date) Beginning phase 2..." -ForegroundColor Green
    
    # Get all mailboxes because client-side processing is necessary
    Write-Host "$(Get-Date) Getting all non-Group mailboxes. This may take some time..." `
        -ForegroundColor Green
    [array]$allMailboxes = Get-EXOMailbox -ResultSize Unlimited `
        -Properties DefaultAuditSet,AuditAdmin,AuditDelegate,AuditOwner,PersistedCapabilities
    
    Write-Host "$(Get-Date) $($allMailboxes.Count) mailboxes were returned." -ForegroundColor Green
    
    # Filter results to include those that don't have the recommended actions enabled for any of
    # Admin, Delegate, and Owner
    Write-Host "$(Get-Date) Filtering mailboxes that do not have the recommended actions enabled..." -ForegroundColor Green
    [System.Collections.ArrayList]$nonDefaultActionsMB =@()
    $j = 1
    foreach ($mb in $allMailboxes){
        Write-Progress -Activity "Determining mailboxes that do not have the recommended actions enabled" `
            -Status "Processing mailbox for $($mb.DisplayName)" -PercentComplete ($j/$allMailboxes.Count*100)

        # Check only mailboxes not using default audit set for any logon type
        if ($mb.DefaultAuditSet -join '' -ne 'AdminDelegateOwner') {
            
            $mbActionsCheck = Compare-AuditActions -mailbox $mb
            
            # Add mailbox to collection of those that need to be updated and whether extra actions need to be re-added
            if ($mbActionsCheck.IsMissingDefaultActions -eq $true) {
                $mbToAdd = New-Object -TypeName psobject -Property @{
                    DisplayName = $mb.DisplayName
                    DistinguishedName = $mb.DistinguishedName
                    UserPrincipalName = $mb.UserPrincipalName
                    HasCustomActions = $mbActionsCheck.HasCustomActions
                    HasCustomActionsAdmin = $mbActionsCheck.HasCustomActionsAdmin
                    HasCustomActionsDelegate = $mbActionsCheck.HasCustomActionsDelegate
                    HasCustomActionsOwner = $mbActionsCheck.HasCustomActionsOwner
                    MissingActionsAdmin = $mbActionsCheck.MissingActionsAdmin
                    MissingActionsDelegate = $mbActionsCheck.MissingActionsDelegate
                    MissingActionsOwner = $mbActionsCheck.MissingActionsOwner
                    AuditAdmin = $mb.AuditAdmin
                    AuditDelegate = $mb.AuditDelegate
                    AuditOwner = $mb.AuditOwner
                    PersistedCapabilities = $mb.PersistedCapabilities
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
        Write-Host "$(Get-Date) Enabling recommended actions..." `
            -ForegroundColor Green
        
        [System.Collections.ArrayList]$AuditActionsLog = @()
        if ($PSBoundParameters.ContainsKey('WhatIf')) {
            $fileTag = '-WhatIf'
        } 
        $AuditActionsLogFile = "MailboxAuditActionsUpdated$fileTag-$today.csv"
        $k = 1
        foreach ($mb in $nonDefaultActionsMB)
            {
            Write-Progress -Activity "Enabling recommended actions" -Status "Processing mailbox for $($mb.DisplayName)" `
                -PercentComplete ($k/$nonDefaultActionsMB.Count*100)

            # Configure the mailbox to use the default audit set for all logon types
            Write-Verbose -Message "Enabling recommended actions for $($mb.DisplayName)."
            if ($PSCmdlet.ShouldProcess($mb.UserPrincipalName,"Set-Mailbox -DefaultAuditSet")) {
                Set-Mailbox -Identity $mb.DistinguishedName -DefaultAuditSet Admin,Delegate,Owner
            }

            if ($DoNotRetainCustomActions -eq $false) {
                # If additional actions were being logged, add them back
                # Actions being added that are already logged do not result in error
                if ($mb.HasCustomActions -eq $true) {
                    Write-Verbose -Message "Re-adding non-default actions for $($mb.DisplayName)."
                    # Build command based on logon types that need actions re-added
                    $command = {Set-Mailbox -Identity $mb.DistinguishedName}
                    if ($mb.HasCustomActionsAdmin -eq $true) {
                        # MessageBind cannot be added if Purview Audit (Premium) service plan assigned. MailItemsAccessed
                        # is to be used instead, which is enabled by default. Therefore, to avoid error, remove MessageBind
                        # if enabled and the service plan is assigned
                        if ($mb.PersistedCapabilities -contains 'M365Auditing' -and $mb.AuditAdmin -contains 'MessageBind') {
                            Write-Verbose "Removing MessageBind as a logged action to re-add for AuditAdmin for $($mb.DisplayName) because the mailbox has the Purview Audit (Premium) service plan enabled."
                            $mb.AuditAdmin = $mb.AuditAdmin | Where-Object {$_ -ne 'MessageBind'}
                            # Only add AuditAdmin if additional custom actions besides MessageBind were enabled (to avoid error)
                            if (Compare-Object -ReferenceObject $AuditAdminActions -DifferenceObject $mb.AuditAdmin  | Where-Object {$_.SideIndicator -eq '=>'}) {
                                $command = [ScriptBlock]::Create($command.ToString() + ' -AuditAdmin @{add=$mb.AuditAdmin}')
                            }
                            else {
                                $mb.HasCustomActionsAdmin = $false
                            }
                        }
                        else {
                            $command = [ScriptBlock]::Create($command.ToString() + ' -AuditAdmin @{add=$mb.AuditAdmin}')   
                        }  
                    }
                    if ($mb.HasCustomActionsDelegate -eq $true) {
                        $command = [ScriptBlock]::Create($command.ToString() + ' -AuditDelegate @{add=$mb.AuditDelegate}')
                    }
                    if ($mb.HasCustomActionsOwner -eq $true) {
                        $command = [ScriptBlock]::Create($command.ToString() + ' -AuditOwner @{add=$mb.AuditOwner}')
                    }
                    # Determine if command still needs to be run (because MessageBind may have been removed above)
                    if ($mb.HasCustomActionsAdmin -eq $true -or $mb.HasCustomActionsDelegate -eq $true -or $mb.HasCustomActionsOwner -eq $true) {
                        # Execute built command only if WhatIf not used
                        Write-Verbose -Message "Command that will be executed to re-add non-default actions: $command"
                        if (-not $PSBoundParameters.ContainsKey('WhatIf')) {
                            &$command
                        }
                    }
                    else {
                        Write-Verbose "Skipping re-adding non-default actions for $($mb.DisplayName) because there are none to add (after removing MessageBind for AuditAdmin, which is not applicable)."
                    }
                }
            }

            $mbUpdated = [pscustomobject] [ordered] @{
                Mailbox = $mb.UserPrincipalName
                AddedAuditAdmin = $mb.MissingActionsAdmin -join ' '
                AddedAuditDelegate = $mb.MissingActionsDelegate -join ' '
                AddedAuditOwner = $mb.MissingActionsOwner -join ' '
            }
            $AuditActionsLog.Add($mbUpdated) | Out-Null
            $k++
            }
        Write-Progress -Activity "Enabling recommended actions" -Status " " -Completed
        #$AuditActionsLog | Select-Object -Property Mailbox,AddedAuditAdmin,AddedAuditDelegate,AddedAuditOwner | Export-CSV -Path $AuditActionsLogFile -NoTypeInformation -Append
        $AuditActionsLog | Export-CSV -Path $AuditActionsLogFile -NoTypeInformation -Append
        Write-Host "$(Get-Date) Log file $AuditActionsLogFile is in the current directory." -ForegroundColor Green
        }
    Write-Host "$(Get-Date) Phase 2 is complete." -ForegroundColor Green
    }
else
    {
    Write-Host "$(Get-Date) Phase 2 has been skipped because the ConfigureAuditActions switch was not used." `
        -ForegroundColor Green
    }
