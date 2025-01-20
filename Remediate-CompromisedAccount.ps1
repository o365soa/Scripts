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
#Requires -Modules @{ModuleName='Microsoft.Graph.Authentication';ModuleVersion='2.0.0'},ExchangeOnlineManagement

<#
	.SYNOPSIS
		Remediate an account which has been successfully compromised.

	.DESCRIPTION
		This script will remediate an account which has had credentials breached. Commonly, an attacker will extend
        their access by sharing their data. Simply resetting the users password is not enough to prevent this.

        The following actions are performed
            1. Reset user's password (unless synced user without password writeback enabled)
            2. Enable per-user multi-factor authentication
            3. Revoke (invalidate) all refresh tokens, forcing user to re-authenticate to all applications
            4. Disable email forwarding rules
            5. Disable anonymous calendar sharing
            6. Remove mailbox delegates
            7. Remove mailbox forwarding configuration

        Actions can be disabled using No* parameters. For example, to not enable MFA, use -NoMFA.

        Forensic information is exported unless the -NoForensics parameter is used. This information contains
        details about the mailbox, inbox rules, delegates, calendar sharing, and auditing information of the user prior
        to the remediation actions. This can be useful for further investigation or potential reversal of any of the
        actions performed.

        Parts of this script and some actions have been taken from Brandon Koeller's script
        https://github.com/OfficeDev/O365-InvestigationTooling/blob/master/RemediateBreachedAccount.ps1

        For more information on this script, please visit
        https://github.com/o365soa/Scripts/

    .PARAMETER NoForensics
    .PARAMETER NoPasswordReset
    .PARAMETER NoMFA
    .PARAMETER NoDisableForwardingRules
    .PARAMETER NoRevokeRefreshToken
    .PARAMETER NoRemoveCalendarPublishing
    .PARAMETER NoRemoveDelegates
    .PARAMETER NoRemoveMailboxForwarding

    .PARAMETER CloudEnvironment
        Cloud instance of the tenant. Possible values are Commercial, USGovGCC, USGovGCCHigh, USGovDoD, and China.
        Default value is Commercial.

    .PARAMETER  ConfirmAll
        Specifying this parameter will automate the remediation process, by default, confirmation is required.
        WARNING: This does not allow you to confirm the mailbox before remediation

	.EXAMPLE
		.\Remediate-CompromisedAccount.ps1 -UPN joe@contoso.com

	.EXAMPLE
		.\Remediate-CompromisedAccount.ps1 -UPN joe@contoso.com -NoMFA

	.NOTES
        Version 2.0
		January 8, 2025

	.LINK
		about_functions_advanced

#>

Param(
    [CmdletBinding()]
    [Parameter(Mandatory=$True)]
    [String]$UPN,
    [switch]$NoForensics,
    [Switch]$NoPasswordReset,
    [switch]$NoMFA,
    [switch]$NoDisableForwardingRules,
    [switch]$NoRevokeRefreshTokens,
    [switch]$NoDisableCalendarPublishing,
    [switch]$NoRemoveDelegates,
    [switch]$NoRemoveMailboxForwarding,
    [switch]$NoDisableMobileDevices,
    [switch]$ConfirmAll,
    [ValidateSet("Commercial", "USGovGCC", "USGovGCCHigh", "USGovDoD", "China")][string]$CloudEnvironment="Commercial"
)

#region Functions

Function Reset-Password {
    # Function reset's the user password with a random password
    # Requires delegated scope UserAuthenticationMethod.ReadWrite.All and at least Entra role Authentication Administrator
	Param(
		[string]$UPN
	)
    Write-Host "[$UPN] Resetting password..."

    $Password = [System.Web.Security.Membership]::GeneratePassword(12,2)
    $headers = @{
        'Content-Type' = 'application/json'
    }
    $body = @{
        newPassword = $Password
    }
    # The user is forced to change their password after logging in with the new password
    Invoke-MgGraphRequest -Method POST -Uri "/v1.0/users/$UPN/authentication/methods/28c10230-6103-485e-b985-444c60001490/resetPassword" -Body ($body | ConvertTo-Json) -Headers $headers -ResponseHeadersVariable rHeaders -StatusCodeVariable $rCode
    if ($?) {
        if ($rCode -eq 202) {
            return $Password
        }
        if ($rHeaders.Location) {
            $locationResponse = Invoke-MgGraphRequest -Method GET -Uri $rHeaders.Location[0]
            if ($locationResponse.status -eq 'succeeded' -or $locationResponse.status -eq 'running') {
                return $Password
            } else {
                Write-Warning -Message "Error occurred resetting password for $UPN. Reset response status: $($locationResponse.status)"
                return $null
            }
        }
    } else {
        Write-Error "Failed to reset password for $UPN."
    }
}

Function Enable-MFA {
    # Turns on per-user MFA for the user
    # Requires delegated scope Policy.ReadWrite.AuthenticationMethod and at least Entra role Authentication Policy Administrator
    Param(
        [string]$UPN
    )

    Write-Host "[$UPN] Enabling MFA..."

    $headers = @{
        'Content-Type' = 'application/json'
    }
    $body = @{
        perUserMfaState = 'enforced'
    }

    Invoke-MgGraphRequest -Method PATCH -Uri "/beta/users/$UPN/authentication/requirements" -Body ($body | ConvertTo-Json) -Headers $headers -StatusCodeVariable rCode
    if ($rCode -eq 204) {
        return $true
    } else {
        Write-Error "Failed to enable MFA for $UPN"
    }
    
}

Function Disable-ForwardingRules {
    # Disable forwarding rules
	Param(
		[string]$UPN
	)
    Write-Host "[$UPN] Disabling email forwarding rules..."
    
    if($ConfirmAll) { $Confirmation = $false; } else { $Confirmation = $true; }
    $rules = Get-InboxRule -Mailbox $upn | Where-Object {(($_.Enabled -eq $true) -and (($null -ne $_.ForwardTo) -or ($null -ne $_.ForwardAsAttachmentTo) -or ($null -ne $_.RedirectTo)))}
    if ($rules) {
        try {
            $rules | Disable-InboxRule -Confirm:$Confirmation -ErrorAction Stop
            return 1
        } catch {
            return 0
        }
    } else {
        return -1
    }
}

Function Export-Forensics {
    # This script exports current settings about the user which can be used for forensics information later
	Param(
		[string]$UPN,
        [string]$MailboxIdentity
	)
    
    $ForensicsFolder = "$PSScriptRoot\Forensics\$UPN\"

    Write-Host "[$UPN] Exporting forensics to $ForensicsFolder..."
    if(!(Test-Path($ForensicsFolder))) { 
        try {
            mkdir $ForensicsFolder -ErrorAction:Stop | Out-Null 
        } catch {
            Write-Error "Cannot create directory $ForensicsFolder"
            exit
        }
    }

    Get-Mailbox -Identity $UPN | ConvertTo-Json -Depth 20 | Out-File -FilePath "$ForensicsFolder\$UPN-mailbox.json"
    Get-InboxRule -Mailbox $UPN | ConvertTo-Json -Depth 20 | Out-File -FilePath "$ForensicsFolder\$UPN-inboxrules.json"
    Get-MailboxCalendarFolder -Identity "$($MailboxIdentity):\Calendar" | ConvertTo-Json -Depth 20 | Out-File -FilePath "$ForensicsFolder\$UPN-MailboxCalendarFolder.json"
    Get-MailboxPermission -Identity $UPN | Where-Object {($_.IsInherited -ne "True") -and ($_.User -notlike "*SELF*")} | ConvertTo-Json -Depth 20 | Out-File -FilePath "$ForensicsFolder\$UPN-MailboxDelegates.json"
    Get-MobileDevice -Mailbox $UPN | ConvertTo-Json -Depth 20 | Out-File -FilePath "$ForensicsFolder\$UPN-devices.json"
    Get-MobileDevice -Mailbox $UPN | Get-MobileDeviceStatistics | ConvertTo-Json -Depth 20 | Out-File -FilePath "$ForensicsFolder\$UPN-devicestatistics.json"
    
    $startDate = (Get-Date).AddDays(-7).ToString('MM/dd/yyyy') 
    $endDate = (Get-Date).ToString('MM/dd/yyyy')

    $rand = Get-Random -Minimum 100 -Maximum 999
    $recordCount = 0
    do {
        Write-Verbose -Message "Executing paged search of audit log (Current total: $recordCount)..."
        [array]$results = Search-UnifiedAuditLog -StartDate $startDate -EndDate $endDate -UserIds $UPN -Formatted -ResultSize 5000 -SessionCommand ReturnLargeSet -SessionId $UPN$rand
        if ($results.Count -gt 0) {
            $recordCount += $results.Count
            $results | ConvertTo-Json -Depth 25 | Out-File -FilePath "$ForensicsFolder\$UPN-AuditLog.json" -Append
        }
    } until ($results.Count -eq 0)
    if ($recordCount -eq 50000) {
        Write-Warning -Message "Audit log search returned the maxiumum 50000 records. Not all activities may have been returned."
        $script:Notes += "`nAudit log search returned 50000 records, which is the limit for a single paged query. Not all activities may have been returned."
    }

}

Function Revoke-RefreshToken {
    # Revokes Refresh Token for User, forcing logged in applications to re-logon.
    # Requires delegated scope User.RevokeSessions.All
	Param(
		[string]$UPN
	)
    
    Write-Host "[$UPN] Invalidating refresh tokens..."
    $response = Invoke-MgGraphRequest -Method POST -Uri "/v1.0/users/$UPN/revokeSignInSessions"
    if ($response.value -ne $true) {
        Write-Error "Failed to invalidate refresh tokens for $UPN"
    } else {
        return $true
    }
}

Function Remove-CalendarPublishing {
    # Removes anonymous calendar publishing for the user
	Param(
        [string]$UPN,
		[string]$MailboxIdentity
	)
    
    Write-Host "[$UPN] Disabling anonymous calendar publishing..."

    # Check setting first
    if ((Get-MailboxCalendarFolder -Identity "$($MailboxIdentity):\Calendar").PublishEnabled -eq $true) {
        try {
            Set-MailboxCalendarFolder -Identity "$($MailboxIdentity):\Calendar" -PublishEnabled:$false -ErrorAction Stop
            return 1
        } catch {
            return 0
        }
    } else {return -1}
}

Function Remove-MailboxDelegates {
    # Removes Mailbox Delegates from Mailbox where not SELF
	Param(
        [string]$UPN
	)
    
    Write-Host "[$UPN] Removing mailbox delegates..."
    $mailboxDelegates = Get-MailboxPermission -Identity $upn | Where-Object {$_.IsInherited -ne "True" -and $_.User -notlike "*SELF*"}
    if ($mailboxDelegates) {
        foreach ($delegate in $mailboxDelegates) {
            try {
                Remove-MailboxPermission -Identity $upn -User $delegate.User -AccessRights $delegate.AccessRights -InheritanceType All -Confirm:$false -ErrorAction Stop
            } catch {
                Write-Error -Message "Failed to remove delegate $($delegate.User) from mailbox $UPN"
                $fail = $true
            }
        }
        if ($fail -eq $true) {return 0} else {return 1}
    } else {return -1}

}

Function Remove-MailboxForwarding {
    # Removes Mailbox Forwarding Configuration from Mailbox
	Param(
        [string]$UPN
	)

    Write-Host "[$UPN] Removing mailbox forwarding..."
    $mb = Get-Mailbox -Identity $UPN
    if ($mb.DeliverToMailboxAndForward -or $mb.ForwardingSmtpAddress) {
        try {
            Set-Mailbox -Identity $upn -DeliverToMailboxAndForward $false -ForwardingSmtpAddress $null -WarningAction SilentlyContinue -ErrorAction Stop
            return 1
        } catch {
            return 0
        }
    } else {return -1}

}

Function Disable-MobileDevices {
    # Disable Mobile Devices for the User
	Param(
		[string]$UPN
	)
    Write-Host "[$UPN] Disabling ActiveSync devices..."

    [array]$MobileDevices = Get-MobileDevice -Mailbox $UPN

    $DisableDevices = @()
    $WipeDevices = @()

    if ($MobileDevices.Count -gt 0) {
        foreach($MobileDevice in $MobileDevices) {
            $Stats = $null
            $Stats = $MobileDevice | Get-MobileDeviceStatistics

            if(!$ConfirmAll) {
                $options = [System.Management.Automation.Host.ChoiceDescription[]] @("&Block", "&Wipe","&Allow")
                $answer = $host.UI.PromptForChoice($null , "`nBlock Mobile Device $($Stats.DeviceUserAgent) First Sync Time $($Stats.FirstSyncTime) Last Sync Attempt Time $($Stats.LastSyncAttemptTime)?" , $Options,0)
                if($answer -eq 0) { $DisableDevices += $MobileDevice.DeviceId }
                if($answer -eq 1) { $WipeDevices += $MobileDevice }
            } else {
                $DisableDevices += $MobileDevice.DeviceId
            }
        }

        if ($DisableDevices) {
            try {
                Set-CASMailbox $UPN -ActiveSyncBlockedDeviceIDs $DisableDevices -ErrorAction Stop
            } catch {
                Write-Error -Message "Failed to disabled devices for $UPN"
                $fail = $true
            }
        }

        foreach($WipeDevice in $WipeDevices) {
            Write-Host "[$UPN] Wiping device $($WipeDevice.Identity)..."
            try {
                Clear-MobileDevice -Identity "$($WipeDevice.Identity)" -ErrorAction Stop
            } catch {
                Write-Error -Message "Failed to issue wipe device command for $($WipeDevice.Identity)"
                $fail = $true
            }
        }
        if ($fail -eq $true) {return 0} else {return 1}
    }
    else {return -1}

}

#endregion functions

#region start

Start-Transcript -Path "$PSScriptRoot\Remediate-$UPN.txt"
$Notes = ""

#endregion start

#region prechecks
$requiredScopes = @()
if ($NoPasswordReset -eq $false) {
    $requiredScopes += 'UserAuthenticationMethod.ReadWrite.All','Domain.Read.All'
}
if ($NoMFA -eq $false) {
    $requiredScopes += 'Policy.ReadWrite.AuthenticationMethod'
}
if ($NoRevokeRefreshToken -eq $false) {
    $requiredScopes += 'User.RevokeSessions.All'
}
if ($requiredScopes) {
    $currentScopes = (Get-MgContext).Scopes
    if ($currentScopes) {
        foreach ($scope in $requiredScopes) {
            if ($currentScopes -notcontains $scope) {
                $scopeNeeded = $true
                break
            }
        }
    }
    if ($scopeNeeded -or -not $currentScopes) {
        switch ($CloudEnvironment) {
            "Commercial"   {$cloud = "Global"}
            "USGovGCC"     {$cloud = "Global"}
            "USGovGCCHigh" {$cloud = "USGov"}
            "USGovDoD"     {$cloud = "USGovDoD"}
            "China"        {$cloud = "China"}            
        }
        Write-Host "Connecting to Microsoft Graph..."
        Connect-MgGraph -ContextScope Process -Scopes $requiredScopes -Environment $cloud -NoWelcome
    }
}
if (-not (Get-Command -Name Set-Mailbox -ErrorAction SilentlyContinue)) {
    switch ($CloudEnvironment) {
        "Commercial"   {Connect-ExchangeOnline -ShowBanner:$false -WarningAction SilentlyContinue | Out-Null}
        "USGovGCC"     {Connect-ExchangeOnline -ShowBanner:$false -WarningAction SilentlyContinue | Out-Null}
        "USGovGCCHigh" {Connect-ExchangeOnline -ExchangeEnvironmentName O365USGovGCCHigh -ShowBanner:$false -WarningAction SilentlyContinue | Out-Null}
        "USGovDoD"     {Connect-ExchangeOnline -ExchangeEnvironmentName O365USGovDoD -ShowBanner:$false -WarningAction SilentlyContinue | Out-Null}
        "China"        {Connect-ExchangeOnline -ExchangeEnvironmentName O365China -ShowBanner:$false -WarningAction SilentlyContinue | Out-Null}
    }
}

$Mailbox = Get-Mailbox $UPN -ErrorAction:Stop

if (-not $Mailbox) {
    Write-Error "Failed to get mailbox for $UPN."
    exit
}

Write-Host "[$UPN] Mailbox Identity: $($Mailbox.Identity), Display Name: $($Mailbox.DisplayName)"
 
if(!$ConfirmAll) {
    # Perform confirmation of the mailbox before continuing
    $options = [System.Management.Automation.Host.ChoiceDescription[]] @("&Remediate", "&Quit")
    $result = $host.UI.PromptForChoice($null , "`nConfirm Account?" , $Options,1)
    if($result -eq 1) { exit }
}

if(!$NoPasswordReset) {
    # Determine if user is a federated user, turn off password reset if it is federated and notify that we must set on-premises
    $Domain = $UPN.Split("@")[1]
    if((Invoke-MgGraphRequest -Method GET -Uri "/v1.0/domains/$Domain").authenticationType -ne 'Managed') {
        Write-Host "User's sign-in domain is a federated domain, so the password will need to be managed on-premises unless password-write back is enabled."
        $Notes += "`nUser domain is federated. If password write-back is not enabled, password should be set on-premises."
    }
}

#endregion prechecks

#region remediation

# Remediation actions
if(!$NoForensics) {
    Write-Host "[$UPN] Exporting logs for forensics..."
    Export-Forensics -UPN $UPN -MailboxIdentity $Mailbox.Identity
}
Write-Host "[$UPN] Remediation beginning..."

if(!$NoPasswordReset) {$NewPassword = Reset-Password -UPN $UPN}
if(!$NoMFA) {$MFAResult = Enable-MFA -UPN $UPN}
if(!$NoRevokeRefreshTokens) {$RevokeResult = Revoke-RefreshToken -UPN $UPN}
if(!$NoDisableForwardingRules) {$ForwardingRulesResult = Disable-ForwardingRules -UPN $UPN}
if(!$NoDisabledCalendarPublishing) {$CalPublishingResult = Remove-CalendarPublishing -UPN $UPN -MailboxIdentity $Mailbox.Identity}
if(!$NoRemoveDelegates) {$DelegatesResult = Remove-MailboxDelegates -UPN $UPN}
if(!$NoRemoveMailboxForwarding) {$ForwardingResult = Remove-MailboxForwarding -UPN $UPN}
if(!$NoDisableMobileDevices) {$MobileDevicesResult = Disable-MobileDevices -UPN $UPN}

#endregion remediation

#region report

Write-Host "`n`nRemediation report for $UPN" -ForegroundColor Green
if(!$NoPasswordReset) {
    if (-not $NewPassword){
        Write-Host "New Password: $NewPassword"
    }
}

Write-Host "`nActions performed"
if (-not $NoForensics) { Write-Host " - Forensic information exported" }
if (-not $NoPasswordReset) {
    Write-Host " - Password Reset: " -NoNewline
    if ($NewPassword) {Write-Host "Success" -ForegroundColor Green} else {Write-Host "Failed" -ForegroundColor Red}
}
if (-not $NoMFA) {
    Write-Host " - Enable Per-User MFA: " -NoNewline
    if ($MFAResult) {Write-Host "Success" -ForegroundColor Green} else {Write-Host "Failed" -ForegroundColor Red}
}
if (-not $NoRevokeRefreshTokens) {
    Write-Host " - Revoke Refresh Tokens: " -NoNewline
    if ($RevokeResult) {Write-Host "Success" -ForegroundColor Green} else {Write-Host "Failed" -ForegroundColor Red}
}
if (-not $NoDisableForwardingRules) {
    Write-Host " - Disable Forwarding Rules: " -NoNewline
    if ($ForwardingRulesResult -eq 1) {
        Write-Host "Success" -ForegroundColor Green
    } elseif ($ForwardingRulesResult -eq 0) {
        Write-Host "Failed" -ForegroundColor Red
    } else {
        Write-Host "No rules found" -ForegroundColor Gray
    }
}
if (-not $NoRemoveCalendarPublishing) {
    Write-Host " - Remove Calendar Publishing: " -NoNewline
    if ($CalPublishingResult -eq 1) {
        Write-Host "Success" -ForegroundColor Green
    } elseif ($CalPublishingResult -eq 0) {
        Write-Host "Failed" -ForegroundColor Red
    } else {
        Write-Host "Publishing was not enabled" -ForegroundColor Gray
    }
}
if (-not $NoRemoveDelegates) {
    Write-Host " - Remove Mailbox Delegates: " -NoNewline
    if ($DelegatesResult -eq 1) {
        Write-Host "Success" -ForegroundColor Green
    } elseif ($DelegatesResult -eq 0) {
        Write-Host "Failed for 1 or more delegates" -ForegroundColor Red
    } else {
        Write-Host "No delegates found" -ForegroundColor Gray
    }
}
if (-not $NoRemoveMailboxForwarding) {
    Write-Host " - Remove Mailbox Forwarding: " -NoNewline
    if ($ForwardingResult -eq 1) {
        Write-Host "Success" -ForegroundColor Green
    } elseif ($ForwardingResult -eq 0) {
        Write-Host "Failed" -ForegroundColor Red
    } else {
        Write-Host "No forwarding configuration found" -ForegroundColor Gray
    }
}
if (-not $NoDisableMobileDevices) {
    Write-Host " - Disable Mobile Devices: " -NoNewline
    if ($MobileDevicesResult -eq 1) {
        Write-Host "Success" -ForegroundColor Green
    } elseif ($MobileDevicesResult -eq 0) {
        Write-Host "Failed for 1 or more devices" -ForegroundColor Red
    } else {
        Write-Host "No mobile devices found" -ForegroundColor Gray
    }
}

Write-Host "`nAdditional notes" -ForegroundColor Green
Write-Host $Notes

#endregion report

Stop-Transcript
