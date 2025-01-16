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
#Requires -Modules @{ModuleName='Microsoft.Graph.Authentication';ModuleVersion='2.0.0'}

<#
	.SYNOPSIS
		Creates a report in either CSV or HTML format of all MFA enrollment.
	.DESCRIPTION
        This script generates a report in either CSV or HTML format for all MFA enrollment

        HTML format uses the opensource bootstrap libraries for a tidy interface.
    .PARAMETER Output
        The name (including optional path) of the file for the report. If not specified, the default is mfa-report.csv or mfa-report.html, depending on the output format.
    .PARAMETER CloudEnvironment
        The cloud instance that hosts the tenant. Used to set the endpoints for authentication and connection.
        Value can be Commercial, USGovGCC, USGovGCCHigh, USGovDoD, China. Default is Commercial.
    .PARAMETER IncludePerUserState
        Include the per-user MFA state in the report.
    .PARAMETER IncludeDisabledUsers
        Include disabled users in the report.
    .PARAMETER IncludeGuests
        Include guest users in the report.
    .PARAMETER IgnoreRoles
        An array of roles, by display name, to not include in the report. Default is Guest Inviter, Partner Tier1 Support,
        Partner Tier2 Support, Directory Readers, Directory Synchronization Accounts, Device Users, Device Join, Workplace Device Join.
	.PARAMETER ExcludeRoleBreakdown
        Do not include a breakdown of Entra roles that have users in the HTML report.
    .EXAMPLE
		.\Get-MFAReport.ps1 -CSV
    .EXAMPLE
        .\Get-MFAReport.ps1 -HTML -IncludeGuests
	.NOTES
		Version 2.0
		January 16, 2025
	.LINK
		about_functions_advanced

#>

param(
  [string]$Output,
  [array]$IgnoreRoles=@("Guest Inviter","Partner Tier1 Support","Partner Tier2 Support","Directory Readers","Directory Synchronization Accounts","Device Users","Device Join","Workplace Device Join"),
  [ValidateSet("Commercial", "USGovGCC", "USGovGCCHigh", "USGovDoD", "China")][string]$CloudEnvironment="Commercial",
  [switch]$IncludePerUserState,
  [switch]$IncludeDisabledUsers,
  [switch]$IncludeGuests,
  [switch]$ExcludeRoleBreakdown,
  [Parameter(ParameterSetName='CSVOutput')]
    [switch]$CSV,
  [Parameter(ParameterSetName='HTMLOutput')]
    [switch]$HTML
)

 # Build a default output file if none specified
 if (!$output) {
    if ($CSV) {
        $Output = "mfa-report.csv"
    }
    if ($HTML) {
        $Output = "mfa-report.html"
    }
}

# Directory.Read.All = Least common scope for Users and DirectoryObjects APIs
# UserAuthenticationMethod.Read.All = Scope for Authentication Methods, Sign-In Preferences, and System-Preferred MFA Method APIs. Must also have Entra role of either Global Reader or Authentication Administrator or Privileged Authentication Administrator
# Policy.Read.All = Scope for Authentication Requirements API. Must also have Entra role of either Global Reader or Authentication Policy Administrator
# RoleManagement.Read.Directory = Scope for Role Definitions and Role Assignments APIs. Must also have Entra role of User Administrator or higher

$requiredScopes = @('Directory.Read.All','UserAuthenticationMethod.Read.All','Policy.Read.All')
if (-not $ExcludeRoleBreakdown) {
    $requiredScopes += 'RoleManagement.Read.Directory'
}
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
    Write-Host -ForegroundColor Green "$(Get-Date) Connecting to Microsoft Graph..."
    Connect-MgGraph -ContextScope Process -Scopes $requiredScopes -Environment $cloud -NoWelcome
}
 <# 
 
    User Checking

    The following goes through all users in Get-MsolUser, and then parses their StrongAuth settings.
 
 #>

Write-Host "$(Get-Date) Getting all users. This may take some time..."
$result = New-Object -TypeName System.Collections.ArrayList
$apiUrl = "/v1.0/users?`$select=id,userPrincipalName,accountEnabled,userType"
do {
    Write-Progress -Id 1 -Activity "Getting Users..."
    $response = Invoke-MgGraphRequest -Method GET -Uri $apiUrl -OutputType PSObject
    $apiUrl = $response."@odata.nextLink"
    if ($apiUrl) { Write-Verbose "@odata.nextLink: $ApiUrl" }
    $result.AddRange($response.Value) | Out-Null
} until ($null -eq $response."@odata.nextLink" )

Write-Host "$(Get-Date) Getting strong authentication settings for all users. This may take some time..."
$mfaUsers = New-Object -TypeName System.Collections.ArrayList
$mfaUserIds = New-Object -TypeName System.Collections.ArrayList
$i = 0
$culture = [System.Globalization.CultureInfo]::CurrentCulture
$textInfo = $culture.TextInfo
foreach ($user in $result) {
    $i++
    if ($IncludeDisabledUsers -eq $false -and $user.accountEnabled -eq $false) {
        continue
    }
    if ($IncludeGuests -eq $false -and $user.userType -eq "Guest") {
        continue
    }
    Write-Progress -Id 1 -Activity "Processing Users for Strong Auth Details..." -Status $($user.userPrincipalName) -PercentComplete ($i / $result.Count * 100) 

    # Reset variables
    $defaultMethod = $null; $mfaState = $null; $mfaMobilePhone = $null; $phoneAppDevice = $null; $fidoDevice = $null; $softwareOTP = $null, $helloDevice = $null

    # Get the default MFA option
    $signInPreferences = Invoke-MgGraphRequest -Method GET -Uri "/beta/users/$($user.id)/authentication/signInPreferences"
    if ($signInPreferences.isSystemPreferredAuthenticationMethodEnabled -eq $true) {
        $defaultMethod = $signInPreferences.systemPreferredAuthenticationMethod
    } else {
        $defaultMethod = $signInPreferences.userPreferredMethodForSecondaryAuthentication
    }

    # Get per-user MFA state
    if ($IncludePerUserState) {
        $authRequirements = Invoke-MgGraphRequest -Method GET -Uri "/beta/users/$($user.id)/authentication/requirements"
        $mfaState = $textInfo.ToTitleCase($authRequirements.perUserMfaState)
    }

    # Get user's registered auth methods
    $authMethods = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/users/$($user.id)/authentication/methods" -OutputType PSObject

    # MFA Phone Number for mobile phone
    $mfaMobilePhone = ($authMethods.value | Where-Object {$_.id -eq '3179e48a-750b-4051-897c-87b9720928f7'}).phoneNumber
    # MFA Phone Number for office phone
    #$mfaOfficePhone = ($authMethods.value | Where-Object {$_.id -eq 'e37fc753-ff3b-4958-9484-eaa9425c82bc'}).phoneNumber
    # MFA Phone Number for alternate phone
    #$mfaAltPhone = ($authMethods.value | Where-Object {$_.id -eq 'b6332ec1-7057-4abe-9331-3d72feddfe417'}).phoneNumber

    # Registered device name(s) for MS Authenticator
    $phoneAppDevice = ($authMethods.value | Where-Object {$_."@odata.type" -eq "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod"}).displayName -join ' | '

    # Registered FIDO2 device name(s)
    $fidoDevice = @()
    foreach ($device in ($authMethods.value | Where-Object {$_."@odata.type" -eq "#microsoft.graph.fido2AuthenticationMethod"})) {
        $fidoDevice += "$device.displayName ($($device.model))"
    }

    # Registered 3P software OTP devices
    $softwareOTP = ($authMethods.value | Where-Object {$_."@odata.type" -eq "#microsoft.graph.softwareOathAuthenticationMethod"}).id -join ' | '

    # Registered Hello for Business devices
    $helloDevice = ($authMethods.value | Where-Object {$_."@odata.type" -eq "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod"}).displayName -join ' | '

    # Add to collection
    $userDetails = New-Object -TypeName psobject -Property @{
		UserPrincipalName=$user.UserPrincipalName
        MobilePhone=$mfaMobilePhone
        PhoneAppDevice=$phoneAppDevice
        FIDO2Device=$fidoDevice -join ' | '
        SoftwareOATH=$softwareOTP
        HelloForBusiness=$helloDevice
        Default=$defaultMethod
        Id=$user.id
    }
    if ($IncludePerUserState) {
        $userDetails | Add-Member -MemberType NoteProperty -Name MFAState -Value $mfaState
    }
    if ($IncludeDisabledUsers) {
        $userDetails | Add-Member -MemberType NoteProperty -Name AccountState -Value $(if($user.accountEnabled){"Enabled"}else{"Disabled"})
    }
    if ($IncludeGuests) {
        $userDetails | Add-Member -MemberType NoteProperty -Name UserType -Value $user.userType
    }
    $mfaUsers.Add($userDetails) | Out-Null
    $mfaUserIds.Add($user.id) | Out-Null
}
if (-not $ExcludeRoleBreakdown) {
    $roleUsers = New-Object -TypeName System.Collections.ArrayList

    Write-Host "$(Get-Date) Getting Entra Role Member Status. This may take some time..."
    $mRoles = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/roleManagement/directory/roleDefinitions?`$select=id,displayName" -OutputType PSObject
    $rolesToProcess = $mRoles.value | Where-Object {$_.displayName -notin $IgnoreRoles} | Sort-Object -Property displayName
    $i = 0
    foreach ($role in $rolesToProcess) {
        $i++
        Write-Progress -Id 1 -Activity "Parsing Role Members" -Status $($role.displayName)

        # Get active user members of the role, accounting for group assignments
        $mRoleMembers = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/roleManagement/directory/roleAssignments?`$filter=roleDefinitionId eq '$($role.id)'" -OutputType PSObject
        $memberToProcess = @()
        foreach ($member in $mRoleMembers.value) {
            # Get Entra ID object to determine its type
            $dirObject = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/directoryObjects/$($member.principalId)?`$select=id" -OutputType PSObject
            if ($dirObject."@odata.type" -eq "#microsoft.graph.group") {
                # v1.0 endpoint does not return service princiapls, but they are not relevant for this script
                $mgm = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/groups/$($member.principalId)/transitiveMembers?`$top=999&`$select=id,accountEnabled,userType" -OutputType PSObject
                # The parent group of nested members is also returned, so filter out the group
                foreach ($mMember in ($mgm.value | Where-Object {$_."@odata.type" -ne "#microsoft.graph.group"})) {
                    if ($IncludeDisabledUsers -eq $false -and $mMember.accountEnabled -eq $false) {
                        continue
                    }
                    if ($IncludeGuests -eq $false -and $mMember.userType -eq "Guest") {
                        continue
                    }
                    # Include member only if not already assigned role directly or from another group
                    if ($mMember.id -notin $memberToProcess) {
                        $memberToProcess += $mMember.id
                    }
                }
            } elseif ($dirObject."@odata.type" -eq "#microsoft.graph.user" -and $dirObject.id -notin $memberToProcess) {
                $memberToProcess += $dirObject.id
            }
        }
        # Loop through each and add settings
        foreach ($roleMember in $memberToProcess) {
            $rUser = $null
            # Get user details from collection based on its position (faster than using Where-Object with large collections)
            try {
                $user = $mfaUsers[$mfaUserIds.IndexOf($roleMember)]
            } catch {
                # User not found in collection
                continue
            }

            # Add role user to collection
            $rUser = New-Object -TypeName PSObject -Property @{
                UserPrincipalName=$user.UserPrincipalName
                Role=$role.displayName
                MobilePhone=$user.MobilePhone
                PhoneAppDevice=$user.PhoneAppDevice
                Default=$user.Default
                FIDO2Device=$user.FIDO2Device
                SoftwareOATH=$user.SoftwareOATH
                HelloForBusiness=$user.HelloForBusiness
            }
            if ($IncludePerUserState) {
                $rUser | Add-Member -MemberType NoteProperty -Name MFAState -Value $user.MFAState
            }
            if ($IncludeDisabledUsers) {
                $rUser | Add-Member -MemberType NoteProperty -Name AccountState -Value $user.AccountState
            }
            if ($IncludeGuests) {
                $rUser | Add-Member -MemberType NoteProperty -Name UserType -Value $user.UserType
            }
            $roleUsers.Add($rUser) | Out-Null
        }
    }
}

<#

    OUTPUT GENERATION

    The next components generate the required output either to CSV or perform the
    bootstrap HTML generation.

#>

Write-Host "$(Get-Date) Generating Output to $($Output)"

if ($CSV) {
    $props = @("UserPrincipalName")
    if ($IncludeGuests) {$props += "UserType"}
    if ($IncludeDisabledUsers) {$props += "AccountState"}
    if ($IncludePerUserState) {$props += "MFAState"}
    $props += "MobilePhone","PhoneAppDevice","PhoneAppType","FIDO2Device","SoftwareOATH","HelloForBusiness","Default"
    $mfaUsers | Select-Object -Property $props | Export-CSV -NoTypeInformation $Output   
}

if ($HTML) {
    # Stats generation
    if ($IncludePerUserState) {
        $groupedState = $mfaUsers | Group-Object -Property MFAState
        $mfa_disabled = ($groupedState | Where-Object {$_.Name -eq "disabled"}).Count
        $mfa_enabled = ($groupedState | Where-Object {$_.Name -eq "enabled"}).Count
        $mfa_enforced = ($groupedState | Where-Object {$_.Name -eq "enforced"}).Count
    }

    # Enrollment stats
    $mfaDefaultSet = $mfaUsers | Where-Object {$_.Default}
    $mfaDefaultNotSet = $mfaUsers | Where-Object {-not $_.Default}

    $roles = $roleUsers | Group-Object -Property Role

    # Header
    $HTMLOutput = "
    <html>
    <head>
    <script src='https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/js/bootstrap.min.js'></script>
    <link href='https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css' rel='stylesheet' integrity='sha384-ggOyR0iXCbMQv3Xipma34MD+dH/1fQ784/j6cY/iJTQUOhcWr7x9JvoRxT2MZw1T' crossorigin='anonymous'>
    </head>
    <body>
    
    <main role='main' class='container-fluid'>
        <div class='jumbotron'>
            <h1>MFA Report</h1> 
            <p>MFA report generated: $(Get-Date)</p>"
    if ($IncludePerUserState -or $IncludeGuests -or $IncludeDisabledUsers) {
        $includeOptions = @()
        $includeString = "Included options: "
        if ($IncludePerUserState) {$includeOptions += "Per-user MFA state"}
        if ($IncludeGuests) {$includeOptions += " Guest users"}
        if ($IncludeDisabledUsers) {$includeOptions += " Disabled users"}
        $includeString += $($includeOptions -join ", ")
        $HTMLOutput += "
            <p>$includeString</p>
            "
    }
    $HTMLOutput += "<p>User details row color scheme:"
    if ($IncludePerUserState) {
        $HTMLOutput += "
            <div class='table-danger'>Per-user MFA disabled</div>
            <div class='table-warning'>Per-user MFA enabled or enforced by MFA registration not complete</div>
        "
        if (-not $ExcludeRoleBreakdown) {
            $HTMLOutput += "
            <div class='table-success'>Per-user MFA enforced and user registered (Role breakdown only)</div>
            "
        }
        $HTMLOutput += "
            </p>
        </div>
        "
    } else {
        $HTMLOutput += "
            <div class='table-danger'>MFA registration not complete</div>
            </p>
        </div>
        "
    }

    if ($IncludePerUserState) {
        $HTMLOutput += "
        <div class='card'>
            <div class='card-header'>
                Summary of MFA Enforcement State
            </div>
            <div class='card-body'>
                <p><Strong>User MFA Breakdown by State</Strong></p>

                <div class='container'>
                    <div class='row'>
                        <div class='col-sm'>
                            <div class='card text-white bg-danger mb-3' style='max-width: 18rem;'>
                                <div class='card-header'>
                                    <center>Disabled</center>
                                </div>
                                <div class='card-body'>
                                    <center><h5 class='card-title'>$($mfa_disabled)</h5></center>
                                </div>
                            </div>
                        </div>
                        <div class='col-sm'>
                            <div class='card text-white bg-warning mb-3' style='max-width: 18rem;'>
                                <div class='card-header'>
                                    <center>Enabled</center>
                                </div>
                                <div class='card-body'>
                                    <center><h5 class='card-title'>$($mfa_enabled)</h5></center>
                                </div>
                            </div>
                        </div>
                        <div class='col-sm'>
                            <div class='card text-white bg-success mb-3' style='max-width: 18rem;'>
                            <div class='card-header'>
                                <center>Enforced</center>
                            </div>
                            <div class='card-body'>
                                <center><h5 class='card-title'>$($mfa_enforced)</h5></center>
                            </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        "
    }
    $HTMLOutput += "
    <div class='card mt-4'>
        <div class='card-header'>
            Summary of MFA Registration
        </div>
        <div class='card-body'>
            <p><Strong>User MFA Breakdown by Registration</Strong></p>

            <div class='container'>
                <div class='row'>
                    <div class='col'>
                        <div class='card text-white bg-danger mb-3' style='max-width: 18rem;'>
                            <div class='card-header'>
                                <center>Not Registered</center>
                            </div>
                            <div class='card-body'>
                                <center><h5 class='card-title'>$($mfaDefaultNotSet.Count)</h5></center>
                            </div>
                        </div>
                    </div>
                    <div class='col'>
                        <div class='card text-white bg-success mb-3' style='max-width: 18rem;'>
                            <div class='card-header'>
                                <center>Registered</center>
                            </div>
                            <div class='card-body'>
                                <center><h5 class='card-title'>$($mfaDefaultSet.Count)</h5></center>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

        <p><Strong>User MFA Registration Breakdown by Default Method</Strong></p>
        <table class='table'>
        "

        foreach ($type in $($mfaDefaultSet | Group-Object -Property Default | Sort-Object -Property Count -Descending)) {
            $pctTotal = [math]::Round(($type.Count/$mfaDefaultSet.Count)*100)
            $HTMLOutput += "
            <tr>
            <td style='width: 200px'>
            <p><strong>$($type.Name)</strong></p>
            <p><i>$($type.Count) of $($mfaDefaultSet.Count) ($pctTotal%)</i></p>
            </td>
            <td>
            <div class='progress'>
                <div class='progress-bar progress-bar-striped bg-success' role='progressbar' style='width: $($pctTotal)%' aria-valuenow='$($pctTotal)' aria-valuemin='0' aria-valuemax='100'></div>
            </div>
            </td>
            </tr>
            "
        }

        $HTMLOutput += "
        </table>
        </div>
        </div>
        "

    if (-not $ExcludeRoleBreakdown) {
        # Create the per-role breakdown summary
        foreach ($role in $roles) {
            $mfaRoleDefaultSet = @($role.Group | Where-Object {$_.Default})
            $mfaRoleDefaultNotSet = @($role.Group | Where-Object {-not $_.Default})

            $HTMLOutput += "

            <div class='card mt-4'>
            <div class='card-header'>
                Summary of MFA Registration - $($role.Name)
            </div>
            <div class='card-body'>
                <p><Strong>$($Role.Name) MFA Breakdown by Registration</Strong></p>

                <div class='container'>
                    <div class='row'>
                        <div class='col'>
                            <div class='card text-white bg-danger mb-3' style='max-width: 18rem;'>
                                <div class='card-header'>
                                    <center>Not Registered</center>
                                </div>
                                <div class='card-body'>
                                    <center><h5 class='card-title'>$($mfaRoleDefaultNotSet.Count)</h5></center>
                                </div>
                            </div>
                        </div>
                        <div class='col'>
                            <div class='card text-white bg-success mb-3' style='max-width: 18rem;'>
                                <div class='card-header'>
                                    <center>Registered</center>
                                </div>
                                <div class='card-body'>
                                    <center><h5 class='card-title'>$($mfaRoleDefaultSet.Count)</h5></center>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>

            <p><Strong>User MFA Registration Breakdown by Default Method</Strong></p>
            <table class='table'>
            "

            foreach ($type in $($mfaRoleDefaultSet | Group-Object -Property Default | Sort-Object -Property Count -Descending)) {
                $pctTotal = [math]::Round(($type.Count/$mfaRoleDefaultSet.Count)*100)
                $HTMLOutput += "
                <tr>
                <td style='width: 200px'>
                <p><strong>$($type.Name)</strong></p>
                <p><i>$($type.Count) of $($mfaRoleDefaultSet.Count) ($pctTotal%)</i></p>
                </td>
                <td>
                <div class='progress'>
                    <div class='progress-bar progress-bar-striped bg-success' role='progressbar' style='width: $($pctTotal)%' aria-valuenow='$($pctTotal)' aria-valuemin='0' aria-valuemax='100'></div>
                </div>
                </td>
                </tr>
                "
            }

            $HTMLOutput += "
            </table>
            <p><strong>Users in Role</strong></p>
            <table class='table table-sm table-striped'>
            <thead class='thead-light'>
                <tr>
                    <th scope='col'>User Principal Name</th>"

            if ($IncludeGuests) {$HTMLOutput += "<th scope='col'>User Type</th>"}
            if ($IncludeDisabledUsers) {$HTMLOutput += "<th scope='col'>Account State</th>"}
            if ($IncludePerUserState) {$HTMLOutput += "<th scope='col'>MFA State</th>"}
            $HTMLOutput += "
                    <th scope='col'>Mobile Phone</th>
                    <th scope='col'>Phone App Device</th>
                    <th scope='col'>FIDO2 Device</th>
                    <th scope='col'>Software OATH</th>
                    <th scope='col'>Hello for Business</th>
                    <th scope='col'>Default</th>
                </tr>
            </thead>"

            # Add rows for each MFA User
            foreach ($u in $($role.Group | Sort-Object -Property UserPrincipalName)) {
                if ($IncludePerUserState) {
                    # Enrollment color scheme
                    if ($u.MFAState -eq "Disabled") {
                        $TRClass = "table-danger"
                    } elseif ($u.MFAState -eq "Enabled") {
                        $TRClass = "table-warning"
                    } elseif (-not $u.Default) {
                        $TRClass = "table-warning"
                    } else {
                        $TRClass = ""
                    }
                } elseif (-not $u.Default) {
                    $TRClass = "table-danger"
                } else {
                    $TRClass = ""
                }

                $HTMLOutput += "
                <tr class='$($TRClass)'>
                    <td>$($u.UserPrincipalName)</td>
                    "
                if ($IncludeGuests) {$HTMLOutput += "<td>$($u.UserType)</td>"}
                if ($IncludeDisabledUsers) {$HTMLOutput += "<td>$($u.AccountState)</td>"}
                if ($IncludePerUserState) {$HTMLOutput += "<td>$($u.MFAState)</td>"}
                $HTMLOutput += "
                    <td>$($u.MobilePhone)</td>
                    <td>$($u.PhoneAppDevice.Replace(' | ','<br>'))</td>
                    <td>$($u.FIDO2Device.Replace(' | ','<br>'))</td>
                    <td>$($u.SoftwareOATH.Replace(' | ','<br>'))</td>
                    <td>$($u.HelloForBusiness.Replace(' | ','<br>'))</td>
                    <td>$($u.Default)</td>
                </tr>
                "
            }

            $HTMLOutput += "</table>
            </div>
            </div>"
        }
    }

    $HTMLOutput += "
    
    <div class='card mt-4'>
    <div class='card-header'>User List</div>

    <table class='table table-sm table-striped'>
    <thead class='thead-light'>
        <tr>
            <th scope='col'>User Principal Name</th>"
    if ($IncludeGuests) {$HTMLOutput += "<th scope='col'>User Type</th>"}
    if ($IncludeDisabledUsers) {$HTMLOutput += "<th scope='col'>Account State</th>"}
    if ($IncludePerUserState) {$HTMLOutput += "<th scope='col'>MFA State</th>"}
    $HTMLOutput += "
            <th scope='col'>Mobile Phone</th>
            <th scope='col'>Phone App Device</th>
            <th scope='col'>FIDO2 Device</th>
            <th scope='col'>Software OATH</th>
            <th scope='col'>Hello for Business</th>
            <th scope='col'>Default</th>
        </tr>
    </thead>
    "

    # Add rows for each MFA User
    foreach($u in $($mfaUsers | Sort-Object -Property UserPrincipalName)) {

        # Enrollment color scheme
        if ($IncludePerUserState) {
            # Enrollment color scheme
            if ($u.MFAState -eq "Disabled") {
                $TRClass = "table-danger"
            } elseif ($u.MFAState -eq "Enabled") {
                $TRClass = "table-warning"
            } elseif (-not $u.Default) {
                $TRClass = "table-warning"
            } else {
                $TRClass = ""
            }
        } elseif (-not $u.Default) {
            $TRClass = "table-danger"
        } else {
            $TRClass = ""
        }

        $HTMLOutput += "
        <tr class='$($TRClass)'>
            <td>$($u.UserPrincipalName)</td>
            "
        if ($IncludeGuests) {$HTMLOutput += "<td>$($u.UserType)</td>"}
        if ($IncludeDisabledUsers) {$HTMLOutput += "<td>$($u.AccountState)</td>"}
        if ($IncludePerUserState) {$HTMLOutput += "<td>$($u.MFAState)</td>"}
        $HTMLOutput += "
            <td>$($u.MobilePhone)</td>
            <td>$($u.PhoneAppDevice.Replace(' | ','<br>'))</td>
            <td>$($u.FIDO2Device.Replace(' | ','<br>'))</td>
            <td>$($u.SoftwareOATH.Replace(' | ','<br>'))</td>
            <td>$($u.HelloForBusiness.Replace(' | ','<br>'))</td>
            <td>$($u.Default)</td>
        </tr>
        "
    }

    # Closing
    $HTMLOutput += "
    </table>
    </div>
    </main>
    </body>
    </html>
    "

    # Output
    $HTMLOutput | Out-File -FilePath $Output

}
