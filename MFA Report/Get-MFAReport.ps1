#Requires -Version 4

<#
	.SYNOPSIS
		Creates a report in either CSV or HTML format of all MFA enrollment.

	.DESCRIPTION
        This script generates a report in either CSV or HTML format for all MFA enrollment

        HTML format uses the opensource bootstrap libraries for a tidy interface.

	.EXAMPLE
		PS C:\> .\Get-MFASettings.ps1

	.NOTES
		Cam Murray
		Field Engineer - Microsoft
		cam.murray@microsoft.com
		
		For updates, and more scripts, visit https://github.com/O365SOA/Scripts
		
		Last update: Feb 2019

	.LINK
		about_functions_advanced

#>

Param(
  [string]$Output,
  [Array]$IgnoreRoles=@("Guest Inviter","Partner Tier1 Support","Partner Tier2 Support","Directory Readers","Device Users","Device Join","Workplace Device Join"),

  [Parameter(ParameterSetName='CSVOutput')]

  [switch]$CSV,

  [Parameter(ParameterSetName='HTMLOutput')]

  [switch]$HTML

 )

 # Build a default output file if none specified
 If(!$output) {
    If($CSV) {
        $Output = "mfa-report.csv"
    }
    If($HTML) {
        $Output = "mfa-report.html"
    }
 }

 # Determine if connected to MSOL
 Try {
     Get-MsolCompanyInformation -ErrorAction:Stop | Out-Null
 } Catch {
     Write-Error "Error running MSOL command, ensure you run Connect-MsolService first!"
     Exit
 }

 <# 
 
    User Checking

    The following goes through all users in Get-MsolUser, and then parses their StrongAuth settings.
 
 #>

Write-Host "$(Get-Date) Getting all MSOL Users - this may take some time.."
$MFAUsers = @()

ForEach($User in (Get-MsolUser -All)) {

    Write-Progress -Id 1 -Activity "Parsing MFA Users.." -Status $($User.UserPrincipalName)

    # Reset variables
    $Default = $null; $MFAState = $null; $PhoneAppType = $null;

    # Get the default MFA option
    $Default = ($User.StrongAuthenticationMethods | Where-Object {$_.IsDefault -eq $true}).MethodType

    # If StrongAuthenticationRequirements is not set
    If($($User.StrongAuthenticationRequirements.State) -eq $null) {
        $MFAState = "Not Set"
    } else {
        $MFAState = $($User.StrongAuthenticationRequirements.State)
    }

    # If PhoneAppType is specified convert to string and replace commas for CSV compat
    If($User.StrongAuthenticationPhoneAppDetails.AuthenticationType) {
        $PhoneAppType = $($User.StrongAuthenticationPhoneAppDetails.AuthenticationType.ToString().Replace(",","+"))
    }

    # Add to variable
    $MFAUsers += New-Object -TypeName psobject -Property @{
		UserPrincipalName=$($User.UserPrincipalName)
		State=$MFAState
        Phone=$($User.StrongAuthenticationUserDetails.PhoneNumber)
        PhoneAppDevice=$($User.StrongAuthenticationPhoneAppDetails.DeviceName)
        PhoneAppType=$PhoneAppType
        Default=$Default
        ObjectId=$($User.ObjectId)
    }
}

$RoleUsers = @()

Write-Host "$(Get-Date) Getting Role Member Status.. this may take sometime.."

ForEach($Role in (Get-MsolRole)) {

    If($IgnoreRoles -notcontains $Role.Name) {
        Write-Progress -Id 1 -Activity "Parsing Group Users" -Status $($Role.Name)

        # Get members of this role
        $RoleMembers = Get-MsolRoleMember -RoleObjectId $Role.ObjectId -All | Where-Object {$_.RoleMemberType -eq "User"}
    
        # Loop through each and add settings
        ForEach($RoleMember in $RoleMembers) {
    
            # Get responding user object
            $User = $MFAUsers | Where-Object {$_.ObjectId -eq $($RoleMember.ObjectId)}
    
            # Add to array
            $RoleUsers += New-Object -TypeName PSObject -Property @{
                UserPrincipalName=$($User.UserPrincipalName)
                Role=$($Role.Name)
                State=$User.State
                Phone=$($User.Phone)
                PhoneAppDevice=$($User.PhoneAppDevice)
                PhoneAppType=$User.PhoneAppType
                Default=$User.Default
            }
        }
    }
}

<#

    OUTPUT GENERATION

    The next components generate the required output either to CSV or perform the
    bootstrap HTML generation.

#>

Write-Host "$(Get-Date) Generating Output to $($Output)"

If($CSV) {

    <#
    
        CSV Generation
    
    #>

    $MFAUsers | Select UserPrincipalName,State,Default,Phone,PhoneAppDevice,PhoneAppType | Export-CSV -NoTypeInformation $Output

}

If($HTML) {

    <#
    
        HTML Generation
        Slightly more complicated because we use bootstrap.

    #>

    # Stats generation

    $GroupedState = ($MFAUsers | Group-Object State)

    $mfa_Total = $MFAUsers.Count

    $mfa_NotSet = ($GroupedState | Where-Object {$_.Name -eq "Not Set"}).Count
    $mfa_NotSet_pct = [Math]::Round($($mfa_NotSet)/$($mfa_Total)*100)

    $mfa_Enabled = ($GroupedState | Where-Object {$_.Name -eq "Enabled"}).Count
    $mfa_Enabled_pct = [Math]::Round($($mfa_Enabled)/$($mfa_Total)*100)

    $mfa_Enforced = ($GroupedState | Where-Object {$_.Name -eq "Enforced"}).Count
    $mfa_Enforced_pct = [Math]::Round($($mfa_Enforced)/$($mfa_Total)*100)
    


    # Enrollment stats

    $MFADefaultSet = ($MFAUsers | Where-Object {$_.Default -ne $Null})
    $MFADefaultNotSet = ($MFAUsers | Where-Object {$_.Default -eq $Null})

    $mfa_Methods_Group = ($MFADefaultSet | Group-Object Default)

    $Roles = ($RoleUsers | Group-Object Role)

    # Header
    $HTMLOutput = "
    <html>
    <head>
    <script src='https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/js/bootstrap.min.js'></script>
    <link href='https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css' rel='stylesheet' integrity='sha384-ggOyR0iXCbMQv3Xipma34MD+dH/1fQ784/j6cY/iJTQUOhcWr7x9JvoRxT2MZw1T' crossorigin='anonymous'>
    </head>
    <body>
    
    <main role='main' class='container'>
        <div class='jumbotron'>
            <h1>MFA Report</h1> 
            <p>MFA report generated at $(Get-Date)</p>
        </div>
        
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
                                    <center>Not Set</center>
                                </div>
                                <div class='card-body'>
                                    <center><h5 class='card-title'>$($mfa_NotSet)</h5></center>
                                </div>
                            </div>
                        </div>
                        <div class='col-sm'>
                            <div class='card text-white bg-warning mb-3' style='max-width: 18rem;'>
                                <div class='card-header'>
                                    <center>Enabled</center>
                                </div>
                                <div class='card-body'>
                                    <center><h5 class='card-title'>$($mfa_Enabled)</h5></center>
                                </div>
                            </div>
                        </div>
                        <div class='col-sm'>
                            <div class='card text-white bg-success mb-3' style='max-width: 18rem;'>
                            <div class='card-header'>
                                <center>Enforced</center>
                            </div>
                            <div class='card-body'>
                                <center><h5 class='card-title'>$($mfa_Enforced)</h5></center>
                            </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

    <div class='card mt-4'>
        <div class='card-header'>
            Summary of MFA Enrollment
        </div>
        <div class='card-body'>
            <p><Strong>User MFA Breakdown by Enrollment</Strong></p>

            <div class='container'>
                <div class='row'>
                    <div class='col'>
                        <div class='card text-white bg-danger mb-3' style='max-width: 18rem;'>
                            <div class='card-header'>
                                <center>Not Enrolled</center>
                            </div>
                            <div class='card-body'>
                                <center><h5 class='card-title'>$($MFADefaultNotSet.Count)</h5></center>
                            </div>
                        </div>
                    </div>
                    <div class='col'>
                        <div class='card text-white bg-success mb-3' style='max-width: 18rem;'>
                            <div class='card-header'>
                                <center>Enrolled</center>
                            </div>
                            <div class='card-body'>
                                <center><h5 class='card-title'>$($MFADefaultSet.Count)</h5></center>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

        <p><Strong>User MFA Enrollment Breakdown by Default Method</Strong></p>
        <table class='table'>
        "

        ForEach($Type in $($MFADefaultSet | Group-Object Default | Sort-Object Count -Descending)) {

            $PCTTotal = [Math]::Round(($Type.Count/$MFADefaultSet.Count)*100)
            $HTMLOutput += "
            <tr>
            <td style='width: 200px'>
            <p><strong>$($Type.Name)</strong></p>
            <p><i>$($Type.Count) of $($MFADefaultSet.Count) ($PCTTotal%)</i></p>
            </td>
            <td>
            <div class='progress'>
                <div class='progress-bar progress-bar-striped bg-success' role='progressbar' style='width: $($PCTTotal)%' aria-valuenow='$($PCTTotal)' aria-valuemin='0' aria-valuemax='100'></div>
            </div>
            </td>
            </tr>"
            
        }


        $HTMLOutput += "
        </table>
        </div>
    </div>"

    # Create the per role break down summary

    ForEach ($Role in $Roles) {

        $MFARoleDefaultSet = @($Role.Group | Where-Object {$_.Default -ne $Null})
        $MFARoleDefaultNotSet = @($Role.Group | Where-Object {$_.Default -eq $Null})

    $HTMLOutput += "

    <div class='card mt-4'>
        <div class='card-header'>
            Summary of MFA Enrollment - $($Role.Name)
        </div>
        <div class='card-body'>
            <p><Strong>$($Role.Name) MFA Breakdown by Enrollment</Strong></p>

            <div class='container'>
                <div class='row'>
                    <div class='col'>
                        <div class='card text-white bg-danger mb-3' style='max-width: 18rem;'>
                            <div class='card-header'>
                                <center>Not Enrolled</center>
                            </div>
                            <div class='card-body'>
                                <center><h5 class='card-title'>$($MFARoleDefaultNotSet.Count)</h5></center>
                            </div>
                        </div>
                    </div>
                    <div class='col'>
                        <div class='card text-white bg-success mb-3' style='max-width: 18rem;'>
                            <div class='card-header'>
                                <center>Enrolled</center>
                            </div>
                            <div class='card-body'>
                                <center><h5 class='card-title'>$($MFARoleDefaultSet.Count)</h5></center>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

        <p><Strong>User MFA Enrollment Breakdown by Default Method</Strong></p>
        <table class='table'>
        "

        ForEach($Type in $($MFARoleDefaultSet | Group-Object Default | Sort-Object Count -Descending)) {

            $PCTTotal = [Math]::Round(($Type.Count/$MFARoleDefaultSet.Count)*100)
            $HTMLOutput += "
            <tr>
            <td style='width: 200px'>
            <p><strong>$($Type.Name)</strong></p>
            <p><i>$($Type.Count) of $($MFARoleDefaultSet.Count) ($PCTTotal%)</i></p>
            </td>
            <td>
            <div class='progress'>
                <div class='progress-bar progress-bar-striped bg-success' role='progressbar' style='width: $($PCTTotal)%' aria-valuenow='$($PCTTotal)' aria-valuemin='0' aria-valuemax='100'></div>
            </div>
            </td>
            </tr>"
            
        }


        $HTMLOutput += "
        </table>
        <p><strong>Users in Role</strong></p>
        <table class='table table-striped table-dark'>
        <thead>
            <tr>
                <th scope='col'>UserPrincipalName</th>
                <th scope='col'>State</th>
                <th scope='col'>Default</th>
                <th scope='col'>Phone Number</th>
                <th scope='col'>Phone App Device</th>
                <th scope='col'>Phone App Type</th>
            </tr>
        </thead>"

            # Add rows for each MFA User
    ForEach($U in $($Role.Group | Sort-Object UserPrincipalName)) {

        # Enrollment colour scheming
        If($U.State -eq "Not Set") {
            $TRClass = ""
        } ElseIf($U.State -eq "Enabled") {
            $TRClass = "bg-warning"
        } Else {
            $TRClass = "bg-success"
        }

        $HTMLOutput += "
        <tr class='$($TRClass)'>
            <td>$($u.UserPrincipalName)</td>
            <td>$($u.State)</td>
            <td>$($u.Default)</td>
            <td>$($u.Phone)</td>
            <td>$($u.PhoneAppDevice)</td>
            <td>$($u.PhoneAppType)</td>
        </tr>
        "
    }

        $HTMLOutput += "</table>
        </div>
    </div>"
    }

    $HTMLOutput += "
    
    <div class='card mt-4'>
        <div class='card-header'>
        User List
        </div>

        <table class='table table-striped table-dark'>
        <thead>
            <tr>
                <th scope='col'>UserPrincipalName</th>
                <th scope='col'>State</th>
                <th scope='col'>Default</th>
                <th scope='col'>Phone Number</th>
                <th scope='col'>Phone App Device</th>
                <th scope='col'>Phone App Type</th>
            </tr>
        </thead>
    "

    # Add rows for each MFA User
    ForEach($U in $($MFAUsers | Sort-Object UserPrincipalName)) {

        # Enrollment colour scheming
        If($U.State -eq "Not Set") {
            $TRClass = ""
        } ElseIf($U.State -eq "Enabled") {
            $TRClass = "bg-warning"
        } Else {
            $TRClass = "bg-success"
        }

        $HTMLOutput += "
        <tr class='$($TRClass)'>
            <td>$($u.UserPrincipalName)</td>
            <td>$($u.State)</td>
            <td>$($u.Default)</td>
            <td>$($u.Phone)</td>
            <td>$($u.PhoneAppDevice)</td>
            <td>$($u.PhoneAppType)</td>
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