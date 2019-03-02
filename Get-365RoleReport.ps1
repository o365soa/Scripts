#Requires -Version 4

<#
	.SYNOPSIS
		Office 365 Role Report

	.DESCRIPTION

	.NOTES
		Cam Murray
		Field Engineer - Microsoft
        cam.murray@microsoft.com
        
        This script uses Bootstrap in the output report. For more information https://www.getbootstrap.com/

	.LINK
		about_functions_advanced

#>


Param(
    [CmdletBinding()]
    [Int16]$ThresholdPasswordReset=30,
    [String]$Output = "report.html",
    [Array]$IgnoredRoles="Directory Synchronization Accounts"
)


# Connect to the MSOL Service. MSOL still required due to MFA support.

Connect-MsolService

$pUsers = @() 
$mRoles = Get-MsolRole

ForEach($mRole in $mRoles) {

    # Get the member
    $mRoleUsers = Get-MsolRoleMember -RoleObjectId $mRole.ObjectId | Where-Object {$_.RoleMemberType -ne 'ServicePrincipal'}

    # Itterate each user
    ForEach($mRoleUser in $mRoleUsers) {

        # Get underlying msol user
        $mUser = Get-MsolUser -ObjectId $mRoleUser.ObjectId

        # Determine password age
        $PasswordAge = ((Get-Date) - $mUser.LastPasswordChangeTimestamp).Days

        # Determine MFA state
        $MFA_State = $mUser.StrongAuthenticationRequirements.State

        # Determine default MFA method
        $MFA_Default = ($mUser.StrongAuthenticationMethods | Where-Object {$_.IsDefault -eq $true})

        # MFA Phone Number
        $MFA_Phone = $mUser.StrongAuthenticationUserDetails.PhoneNumber

        # Determine if cloud user
        If($mUser.ImmutableID -eq $null) {
            $Type = "Cloud"
        } else {
            $Type = "Synchronised"
        }


        # Add to final object

        $pUsers += New-Object -TypeName PSObject -Property @{
            SignInName=$mUser.SignInName
            PasswordAge=$PasswordAge
            Role=$($mRole.Name)
            MFA_Default=$($MFA_Default.MethodType)
            MFA_Phone=$($MFA_Phone)
            MFA_State=$($MFA_State)
            Type = $Type
        }
    }

}

# Write the report

$Report = "<html><head><link rel='stylesheet' href='https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0-beta.3/css/bootstrap.min.css' integrity='sha384-Zug+QiDoJOrZ5t4lssLdxGhVrurbmBWopoEl+M6BdEfwnCJZtKxi1KgxUyJq13dy' crossorigin='anonymous'></head><body>"
$Report += "<div class='jumbotron jumbotron-fluid'>
<div class='container'>
  <h1 class='display-4'>Office 365 Administrator Report</h1>
  <p class='lead'>This report contains details for Office 365 Administrators. Generated on $(Get-Date)</p>
</div>
</div>"

$RoleGrouping = $pUsers | Group-Object Role

ForEach($r in $RoleGrouping) {
    if($IgnoredRoles -notcontains $r.Name) {
    $Report += "<div class='card'>
    <div class='card-header'>
      $($r.Name)
    </div>"

    $Report += "<div class='card-body'><table class='table'>
    <thead>
      <tr>
        <th>Sign In Name</th>
        <th>Type</th>
        <th>Password Age</th>
        <th>MFA State</th>
        <th>MFA Default</th>
        <th>MFA Phone</th>
      </tr>
    </thead>
    <tbody>"

    ForEach($u in $r.Group) {
        $Report += "<tr>"
        $Report += "<td>$($u.SignInName)</td>"
        $Report += "<td>$($u.Type)</td>"
        
        If($u.PasswordAge -ge $ThresholdPasswordReset) { $Class = "table-danger" } else { $Class = "table-success" }
        $Report += "<td class='$Class'>$($u.PasswordAge) Days</td>"

        If($u.MFA_State -ne "Enforced") { $Class = "table-danger" } else { $Class = "table-success" }
        $Report += "<td class='$Class'>$($u.MFA_State)</td>"

        If($u.MFA_Default -eq $null) { $Class = "table-danger" } else { $Class = "table-success" }
        $Report += "<td class='$Class'>$($u.MFA_Default)</td>"

        $Report += "<td>$($u.MFA_Phone)</td>"
        $Report += "</tr>"

    }
  

    $Report += "</tbody></table></div>
  </div>"
}
}

$Report | Out-File $Output
Invoke-Item $Output