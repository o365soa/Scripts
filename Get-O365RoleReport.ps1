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
#Requires -Modules MSOnline, ExchangeOnlineManagement

<#
	.SYNOPSIS
		Office 365 Administrative Role Report

	.DESCRIPTION
		Enumerates members of administrative roles in Azure AD, Security & Compliance, and Exchange Online.
	.PARAMETER PasswordAgeThreshold
		Passwords older than this age (in days) are highlighted in red.  Default is 90 days.
	.PARAMETER SkipWorkload
		Provide one or more workloads (comma-separated or array) to skip.  Valid values are AAD, SCC, EXO.
	.PARAMETER AdminUPN
		UPN of account to use when connecting to Exchange and SCC.  Helps to avoid unnecessary auth
		prompts if you have connected with the account before.
	.PARAMETER IgnoredRoles
		Array of roles to exclude from the report.  Default is AAD role of "Directory Synchronization Accounts".
	.PARAMETER Output
		Path and filename of the report.  Default is O365RoleReport.html in the current directory.
    .PARAMETER UseIEProxyConfig
        When a proxy is required, you may need to use the proxy configuration from IE to connect.
        This creates a PSSession Option which is used for Remote PowerShell.
	.NOTES
        Version 2.3
		April 27, 2022
		
		This script uses Bootstrap to format the report. For more information https://www.getbootstrap.com/

#>

[CmdletBinding()]
Param(
    [Int16]$PasswordAgeThreshold=90,
    [String]$Output = "O365RoleReport.html",
    [Array]$IgnoredRoles="Directory Synchronization Accounts",
	[ValidateSet('AAD','SCC','EXO')]$SkipWorkload,
	[string]$AdminUPN,
    [Switch]$UseIEProxyConfig
)


function Get-UserDetails ($id) {
	$user = Get-MsolUser -ObjectId $id
	
	# Determine password age
    $passwordAge = ((Get-Date) - $user.LastPasswordChangeTimestamp).Days

    # Determine default MFA method
    $mfaDefault = ($user.StrongAuthenticationMethods | Where-Object {$_.IsDefault -eq $true})

    # MFA Phone Number
    $mfaPhone = $user.StrongAuthenticationUserDetails.PhoneNumber

	# Determine if cloud user
    if ($user.ImmutableID -eq $null) {
    	$type = "Cloud"
    }
	else {
        $type = "Synchronized"
    }
	
	return New-Object -TypeName PSObject -Property @{
        SignInName = $user.SignInName
        PasswordAge = $passwordAge
        MFADefault = $mfaDefault.MethodType
        MFAPhone = $mfaPhone
        MFAState = $mfaState
        UserType = $type
    }
}

function Get-ExoRoleGroupMembers {
	param (
		$roleGroup,
		$roleName,
		$parentGroupName
	)
	$rgm = Get-RoleGroupMember -Identity $roleGroup.Identity
	$members = @()
	foreach ($gMember in $rgm) {
		if ($gMember.RecipientType -eq 'Group') {
			# Member is Exchange role group
			if ($parentGroupName) {
				Write-Verbose -Message "Nested role group of $parentGroupName in $roleName role group: $($gMember.Name)"
			}
			else {
				Write-Verbose -Message "Role group in $roleName role group: $($gMember.Name)"
			}
			$members += Get-ExoRoleGroupMembers -roleGroup $gMember -roleName $roleName -parentGroup $gMember.Name
		}
		elseif ($rgm.RecipientType -eq 'MailUniversalSecurityGroup') {
			# Member is Security DL
			if ($parentGroupName) {
				Write-Verbose -Message "Nested mail-enabled security group of $parentGroupName in $roleName role group: $($gMember.Name)"
			}
			else {
				Write-Verbose -Message "Mail-enabled security group in $roleName role group: $($gMember.Name)"				
			}
			$members += Get-ExoSecurityGroupMembers -group $gMember -roleName $roleName -parentGroupName $gMember.Name
		}
		else {
			# Member is individual
			if ($parentGroupName) {
				Write-Verbose -Message "User in role group $parentGroupName assigned roles of $roleName role group: $($gMember.Name) ($($gMember.WindowsLiveId))"
				$pgName = $parentGroupName + "\"
			}
			else {
				Write-Verbose -Message "User assigned roles of $roleName role group: $($gMember.Name) ($($gMember.WindowsLiveId))"
				$pgName = ""
			}
			$members += New-Object -TypeName PSObject -Property @{
				Id = $gMember.ExternalDirectoryObjectId
				ParentGroup = $pgName
			}
		}
	}
	return $members
}

function Get-ExoSecurityGroupMembers {
	param (
		$group,
		$roleName,
		$parentGroupName
	)
	$sgm = Get-DistributionGroupMember -Identity $group.Identity
	$members = @()
	foreach ($gMember in $sgm) {
		if ($gMember.RecipientType -like "*Group") {
			# Member is security group
			if ($parentGroupName) {
				Write-Verbose -Message "Nested security group of $parentGroupName in $roleName role group: $($gMember.Name)"
			}
			else {
				Write-Verbose -Message "Security group in $roleName role group: $($gMember.Name)"
			}
			$members += Get-ExoSecurityGroupMembers -group $gMember -roleName $roleName -parentGroupName $gMember.Name
		}
		else {
			# Member is individual
			if ($parentGroupName) {
				Write-Verbose -Message "User in security group $parentGroupName assigned roles of $roleName role group: $($gMember.Name) ($($gMember.WindowsLiveId))"
				$pgName = $parentGroupName + "\"
			}
			else {
				Write-Verbose -Message "User assigned roles of $roleName role group: $($gMember.Name) ($($gMember.WindowsLiveId))"
				$pgName = ""
			}
			$members += New-Object -TypeName PSObject -Property @{
				Id = $gMember.ExternalDirectoryObjectId
				ParentGroup = $pgName
			}
		}
	}
	return $members
}

$workLoads = @()
if ($SkipWorkload -notcontains 'AAD') {$workLoads += 'AAD'}
if ($SkipWorkload -notcontains 'SCC') {$workLoads += 'SCC'}
if ($SkipWorkload -notcontains 'EXO') {$workLoads += 'EXO'}

if ($SkipWorkload -contains 'AAD' -and $SkipWorkload -contains 'SCC' -and $SkipWorkload -contains 'EXO') {
	Write-Error -Message 'At least one workload must not be excluded.'
	exit
}

If ($UseIEProxyConfig) {
    Write-Host "$(Get-Date) [INFO] Engineer has specified using IE Proxy Settings" -ForegroundColor Green
    $ProxySetting = New-PSSessionOption -ProxyAccessType IEConfig -IdleTimeout 9000000 -OperationTimeout 9000000
} Else {
    $ProxySetting = New-PSSessionOption -ProxyAccessType None -IdleTimeout 9000000 -OperationTimeout 9000000
}

# Always connect to MSOL, if necessary, for password and MFA details. MSOL still required due to MFA details not available in AAD v2.
if (-not(Get-MsolCompanyInformation -ErrorAction SilentlyContinue)) {
	Write-Host 'Connecting to Azure AD...'
	Connect-MsolService
}

# Connect to SCC if not skipped, if necessary
# Prefix is used to support connecting to SCC and EXO at the same time
if ($SkipWorkload -notcontains 'SCC') {
	if (-not(Get-Command -Name Get-SCCRoleGroup -ErrorAction SilentlyContinue)) {
			Write-Host 'Connecting to Security & Compliance Center...'
			Connect-IPPSSession -Prefix SCC -UserPrincipalName $AdminUPN -PSSessionOption $ProxySetting
	}
}

# Connect to EXO if not skipped, if necessary
if ($SkipWorkload -notcontains 'EXO') {
	if (-not(Get-Command -Name Get-OrganizationConfig -ErrorAction SilentlyContinue)) {
		Write-Host 'Connecting to Exchange Online...'
		Connect-ExchangeOnline -UserPrincipalName $AdminUPN -ShowBanner:$false -PSSessionOption $ProxySetting
	}
}

$pUsers = @() 

#Process AAD roles
if ($SkipWorkload -notcontains 'AAD') {
	Write-Host 'Getting users with an Azure AD role assignment...'
	
	$mRoles = Get-MsolRole

	foreach ($mRole in $mRoles) {
		Write-Verbose -Message "Processing role $($mRole.Name)"
		$mRoleUsers = @()

		# Get the members, excluding service principals
	    $mRoleUsers = Get-MsolRoleMember -RoleObjectId $mRole.ObjectId | Where-Object {$_.RoleMemberType -ne 'ServicePrincipal'}

	    # Iterate each user
	    foreach ($mRoleUser in $mRoleUsers) {
			Write-Verbose -Message "User assigned $($mRole.Name) role: $($mRoleUser.EmailAddress)"
			
		    # Get underlying MSOL user details
	        $mUser = Get-UserDetails -id $mRoleUser.ObjectId

	        # Add to final object
	        $pUsers += New-Object -TypeName PSObject -Property @{
	            SignInName = $mUser.SignInName
	            PasswordAge = $mUser.PasswordAge
	            Role = $mRole.Name
	            MFADefault = $mUser.MFADefault
	            MFAPhone = $mUser.MFAPhone
	            UserType = $mUser.UserType
				Workload = 'Azure AD'
	        }
	    }
	}
}

# Process SCC roles
if ($SkipWorkload -notcontains 'SCC') {
	Write-Host 'Getting users with a Security & Compliance Center role assignment...'
	
	$sccRoles = Get-SCCRoleGroup
	
	foreach ($sccRole in $sccRoles) {
		Write-Verbose "Processing role $($sccRole.Name)"
		$roleUsers = @()
		
		# Get the members
	    $sgm = Get-SCCRoleGroupMember -Identity $sccRole.Guid.Guid #| Where-Object {$_.RoleMemberType -ne 'ServicePrincipal'}
	    
		# Iterate each member
	   	foreach ($sMember in $sgm) {
			
	        if ($sMember.RecipientType -eq 'Group') {
				Write-Verbose -Message "Group assigned $($sccRole.Name) role: $($sMember.DisplayName)"
				$mgm = Get-MsolGroupMember -GroupObjectId $sMember.ExternalDirectoryObjectId
				foreach ($mMember in $mgm) {
					Write-Verbose -Message "User in $($sMember.DisplayName) group assigned $($sccRole.Name) role: $($mMember.DisplayName) ($($mMember.EmailAddress))"
					$roleUsers += New-Object -TypeName PSObject -Property @{
						Id = $mMember.ObjectId.Guid
						ParentGroup = $sMember.Displayname + "\"
					}
				}
			}
			else {
				if ($sMember.PrimarySMTPAddress) {
					$memberID = $sMember.PrimarySMTPAddress
				}
				else {
					$memberID = "No email address"
				}
				Write-Verbose -Message "User assigned $($sccRole.Name) role: $($sMember.Name) ($memberID)"
				$roleUsers += New-Object -TypeName PSObject -Property @{
					Id = $sMember.ExternalDirectoryObjectId
					ParentGroup = ""
				}
			}
		}

	    # Iterate each user
	    foreach ($user in $roleUsers) {
			
	        # Get underlying MSOL user details
	        $mUser = Get-UserDetails -id $user.Id
					
			# Add to final object
	        $pUsers += New-Object -TypeName PSObject -Property @{
	            SignInName = $user.ParentGroup + $mUser.SignInName
	            PasswordAge = $mUser.PasswordAge
	            Role = $sccRole.Name
	            MFADefault = $mUser.MFADefault
	            MFAPhone = $mUser.MFAPhone
	            UserType = $mUser.UserType
				Workload = 'Security and Compliance'
	        }
		}			
	}
}

if ($SkipWorkload -notcontains 'EXO') {
	Write-Host 'Getting users with an Exchange Online role assignment...'
	#$exoRoleGroups = Get-RoleGroup | Where-Object {$_.Description -notlike "Membership in this role group is synchronized*" -or $null -eq $_.WellKnownObject}
	$exoRoleGroups = Get-RoleGroup | Select-Object  -Property Name,Identity,@{n="AssigneeType";e={"RoleGroup"}},@{n="User";e={""}}
	$directAssignments = Get-ManagementRoleAssignment | Where-Object {$_.RoleAssigneeType -eq 'User' -or $_.RoleAssigneeType -eq 'SecurityGroup'} | Select-Object -Property Name,Identity,@{n="AssigneeType";e={$_.RoleAssigneeType}},User
	$exoRoleAssignments = $exoRoleGroups + $directAssignments

	foreach ($rm in $exoRoleAssignments) {
		$roleUsers = @()
		
		# Get the members
	    if ($rm.AssigneeType -eq 'RoleGroup') {
			# Type is Exchange role group
			Write-Verbose -Message "Processing role group $($rm.Name)"
			$roleUsers += Get-ExoRoleGroupMembers -roleGroup $rm -roleName $rm.Name
			}
		elseif ($rm.AssigneeType -eq 'SecurityGroup') {
			# Type is Exchange mail-enabled security group
			Write-Verbose -Message "Processing role group $($rm.Name)"
			$roleUsers += Get-ExoSecurityGroupMembers -group (Get-DistributionGroup -Identity $rm.User) -roleName $rm.Name
		}
		else {
			# Type is user
			Write-Verbose -Message "Processing role $($rm.Name)"
			Write-Verbose -Message "User directly assigned $($rm.Name) role: $($rm.User)"
			$roleUsers += New-Object -TypeName PSObject -Property @{
				Id = @((Get-User -Identity $rm.User).ExternalDirectoryObjectId)[0]
				ParentGroup = ""
			}
		}
		
	    # Iterate each user
	    foreach ($user in $roleUsers) {
			
	        # Get underlying MSOL user details
	        $mUser = Get-UserDetails -id $user.Id
					
			# Add to final object
	        $pUsers += New-Object -TypeName PSObject -Property @{
	            SignInName = $user.ParentGroup + $mUser.SignInName
	            PasswordAge = $mUser.PasswordAge
	            Role = $rm.Name
	            MFADefault = $mUser.MFADefault
	            MFAPhone = $mUser.MFAPhone
	            UserType = $mUser.UserType
				Workload = 'Exchange Online'
	        }
		}
	}
}

if ($pUsers.Count -gt 0) {
	# Write the report

	$Report = "<html><head><link rel='stylesheet' href='https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0-beta.3/css/bootstrap.min.css' integrity='sha384-Zug+QiDoJOrZ5t4lssLdxGhVrurbmBWopoEl+M6BdEfwnCJZtKxi1KgxUyJq13dy' crossorigin='anonymous'></head><body>"
	$Report += "<div class='jumbotron jumbotron-fluid'>
	<div class='container'>
	  <h1 class='display-4'>Office 365 Admin Role Assignment Report</h1>
	  <p class='lead'>This report contains details of accounts assigned an administrative role in the following Office 365 workloads: $($workLoads -join ', '). If an assignment is via group membership, the sign-in name is prefixed with the name of the group.<br>Generated on $((Get-Date).ToLocalTime())</p>
	</div>
	</div>"

	$workloadGrouping = $pUsers | Group-Object -Property Workload
	foreach ($w in $workloadGrouping) {
	
	$Report+= "<div class='card'>
	    <h3 class='card-header'>
	      Workload: $($w.Name)
	    </h3>"
	
	$roleGrouping = $w.Group | Group-Object -Property	Role
	ForEach($r in $RoleGrouping) {
	    if($IgnoredRoles -notcontains $r.Name) {
	    $Report += "<div class='card'>
	    <div class='card-header'>
	      Role: $($r.Name)
	    </div>"

	    $Report += "<div class='card-body'><table class='table'>
	    <thead>
	      <tr>
	        <th>Sign In Name</th>
	        <th>Type</th>
	        <th>Password Age</th>
	        <th>MFA Default</th>
	        <th>MFA Phone</th>
	      </tr>
	    </thead>
	    <tbody>"

	    ForEach($u in $r.Group) {
	        $Report += "<tr>"
	        $Report += "<td>$($u.SignInName)</td>"
	        $Report += "<td>$($u.UserType)</td>"
	        
	        If($u.PasswordAge -ge $PasswordAgeThreshold) { $Class = "table-danger" } else { $Class = "table-success" }
	        $Report += "<td class='$Class'>$($u.PasswordAge) Days</td>"

	        If($u.MFADefault -eq $null) { $Class = "table-danger" } else { $Class = "table-success" }
	        $Report += "<td class='$Class'>$($u.MFADefault)</td>"

	        $Report += "<td>$($u.MFAPhone)</td>"
	        $Report += "</tr>"

	    }
	  

	    $Report += "</tbody></table></div>
	  </div>"
	}
	}
	}

	$Report | Out-File $Output
	Invoke-Item $Output
}
