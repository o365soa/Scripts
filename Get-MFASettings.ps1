#Requires -Version 4

<#
	.SYNOPSIS
		This script exports all user MFA settings

	.DESCRIPTION
		Iterates through users and reports on MFA settings.

        	Requires Version 1 of Azure AD PowerShell (MSOL)

	.EXAMPLE
		PS C:\> .\Get-MFASettings.ps1

	.NOTES
		Cam Murray
		Field Engineer - Microsoft
		cam.murray@microsoft.com
		
		For updates, and more scripts, visit https://github.com/O365AES/Scripts
		
		Last update: 10 April 2017

	.LINK
		about_functions_advanced

#>

$MFAUsers = @()

ForEach($User in (Get-MsolUser -All)) {

    $Default = $null
    $MFAState = $null

    $Default = ($User.StrongAuthenticationMethods | Where-Object {$_.IsDefault -eq $true}).MethodType

    If($($User.StrongAuthenticationRequirements.State) -eq $null) {

        $MFAState = "Not Set"

    } else {

        $MFAState = $($User.StrongAuthenticationRequirements.State)

    }

    $MFAUsers += New-Object -TypeName psobject -Property @{
		Member=$($User.UserPrincipalName)
		UserMFAState=$MFAState
                Phone=$($User.StrongAuthenticationUserDetails.PhoneNumber)
                Default=$Default
    }
}

return $MFAUsers
