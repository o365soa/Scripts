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
		Get users via Microsoft Graph based on sign-in activity

	.DESCRIPTION
        This script will retrieve a list of users who have not signed in for at least a specified number of days.
        Requires Microsoft.Graph.Authentication module and delegated scope of User.Read.All.

    .PARAMETER SignInType
        Filter users on the type of sign-in: interactive (successful or unsuccessful), non-interactive (successful or unsuccessful),
        or successful (for either type). Valid values are Interactive, NonInteractive, and Successful.
        Successful is the default.

    .PARAMETER DaysOfInactivity
        The number of days of sign-in inactivity for the user to be returned. Default value is 30.
        Note: Users with a null value for the date/time of the sign-in type will not be returned.
    
    .PARAMETER Environment
        Cloud environment of the tenant. Possible values are Commercial, USGovGCC, USGovGCCHigh, USGovDoD, Germany, and China.
        Default value is Commercial.

    .PARAMETER DoNotExportToCSV
        Switch to skip exporting the results to CSV and instead output the result objects to the host.
        
	.NOTES
        Version 1.4
        May 23, 2024

	.LINK
		about_functions_advanced   
#>
#Requires -Modules @{ModuleName = 'Microsoft.Graph.Authentication'; ModuleVersion = '2.0.0'}
[CmdletBinding()]
param (
    [ValidateSet('Interactive','NonInteractive','Successful')]$SignInType = 'Successful',
    [int]$DaysOfInactivity = 30,
    [ValidateSet("Commercial", "USGovGCC", "USGovGCCHigh", "USGovDoD", "Germany", "China")][string]$Environment="Commercial",
    [switch]$DoNotExportToCSV
)

# Start-Transcript -Path "Transcript-inactiveusers.txt" -Append
switch ($Environment) {
    "Commercial"   {$cloud = "Global"}
    "USGovGCC"     {$cloud = "Global"}
    "USGovGCCHigh" {$cloud = "USGov"}
    "USGovDoD"     {$cloud = "USGovDoD"}
    "Germany"      {$cloud = "Germany"}
    "China"        {$cloud = "China"}            
}

# Supported scope from least to most privileged
$supportedScopes = @('User.Read.All', 'User.ReadWrite.All', 'Directory.Read.All', 'Directory.ReadWrite.All')
foreach ($scope in (Get-MgContext).Scopes) {
    if ($scope -in $supportedScopes) {
        $scopeInCurrentContext = $true
        break
    }
}
if (-not($scopeInCurrentContext)) {
    Write-Host -ForegroundColor Green "$(Get-Date) Connecting to Microsoft Graph..."
    Connect-MgGraph -ContextScope CurrentUser -Scopes 'User.Read.All', -Environment $cloud -NoWelcome
}

$targetdate = (Get-Date).ToUniversalTime().AddDays(-$DaysOfInactivity).ToString("o")
$result = @()
switch ($SignInType) {
    Interactive {$siFilter = 'signInActivity/lastSignInDateTime'}
    NonInteractive {$siFilter = 'signInActivity/lastNonInteractiveSignInDateTime'}
    Successful {$siFilter = 'signInActivity/lastSuccessfulSignInDateTime'}
}

# beta endpoint is required for the lastSuccessfulSignInDateTime property
$apiUrl = "/beta/users?`$filter=$siFilter lt $($targetdate)&`$select=accountEnabled,id,userType,signInActivity,userprincipalname"
Write-Verbose "Initial URL: $apiUrl"
Write-Host -ForegroundColor Green "$(Get-Date) Getting users based on $DaysOfInactivity days of inactivity..."
do {
    # Get data via Graph and continue paging until complete
    $response = Invoke-MgGraphRequest -Method GET $apiUrl -OutputType PSObject
    $apiUrl = $($response."@odata.nextLink")
    if ($apiUrl) { Write-Verbose "@odata.nextLink: $apiUrl" }
    $result += $response.Value
}
until ($null -eq $response."@odata.nextLink")

if ($result.Count -gt 0) {
    # Processing user data to prepare export
    Write-Host -ForegroundColor Green "$(Get-Date) Processing $($result.Count) returned users..."

    $return=@()
    foreach ($item in $result) {
        if ($null -ne $item.userPrincipalName -and $item.accountEnabled -eq $true) {
            $return += New-Object -TypeName PSObject -Property @{
                UserPrincipalName = $item.userprincipalname
                LastSuccessfulSignIn = $item.signinactivity.lastSuccessfulSignInDateTime
                LastInteractiveSignIn = $item.signinactivity.lastsignindatetime
                LastNonInteractiveSignIn = $item.signinactivity.lastNonInteractiveSignInDateTime
                UserType = $item.usertype
            }
        }
    }

    # Export to CSV unless opted out
    if ($DoNotExportToCSV -eq $false) {
        Write-Host -ForegroundColor Green "$(Get-Date) Exporting EntraID-InactiveUsers.csv in current directory..."  
        $return | Select-Object -Property UserPrincipalName,UserType,LastSuccessfulSignIn,LastInteractiveSignIn,LastNonInteractiveSignIn | Export-CSV "EntraID-InactiveUsers.csv" -NoTypeInformation
    }

    if ($DoNotExportToCSV -eq $true) {
        $return
    }
} else {
    Write-Host -ForegroundColor Green "$(Get-Date) No users were returned based on the search criteria."
}

Write-Host -ForegroundColor Green "$(Get-Date) Script has completed."
# Stop-Transcript
