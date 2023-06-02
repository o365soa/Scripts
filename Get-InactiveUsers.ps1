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
		Inactive users standalone script

	.DESCRIPTION
        This script is designed to collect a list of users who have not signed in for 30 days or more.
        The Office 365: Security Optimization Assessment Azure AD application must exist 
        for this to function.

    .PARAMETER SignInType
        Filter users on interactive or non-interactive sign-ins. Valid value is Interactive or NonInterctive.
        Interactive is the default.
        
	.NOTES
        Version 1.3
        02 June, 2023

	.LINK
		about_functions_advanced   
#>
#Requires -Modules ExchangeOnlineManagement, Microsoft.Graph.Authentication, Microsoft.Graph.Applications
[CmdletBinding()]
param (
    [ValidateSet('Interactive','NonInteractive')]$SignInType = 'Interactive'
)

Start-Transcript -Path "Transcript-inactiveusers.txt" -Append

Connect-MgGraph -ContextScope Process -Scopes 'Application.ReadWrite.All','User.Read'

# Get the AzureAD Application and create a secret

Write-Host -ForegroundColor Green "$(Get-Date) Creating a new client secret for the SOA application."

$GraphApp = Get-MgApplication -Filter "displayName eq 'Office 365 Security Optimization Assessment'" | Where-Object {$_.Web.RedirectUris -Contains "https://security.optimization.assessment.local"}

$Params = @{
    displayName = "Task on $(Get-Date -Format "dd-MMM-yyyy")"
    endDateTime = (Get-Date).AddDays(2)
}

$Secret = Add-MgApplicationPassword -ApplicationId $GraphApp.Id -PasswordCredential $Params

# Let the secret settle

Write-Host -ForegroundColor Green "$(Get-Date) Sleeping for 60 seconds to let the client secret replicate."
Start-sleep 60

# Find a suitable MSAL library - Requires that the ExchangeOnlineManagement module is installed
$ExoModule = Get-Module -Name "ExchangeOnlineManagement" -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1

# Add support for the .Net Core version of the library. Variable doesn't exist in PowerShell v4 and below, 
# so if it doesn't exist it is assumed that 'Desktop' edition is used
If ($PSEdition -eq 'Core'){
    $Folder = "netCore"
} Else {
    $Folder = "NetFramework"
}

$MSAL = Join-Path $ExoModule.ModuleBase "$($Folder)\Microsoft.Identity.Client.dll"
Write-Verbose "$(Get-Date) Loading module from $MSAL"
Try {Add-Type -LiteralPath $MSAL | Out-Null} Catch {} # Load the MSAL library

# Get a token

Write-Host -ForegroundColor Green "$(Get-Date) Getting an access token"

$GraphAppDomain = ((Invoke-MgGraphRequest GET "/v1.0/organization" -OutputType PSObject).Value | Select-Object -ExpandProperty VerifiedDomains | Where-Object { $_.isInitial }).Name
$authority          = "https://login.microsoftonline.com/$GraphAppDomain"
$resource           = "https://graph.microsoft.com"
    
$ccApp = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create($GraphApp.AppId).WithClientSecret($Secret.SecretText).WithAuthority($Authority).Build()

$Scopes = New-Object System.Collections.Generic.List[string]
$Scopes.Add("$($Resource)/.default")

$MgToken = $ccApp.AcquireTokenForClient($Scopes).ExecuteAsync().GetAwaiter().GetResult()
If ($MgToken){Write-Verbose "$(Get-Date) Successfully got a token using MSAL for $($Resource)"}

# Get Graph data and continue paging until data collection is complete

Write-Host -ForegroundColor Green "$(Get-Date) Get Graph data and continue paging until data collection is complete"

$targetdate = (Get-Date).ToUniversalTime().AddDays(-30).ToString("o")

$Result = @()
if ($SignInType -eq 'Interactive') {
    $siFilter = 'signInActivity/lastSignInDateTime'
}
else {
    $siFilter = 'signInActivity/lastNonInteractiveSignInDateTime'
}
$ApiUrl = "https://graph.microsoft.com/beta/users?`$filter=$siFilter lt $($targetdate)&`$select=accountEnabled,id,userType,signInActivity,userprincipalname"

Do {
    $Response = Invoke-MgGraphRequest -Method GET $ApiUrl -Authentication UserProvidedToken -Token ($MgToken.AccessToken | ConvertTo-SecureString -AsPlainText -Force) -OutputType PSObject
    $ApiUrl = $($Response."@odata.nextLink")
    If ($ApiUrl) { Write-Verbose "@odata.nextLink: $ApiUrl" }
    
    $Result = $Response.Value

} Until ($Null -eq $Response."@odata.nextLink")

# Processing user data to prepare export

Write-Host -ForegroundColor Green "$(Get-Date) Processing user data to prepare export"

$return=@()

foreach ($item in $result) {
    if ( $null -ne $item.userPrincipalName -and $item.accountEnabled -eq $true) {
        $Return += New-Object -TypeName PSObject -Property @{
            UserPrincipalName=$item.userprincipalname
            LastInteractiveSignIn=$item.signinactivity.lastsignindatetime
            LastNonInteractiveSignInDateTime=$item.signinactivity.lastNonInteractiveSignInDateTime
            UserType=$item.usertype
        }
    }
}

# Exporting CSV

Write-Host -ForegroundColor Green "$(Get-Date) Exporting AAD-InactiveUsers.csv in current directory."  

$return | Select-Object -Property UserPrincipalName,UserType,LastInteractiveSignIn,LastNonInteractiveSignInDateTime | Export-CSV "AAD-InactiveUsers.csv" -NoTypeInformation

# Remove client secret
Write-Host -ForegroundColor Green "$(Get-Date) Removing client secrets for the SOA application."
$Secrets = (Get-MgApplication -Filter "displayName eq 'Office 365 Security Optimization Assessment'").PasswordCredentials
foreach ($Secret in $Secrets) {
    # Suppress errors in case a secret no longer exists
    try {
        Remove-MgApplicationPassword -ApplicationId $GraphApp.Id -BodyParameter (@{KeyID = $Secret.KeyId})
    }
    catch {}
}

Stop-Transcript
