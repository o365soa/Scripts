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
        This script will retrieve a list of users who have not signed in for at lease a specified number of days.
        Note: The "Office 365 Security Optimization Assessment" enterprise application must exist 
        in Entra ID for this script to work.

    .PARAMETER SignInType
        Filter users on interactive or non-interactive sign-ins. Valid value is Interactive or NonInterctive.
        Interactive is the default.

    .PARAMETER DaysOfInactivity
        The number of days of sign-in inactivity for the user to be returned. Default value is 30.
    
    .PARAMETER Environment
        Cloud environment of the tenant. Possible values are Commercial, USGovGCC, USGovGCCHigh, USGovDoD, Germany, and China.
        Default value is Commercial.

    .PARAMETER DoNotExportToCSV
        Switch to skip exporting the results to CSV and instead output the result objects to the host.
        
	.NOTES
        Version 1.4
        January 30, 2024

	.LINK
		about_functions_advanced   
#>
#Requires -Modules @{ModuleName = 'Microsoft.Graph.Authentication'; ModuleVersion = '2.0.0'}
#Requires -Modules @{ModuleName = 'Microsoft.Graph.Applications'; ModuleVersion = '2.0.0'}
[CmdletBinding()]
param (
    [ValidateSet('Interactive','NonInteractive')]$SignInType = 'Interactive',
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
Write-Host -ForegroundColor Green "$(Get-Date) Connecting to Microsoft Graph. If not already connected, use an account that has permission to create a client secret for the SOA enterprise application..."
Connect-MgGraph -ContextScope CurrentUser -Scopes 'Application.ReadWrite.All','User.Read' -Environment $cloud -NoWelcome

# Use MSAL in the Graph.Authentication module

$GraphAuthModule = Get-Module -Name Microsoft.Graph.Authentication -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
# Support for the .Net Core version of the library. Variable doesn't exist in PowerShell v4 and below, 
# so if it doesn't exist it is assumed that 'Desktop' edition is used
if ($PSEdition -eq 'Core'){$Folder = 'Core'} else {$Folder = 'Desktop'}

$MSALPath = Join-Path $GraphAuthModule.ModuleBase "Dependencies\$($Folder)\Microsoft.Identity.Client.dll"
Write-Verbose "$(Get-Date) Loading module from $MSALPath"
try {
    Add-Type -LiteralPath $MSALPath | Out-Null
} 
catch {
    Write-Error -Message "Unable to load Microsoft Authentication Library from $MSALPath"
    exit
}

# Get the Microsot Entra ID application and create a secret
Write-Host -ForegroundColor Green "$(Get-Date) Creating a new client secret for the SOA application..."
try {
    $graphApp = Get-MgApplication -Filter "displayName eq 'Office 365 Security Optimization Assessment'" -ErrorAction Stop | Where-Object {$_.Web.RedirectUris -Contains "https://security.optimization.assessment.local"}
}
catch {
    Write-Error -Message "This script requires the Microsoft Entra ID application created for the Security Optimization Assessment. Run `"Install-Module SOA`" then `"Install-SOAPrerequisites -AzureADAppOnly`" to provision the application."
    exit
}
$params = @{
    displayName = "Get Inactive Users on $(Get-Date -Format "dd-MMM-yyyy")"
    endDateTime = (Get-Date).AddDays(2)
}
try {
    $secret = Add-MgApplicationPassword -ApplicationId $graphApp.Id -PasswordCredential $params -ErrorAction Stop
}
catch {
    Write-Error -Message "Unable to create client secret for the SOA application in Entra ID."
}

# Wait for any replication latency
Write-Host -ForegroundColor Green "$(Get-Date) Sleeping to allow for client secret replication latency..."
Start-sleep -Seconds 5

# Get tenant domain via delegated call
try {
    $tenantDomain = ((Invoke-MgGraphRequest GET "/v1.0/organization" -OutputType PSObject -ErrorAction Stop).Value | Select-Object -ExpandProperty VerifiedDomains | Where-Object { $_.isInitial }).Name
}
catch {
    Write-Error -Message "Unable to get the tenant's default routing domain."
    exit
}
# Set endpoints according to environment
switch ($Environment) {
    "Commercial"   {$authority = "https://login.microsoftonline.com/$tenantDomain";$resource = "https://graph.microsoft.com/"}
    "USGovGCC"     {$authority = "https://login.microsoftonline.com/$tenantDomain";$resource = "https://graph.microsoft.com/"}
    "USGovGCCHigh" {$authority = "https://login.microsoftonline.us/$tenantDomain";$resource = "https://graph.microsoft.us/"}
    "USGovDoD"     {$authority = "https://login.microsoftonline.us/$tenantDomain";$resource = "https://dod-graph.microsoft.us/"}
    "Germany"      {$authority = "https://login.microsoftonline.de/$tenantDomain";$resource = "https://graph.microsoft.de/"}
    "China"        {$authority = "https://login.partner.microsoftonline.cn/$tenantDomain";$resource = "https://microsoftgraph.chinacloudapi.cn/"}
}

# Get a token
$ccApp = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create($GraphApp.AppId).WithClientSecret($Secret.SecretText).WithAuthority($Authority).WithLegacyCacheCompatibility($false).Build()
$scopes = New-Object System.Collections.Generic.List[string]
$scopes.Add("$($resource)/.default")
$attemptCount = 1
do {
    Write-Verbose "$(Get-Date) Getting an access token"
    try {
        $mgToken = $ccApp.AcquireTokenForClient($scopes).ExecuteAsync().GetAwaiter().GetResult()
        if ($mgToken){Write-Verbose "$(Get-Date) Successfully got a token using MSAL for $($resource)"}
    }
    catch {
        Write-Verbose "Failed to get an access token on attempt number $attemptCount."
        $attemptCount++
        Start-Sleep -Seconds 5
    }
}
until ($mgToken -or $attemptCount -eq 24)
if ($null -eq $mgToken) {
    Write-Error -Message "Unable to get an access token within the timeout period of two minutes."
    exit
}
else {
    Disconnect-MgGraph | Out-Null
    $attemptCount = 1
    # Reconnect to Graph as SOA application
    Write-Verbose "Connecting to Graph as the SOA application..."
    $ssCred = $secret.SecretText | ConvertTo-SecureString -AsPlainText -Force
    $graphCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ($graphApp.AppId), $ssCred
    do {
        try {
            Connect-MgGraph -TenantId $tenantDomain -ClientSecretCredential $graphCred -Environment $cloud -ContextScope "Process" -ErrorAction Stop | Out-Null
        }
        catch {
            Write-Verbose "Failed to connected to Graph as the SOA application on attempt number $attemptCount."
            $attemptCount++
            Start-Sleep -Seconds 5
        }
    }
    until ((Get-MgContext) -or $attemptCount -eq 24)
    if ($null -eq (Get-MgContext)) {
        Write-Error -Message "Unable to connect Graph as the SOA application within the timeout period of two minutes."
        Write-Error $_
        exit
    }
}

# Get Graph data and continue paging until data collection is complete
Write-Host -ForegroundColor Green "$(Get-Date) Getting users based on inactivity..."
$targetdate = (Get-Date).ToUniversalTime().AddDays(-$DaysOfInactivity).ToString("o")

$result = @()
if ($SignInType -eq 'Interactive') {
    $siFilter = 'signInActivity/lastSignInDateTime'
}
else {
    $siFilter = 'signInActivity/lastNonInteractiveSignInDateTime'
}
$apiUrl = "$($resource)beta/users?`$filter=$siFilter lt $($targetdate)&`$select=accountEnabled,id,userType,signInActivity,userprincipalname"
Write-Verbose "Initial URL: $apiUrl"

do {
    $response = Invoke-MgGraphRequest -Method GET $apiUrl -OutputType PSObject
    $apiUrl = $($response."@odata.nextLink")
    if ($apiUrl) { Write-Verbose "@odata.nextLink: $apiUrl" }
    $result += $response.Value
}
until ($null -eq $response."@odata.nextLink")

# Processing user data to prepare export
Write-Host -ForegroundColor Green "$(Get-Date) Processing returned users..."

$return=@()
foreach ($item in $result) {
    if ($null -ne $item.userPrincipalName -and $item.accountEnabled -eq $true) {
        $return += New-Object -TypeName PSObject -Property @{
            UserPrincipalName=$item.userprincipalname
            LastInteractiveSignIn=$item.signinactivity.lastsignindatetime
            LastNonInteractiveSignInDateTime=$item.signinactivity.lastNonInteractiveSignInDateTime
            UserType=$item.usertype
        }
    }
}

# Export to CSV unless opted out
if ($DoNotExportToCSV -eq $false) {
    Write-Host -ForegroundColor Green "$(Get-Date) Exporting AAD-InactiveUsers.csv in current directory..."  
    $return | Select-Object -Property UserPrincipalName,UserType,LastInteractiveSignIn,LastNonInteractiveSignInDateTime | Export-CSV "AAD-InactiveUsers.csv" -NoTypeInformation
}

# Remove client secret
Write-Host -ForegroundColor Green "$(Get-Date) Removing any client secrets for the SOA application. Select the user from the first connection..."
Disconnect-MgGraph | Out-Null
try {
    Connect-MgGraph -ContextScope Process -Scopes 'Application.ReadWrite.All','User.Read' -Environment $cloud -NoWelcome
}
catch {
    Write-Error -Message "Unable to reconnect to Graph as a user to remove the client secret for the SOA application."
}
if (Get-MgContext) {
    $secrets = (Get-MgApplication -Filter "displayName eq 'Office 365 Security Optimization Assessment'").PasswordCredentials
    foreach ($secret in $secrets) {
        # Suppress errors in case a secret no longer exists
        try {
            Remove-MgApplicationPassword -ApplicationId $GraphApp.Id -BodyParameter (@{KeyID = $Secret.KeyId})
        }
        catch {}
    }
}

if ($DoNotExportToCSV -eq $true) {
    $return
}

Write-Host -ForegroundColor Green "$(Get-Date) Script has completed."
# Stop-Transcript
