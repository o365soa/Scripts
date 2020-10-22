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
        
	.NOTES
		Emily Coates
		Customer Engineer - Microsoft
		emily.coates@microsoft.com

	.LINK
		about_functions_advanced   
#>

Start-Transcript -Path "Transcript-inactiveusers.txt" -Append

    # Connect to Azure AD. This is required to get the access token.    
    
    Write-Host -ForegroundColor Green "$(Get-Date) Connecting to Azure AD. Use an administrative account with the ability to manage Azure AD Applications"
    
    Import-Module AzureADPreview
    Connect-AzureAD
    
    # Get the AzureAD Application and create a secret
    
    Write-Host -ForegroundColor Green "$(Get-Date) Creating a new 48 hour secret for the SOA application"
    
    $GraphApp = Get-AzureADApplication -Filter "displayName eq 'Office 365 Security Optimization Assessment'"
    $clientID = $GraphApp.AppId
    $secret = New-AzureADApplicationPasswordCredential -ObjectId $GraphApp.ObjectId -EndDate (Get-Date).AddDays(2) -CustomKeyIdentifier "$(Get-Date -Format "dd-MMM-yyyy")"
    
    # Let the secret settle
    
    Write-Host -ForegroundColor Green "$(Get-Date) Sleeping for 60 seconds to let the Graph secret settle. This prevents a race condition"
    Start-sleep 60
    
    # Use ADAL Library - requires AzureAD module to be installed
    
    $AadModule = Get-Module -Name "AzureADPreview" -ListAvailable
    
    $Latest_Version = ($AadModule | Select-Object version | Sort-Object)[-1]
    $aadModule      = $AadModule | Where-Object { $_.version -eq $Latest_Version.version }
    $adal           = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
    $adalforms      = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"
    
    [System.Reflection.Assembly]::LoadFrom($adal) | Out-Null
    [System.Reflection.Assembly]::LoadFrom($adalforms) | Out-Null
    
    # Get a token

    Write-Host -ForegroundColor Green "$(Get-Date) Getting an access token"
    
    $GraphAppDomain     = (Get-AzureADTenantDetail | Select-Object -ExpandProperty VerifiedDomains | Where-Object { $_.Initial }).Name
    $authority          = "https://login.microsoftonline.com/$graphappdomain"
    $authContext        = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
    $resource           = "https://graph.microsoft.com"
    
    $authContext.TokenCache.Clear()
    
    $ClientCredential   = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.ClientCredential" -ArgumentList @($clientID,$secret.value)
    $authResult         = $authContext.AcquireTokenAsync($Resource,$ClientCredential)
    
    # Set headers
    
    $GraphAuthHeader = @{'Authorization'="$($authresult.result.AccessTokenType) $($authresult.result.AccessToken)"}

# Get Graph data and continue paging until data collection is complete

Write-Host -ForegroundColor Green "$(Get-Date) Get Graph data and continue paging until data collection is complete"

    $targetdate = (Get-Date).AddDays(-30)
    $targetdatestr = $targetdate.ToString("yyyy-MM-dd")

    $Result = @()
    $ApiUrl = "https://graph.microsoft.com/beta/users?`$filter=signInActivity/lastSignInDateTime lt $targetdatestr&`$select=accountEnabled,id,userType,signInActivity,userprincipalname"
    $Response = Invoke-RestMethod -Headers $GraphAuthHeader -Uri $ApiUrl -Method Get
    $Users = $Response.value
    $Result = $Users
    
    While ($Response.'@odata.nextLink' -ne $null) {
    $Response = Invoke-RestMethod -Headers $GraphAuthHeader -Uri $Response.'@odata.nextLink' -Method Get
    $Users = $Response.value
    $Result += $Users
    }

# Processing user data to prepare export

Write-Host -ForegroundColor Green "$(Get-Date) Processing user data to prepare export"

    $return=@()

    foreach ($item in $result)
    {
        $Return += New-Object -TypeName PSObject -Property @{
            UserPrincipalName=$item.userprincipalname
            AccountEnabled=$item.accountenabled
            LastSignIn=$item.signinactivity.lastsignindatetime
            UserType=$item.usertype
        }
    }

# Exporting CSV

Write-Host -ForegroundColor Green "$(Get-Date) Exporting AAD-InactiveUsers.csv in current directory"  

    $return | Export-CSV AAD-InactiveUsers.csv -NoTypeInformation

Write-Host ("")

Stop-Transcript
