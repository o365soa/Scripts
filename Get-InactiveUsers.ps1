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
    
    $GraphAppDomain     = (Get-AzureADTenantDetail | Select-Object -ExpandProperty VerifiedDomains | Where-Object { $_.Initial }).Name
    $authority          = "https://login.microsoftonline.com/$graphappdomain"
    $resource           = "https://graph.microsoft.com"
        
    $ccApp = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create($ClientID).WithClientSecret($Secret.Value).WithAuthority($Authority).Build()

    $Scopes = New-Object System.Collections.Generic.List[string]
    $Scopes.Add("$($Resource)/.default")

    $Token = $ccApp.AcquireTokenForClient($Scopes).ExecuteAsync().GetAwaiter().GetResult()
    If ($Token){Write-Verbose "$(Get-Date) Successfully got a token using MSAL for $($Resource)"}

    # Set authentication headers
    $GraphAuthHeader = @{'Authorization'="$($Token.TokenType) $($Token.AccessToken)"}

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
	if ( $item.userPrincipalName -ne $null )
	{
		$Return += New-Object -TypeName PSObject -Property @{
			UserPrincipalName=$item.userprincipalname
			AccountEnabled=$item.accountenabled
			LastSignIn=$item.signinactivity.lastsignindatetime
			UserType=$item.usertype
		}
	}
    }

# Exporting CSV

Write-Host -ForegroundColor Green "$(Get-Date) Exporting AAD-InactiveUsers.csv in current directory"  

    $return | Export-CSV AAD-InactiveUsers.csv -NoTypeInformation

Write-Host ("")

Stop-Transcript
