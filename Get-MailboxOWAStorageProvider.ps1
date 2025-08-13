#requires -Modules ExchangeOnlineManagement
##############################################################################################
#This sample script is not supported under any Microsoft standard support program or service.
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
	.Synopsis
		Get the configured third-party storage providers for a given mailbox.
	.Description
		Using Exchange Web Services, any third-party storage providers configured for a mailbox (which can be
		configured via Outlook on the web) will be retrieved, including the account name configured
		for a given provider. Supports client secret and certificate authentication.

		* Important: Requires the EWS Managed API to be installed on the local machine.
			1. Run PowerShell as Administrator
			2. Run the following: Install-Package Microsoft.Exchange.Webservices
			   Note: If NuGet is not a registered package source:
			   https://learn.microsoft.com/en-us/powershell/gallery/powershellget/supported-repositories
		* Important: Requires an Entra ID app registration configured for app-only authentication
		  with EWS.AccessAsApp (full_access_as_app).
		    Details for registering an app and adding the role to the manifest:
		    https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth#register-your-application
			Note: Enter the application ID and tenant's default routing domain in the variables at the top of the begin block.
		* Important: Requires the corresponding enterprise application (service principal) to have a role assignment
		  with the "Application EWS.AccessAsApp" role (and a scope that includes the desired mailboxes).
			Details for creating the service principal link and management role assignment in EXO:
			https://learn.microsoft.com/en-us/exchange/permissions-exo/application-rbac
	.Parameter EmailAddress
		Email address of the mailbox from which to retrieve the configuration. Supports pipeline input
		of email addresses or objects with an EmailAddress or PrimarySMTPAddress property, such as
		with Get-Mailbox.
	.Parameter Cloud
		Office 365 environment which hosts the mailboxes. Valid values are Commercial, USGovGCC, China.
		Default value is Commercial. The feature is not available in GCC High and DoD. Unknown if available in China.
	.Parameter CertificateAuthentication
		Use a certificate for authentication instead of a client secret. The app registration must have a
		certificate uploaded, installed on the local machine in Current User\Personal\Certificates,
		and the thumbprint specified in the variables region below.
	.Example
		.\Get-MailboxOWAStorageProvider johndoe@contoso.com
		Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize unlimited | .\Get-MailboxOWAStorageProvider
	.Notes
		Version: 1.2
		Date: August 13, 2025
#>

[CmdletBinding()]
param (
	[Parameter(Mandatory=$true,ValueFromPipelinebyPropertyName=$true,Position=0)][Alias('PrimarySMTPAddress')][string]$EmailAddress,
	[ValidateSet('Commercial','USGovGCC','China')][string]$Cloud = 'Commercial',
	[switch]$CertificateAuthentication
)

begin {
	# Variables
	$tenantDomain = 'tenantname.onmicrosoft.com' # Default routing domain of the tenant
	$appId = '00000000-0000-0000-0000-000000000000' # Application ID of the app registration in Entra ID with EWS permission
	$certThumbprint = '' # Thumbprint of the certificate to use for authentication if CertificateAuthentication switch is used
	# End variables

	if ($tenantDomain -like "tenantname*" -or $appId -like "00000000*") {
		Write-Error "The tenant domain or application ID has not been specified in the Variables section of the `"begin`" block."
		break
	}
	if ($CertificateAuthentication -and -not $certThumbprint) {
		Write-Error "The certificate thumbprint has not been specified in the Variables section of the `"begin`" block."
		break
	}

	# Check for EWS API installed via NuGet
	# If already loaded, save time by reusing the loaded type
	if (-not('Microsoft.Exchange.WebServices.Data.ExchangeVersion' -as [type])) {
		Write-Verbose 'EWS Managed API is not loaded. Checking if package is installed...'
		$apiPackage = Get-Package -Name Microsoft.Exchange.Webservices
		if ($apiPackage) {
			$dllName = 'microsoft.exchange.webservices.dll'
			try {
				Write-Verbose "Loading EWS Managed API from $((Get-Item $apiPackage.Source).DirectoryName)"
				Add-Type -Path (Join-Path -Path (Get-ChildItem -Path (Get-Item $apiPackage.Source).DirectoryName -Recurse -Filter $dllName | 
				Select-Object -ExpandProperty DirectoryName) -ChildPath $dllName) -ErrorAction Stop | Out-Null
			}
			catch {
				Write-Error $_
				break
			}
		}
		else {
			Write-Error 'The Exchange Web Services Managed API is not installed from NuGet and is required by this script.' -Category NotInstalled
			break
		}
	}
	else {
		Write-Verbose 'Using the EWS Managed API that is already loaded.'
	}

	# Import MSAL from Exchange Online module
	if (-not('Microsoft.Identity.Client.ConfidentialClientApplicationBuilder' -as [type])) {
		Write-Verbose 'Loading the MSAL from EXO module installation...'
		if ($PSEdition -eq 'Core') {$folder = 'netCore'} else {	$folder = 'NetFramework'}
		$ExoModule = Get-Module -Name ExchangeOnlineManagement -ListAvailable | Sort-Object -Property Version -Descending | Select-Object -First 1
		$MSAL = Join-Path -Path $ExoModule.ModuleBase -ChildPath "$($folder)\Microsoft.Identity.Client.dll"
		Add-Type -Path $MSAL | Out-Null
	}
	else {
		Write-Verbose 'Using the MSAL that is already loaded.'
	}

	switch ($Cloud) {
	    'Commercial'    { $base = 'https://login.microsoftonline.com/';$ewsUrl = 'https://outlook.office365.com'}
	    'USGovGCC'      { $base = 'https://login.microsoftonline.com/';$ewsUrl = 'https://outlook.office365.com'}
	    'China'         { $base = 'https://login.partner.microsoftonline.cn/';$ewsUrl = 'https://partner.outlook.cn'}
	}
	# Build client app and get access token
	$replyUri = $base + 'common/oauth2/nativeclient'
	$capabilities = New-Object System.Collections.Generic.List[string]
	# cp1 indicates support for CAE, which will result in an access token that is valid for 29 hours
	# (This helps collecting from all mailboxes in a large org without needing to include support for token expiration.)
	$capabilities.Add('cp1')
	if ($CertificateAuthentication) {
		# Use certificate authentication
		$cert = Get-Item "Cert:\CurrentUser\My\$certThumbprint"
		if (-not $cert) {
			Write-Error "The certificate with thumbprint $certThumbprint was not found in the CurrentUser\My store."
			break
		}
		$confidentialClient = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create($appId).WithRedirectUri($replyUri).WithCertificate($cert).WithTenantId($tenantDomain).WithClientCapabilities($capabilities).Build()
	}
	else {
		# Use client secret authentication
		$ssAppSecret = (Get-Credential -Message "Enter the app registration's client secret in the password field." -UserName "EWS Application").Password
		$bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ssAppSecret)
		$appSecret = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
		[System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
		$confidentialClient = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create($appId).WithRedirectUri($replyUri).WithClientSecret($appSecret).WithTenantId($tenantDomain).WithClientCapabilities($capabilities).Build()
	}
	$scope = New-Object System.Collections.Generic.List[string]
	$scope.Add("$ewsUrl/.default")
	$token = $confidentialClient.AcquireTokenForClient($scope).ExecuteAsync().GetAwaiter().GetResult()
	if (-not $CertificateAuthentication) {
		Remove-Variable -Name appSecret,ssAppSecret -ErrorAction SilentlyContinue
	}
}

process {
	if ($token.AccessToken) {
		Write-Progress -Activity 'Getting OWA storage provider settings' -CurrentOperation "Mailbox $EmailAddress"
		# Create EWS service object
		$exchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
		$exchangeService = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ExchangeService($exchangeVersion)
		$exchangeService.Credentials = New-Object -TypeName Microsoft.Exchange.WebServices.Data.OauthCredentials($token.AccessToken)
		$exchangeService.Url = $ewsUrl+'/EWS/Exchange.asmx'
		$exchangeService.ImpersonatedUserId = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,$EmailAddress)
		$exchangeService.HttpHeaders.Add('X-AnchorMailbox', $EmailAddress)
		$folderId = New-Object -TypeName Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$EmailAddress)

		# Get IPM.Configuration message for OWA attachment providers and the roaming dictionary property that contains the settings
		try {
			$userConfig = [Microsoft.Exchange.WebServices.Data.UserConfiguration]::Bind($exchangeservice, 'OWA.AttachmentDataProvider', $folderId, [Microsoft.Exchange.WebServices.Data.UserConfigurationProperties]::All)
			if ($userConfig) {
				# Convert byte array of the settings to string
				[xml]$xmlString = [System.Text.Encoding]::ASCII.GetString($userConfig.XmlData)
				foreach ($sp in $xmlString.AttachmentDataProvider.entry) {
					# Include in output only third-party providers
					if ($sp.isThirdPartyProvider -eq $true) {	
						New-Object -TypeName psobject -Property @{
						    Mailbox = $EmailAddress
						    ProviderName = $sp.DisplayName
						    ProviderAccount = $sp.associatedDataProviderAccountId
						}
					}
				}
			}
		}
		catch {
			if ($_.Exception.InnerException -notlike "*The configuration object was not found.*") {
				Write-Error $_
			}
		}
	}
}
end {
	Write-Progress -Activity 'Getting OWA storage provider settings' -Completed
}