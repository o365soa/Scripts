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
		Power Platform Dataverse auditing script

	.DESCRIPTION
        This script is designed to collect a list of Power Platform Dataverse environments
        which are not configured to audit all event types.
        The 'Office 365: Security Optimization Assessment' Azure AD application must exist 
        for this to retrieve valid access tokens.

    .PARAMETER EnableAuditing
        Configure the script to also remediate and configure Auditing to be enabled on all Dataverses.
        Default value is $False.

    .PARAMETER AsJson
        Configure the output file type to be JSON instead of CSV

    .EXAMPLE
        PS C:\> .\Get-DataverseAuditing.ps1 -EnableAuditing $True
        
	.NOTES
        Version 1.0
        21 June 2023

        Jonathan Devere-Ellery
        Cloud Solution Architect - Microsoft

	.LINK
		about_functions_advanced   
#>

#Requires -Version 5
#Requires -Modules @{ModuleName="ExchangeOnlineManagement"; ModuleVersion="3.0.0"}
#Requires -Modules Microsoft.PowerApps.Administration.PowerShell, Microsoft.Graph.Authentication, Microsoft.Graph.Applications


Param(
    [switch]$EnableAuditing,
    [switch]$AsJson
)

# Load MSAL
$ExoModule = Get-Module -Name "ExchangeOnlineManagement" -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
If ($PSEdition -eq 'Core'){
    $Folder = "netCore"
} Else {
    $Folder = "NetFramework"
}
$MSAL = Join-Path $ExoModule.ModuleBase "$($Folder)\Microsoft.Identity.Client.dll"
Try {Add-Type -LiteralPath $MSAL | Out-Null} Catch {}

Connect-MgGraph -ContextScope Process -Scopes 'Application.ReadWrite.All','User.Read'

$App = Get-MgApplication -Filter "displayName eq 'Office 365 Security Optimization Assessment'" | Where-Object {$_.Web.RedirectUris -Contains "https://security.optimization.assessment.local"}
$GraphAppDomain = ((Invoke-MgGraphRequest GET "/v1.0/organization" -OutputType PSObject).Value | Select-Object -ExpandProperty VerifiedDomains | Where-Object { $_.isInitial }).Name
$Authority = "https://login.microsoftonline.com/$GraphAppDomain"
$pubApp = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($App.AppId).WithRedirectUri('https://login.microsoftonline.com/common/oauth2/nativeclient').WithAuthority($Authority).Build()

Write-Host "Select account with Global or Dynamics 365 Administrator:"
$pubApp.AcquireTokenInteractive($Scopes).ExecuteAsync().GetAwaiter().GetResult() | Out-Null


$Result = @()
$account = $PubApp.GetAccountsAsync().Result.Username

$Environments = Get-AdminPowerAppEnvironment

foreach ($instance in $Environments) {
    $apiUrl = $null
    $apiUrl = $instance.Internal.properties.linkedEnvironmentMetadata.instanceApiUrl
    Write-Verbose "Environment: $($instance.DisplayName)"

    if ($apiUrl -and $instance.Internal.properties.linkedEnvironmentMetadata.instanceState -eq 'Ready') {
        $scope = New-Object System.Collections.Generic.List[string]
        $scope.Add("$apiUrl/.default")

        $token = $PubApp.AcquireTokenSilent($scope, $account).ExecuteAsync().GetAwaiter().GetResult()

        if ($token) {
            Write-Verbose "Successfully retrieved an access token"
            $headers = @{
                'Authorization' = "$($token.TokenType) $($token.AccessToken)"
                'Accept' = 'application/json'
                'OData-MaxVersion' = '4.0'
                'OData-Version' = '4.0'
                'If-None-Match' = 'null'
            }

            $instVer = [version]$instance.Internal.properties.linkedEnvironmentMetadata.version
            $verStr = "v" + $instVer.Major.ToString() + "." + $instVer.Minor.ToString()

            $response = Invoke-RestMethod -Uri "$apiUrl/api/data/$verStr/organizations?`$select=organizationid,isauditenabled,auditretentionperiodv2,isuseraccessauditenabled,isreadauditenabled" -Headers $headers

            $OrgID = $response.value.organizationid

            $result += New-Object -TypeName psobject -Property @{
                OrgID = $OrgID
                EnvDisplayName = $instance.DisplayName
                EnvState = $instance.Internal.properties.linkedEnvironmentMetadata.instanceState
                IsAuditEnabled = $response.value.isauditenabled
                IsAccessAuditEnabled = $response.value.isuseraccessauditenabled
                IsReadLogsEnabled = $response.value.isreadauditenabled
                RetentionPeriod = $response.value.auditretentionperiodv2
            }

            If ($EnableAuditing) {
                $Headers = @{
                    'Authorization'="$($token.TokenType) $($token.AccessToken)"
                    'Content-Type' = 'application/json'
                    'OData-MaxVersion' = '4.0'
                    'OData-Version' = '4.0'
                    'If-Match' = '*'
                }

                $Body = @{
                    'isauditenabled' = $True
                    'isuseraccessauditenabled' = $True
                    'isreadauditenabled' = $True
                }
                Write-Verbose "Setting Auditing for $apiUrl"
                Write-Verbose "OrgID: $OrgID"

                Try {
                    Invoke-RestMethod -Method Patch -Uri "$apiUrl/api/data/$verStr/organizations($OrgID)" -Headers $Headers -Body ($body | ConvertTo-Json)
                } Catch {
                    Write-Warning "Error while making PATCH request to $apiUrl"
                }
            }
        }
    }
}

If ($AsJson) {
    $Result | ConvertTo-Json -Depth 10 | Out-File -FilePath "Dataverse-Auditing.json"
} Else {
    $Result | Export-Csv -NoTypeInformation -Path "Dataverse-Auditing.csv"
}
