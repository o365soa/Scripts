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
        Gets/Sets Power Platform auditing settings

    .DESCRIPTION
        Retrieve the environment-level audit settings for all active environments with a Dataverse
        (excluding Teams-only environments). Can optionally enable environment-level auditing
        for these environments.
        - The 'Microsoft Security Assessment' Entra app registration must exist.
        - Requires ExchangeOnlineManagement module for the MSAL library distributed with it.

    .PARAMETER EnableAuditing
        Switch to enable the three environment-level settings (Audit enabled, Log access, Log read)
        for every environment that does not have all three enabled.
    
    .PARAMETER RetentionPeriod
        When used with EnableAuditing, this sets the retention period, in days, for the audit logs.
        Default is 30. Valid value is -1 (forever) or a number between 1 and 365,000.
        Note: The period will not be applied in an environment that already has any of the three settings enabled.

    .PARAMETER AsJson
        Configure the output file type to be JSON instead of CSV

    .PARAMETER CloudEnvironment
        Cloud instance of the tenant. Valid values: Commercial, USGovGCC, USGovGCCHigh, USGovDoD, China
        Default is Commercial.

    .EXAMPLE
        .\Get-DataverseAuditing.ps1
    .EXAMPLE
        .\Get-DataverseAuditing.ps1 -EnableAuditing
        
    .NOTES
        Version 1.1
        January 6, 2025
#>

#Requires -Version 5
#Requires -Modules @{ModuleName="ExchangeOnlineManagement"; ModuleVersion="3.0.0"}
#Requires -Modules Microsoft.PowerApps.Administration.PowerShell, Microsoft.Graph.Authentication

[CmdletBinding()]
Param(
    [switch]$EnableAuditing,
    [switch]$AsJson,
    [string]$CloudEnvironment = 'Commercial',
    [ValidateScript({if ($_ -eq -1 -or ($_ -ge 1 -and $_ -le 365000)){$true}else{throw "$_ is not valid. Specify a number between 1 and 365000 or -1."}})][Int32]$RetentionPeriod = 30
)

# Load MSAL
$ExoModule = Get-Module -Name "ExchangeOnlineManagement" -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
if ($PSEdition -eq 'Core'){
    $Folder = "netCore"
} else {
    $Folder = "NetFramework"
}
$MSAL = Join-Path -Path $ExoModule.ModuleBase -ChildPath "$($Folder)\Microsoft.Identity.Client.dll"
Try {Add-Type -LiteralPath $MSAL | Out-Null} Catch {}

switch ($CloudEnvironment) {
    "Commercial"   {$cloud = "Global";$authEndpoint = 'https://login.microsoftonline.com'}
    "USGovGCC"     {$cloud = "Global";$authEndpoint = 'https://login.microsoftonline.com'}
    "USGovGCCHigh" {$cloud = "USGov";$authEndpoint = 'https://login.microsoftonline.us'}
    "USGovDoD"     {$cloud = "USGovDoD";$authEndpoint = 'https://login.microsoftonline.us'}
    "China"        {$cloud = "China";$authEndpoint = 'https://login.partner.microsoftonline.cn'}            
}
Connect-MgGraph -ContextScope Process -Scopes 'Application.ReadWrite.All','User.Read' -Environment $cloud -NoWelcome
if (-not (Get-MgContext)) {
    Write-Error -Message "Failed to authenticate with Microsoft Graph"
    exit
}

# Get the SOA app registration
try {
    $App = (Invoke-MgGraphRequest -Method GET -Uri "/v1.0/applications?`$filter=web/redirectUris/any(p:p eq 'https://security.optimization.assessment.local')&`$count=true" -Headers @{'ConsistencyLevel' = 'eventual'} -OutputType PSObject).Value}
catch {
    Write-Error -Message "Failed to retrieve the SOA application. If the application does not exist, run 'Install-Module SOA' then 'Install-SOAPrerequisites -EntraAppOnly'."
    exit
}

# Build public client that will be used to connect to the Dataverse environments
$GraphAppDomain = ((Invoke-MgGraphRequest GET "/v1.0/organization" -OutputType PSObject).Value | Select-Object -ExpandProperty VerifiedDomains | Where-Object { $_.isInitial }).Name
$Authority = "$authEndpoint/$GraphAppDomain"
$pubApp = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($App.AppId).WithRedirectUri('https://login.microsoftonline.com/common/oauth2/nativeclient').WithAuthority($Authority).Build()
Write-Host "Select an account with Global or Dynamics 365 Administrator role:"
$pubApp.AcquireTokenInteractive($Scopes).ExecuteAsync().GetAwaiter().GetResult() | Out-Null
$account = $pubApp.GetAccountsAsync().Result.Username

Write-Host "Getting environments..."
[array]$Environments = Get-AdminPowerAppEnvironment
$Result = @()
$envCount = 0
if ($EnableAuditing) {
    $verbs = "Getting and setting"
} else {
    $verbs = "Getting"
}
Write-Host "Processing environments..."
foreach ($instance in $Environments) {
    $envCount++
    Write-Progress -Activity "$verbs Dataverse audit settings" -CurrentOperation "Environment $envCount of $($Environments.Count): $($instance.DisplayName)" -PercentComplete (($envCount / $Environments.Count)  * 100)
    $apiUrl = $null
    $apiUrl = $instance.Internal.properties.linkedEnvironmentMetadata.instanceApiUrl
    Write-Verbose "Environment: $($instance.DisplayName)"

    # Connect to DV if environment is active (Ready) and not a Teams-only environment
    if ($apiUrl -and $instance.Internal.properties.linkedEnvironmentMetadata.instanceState -eq 'Ready' -and $instance.Internal.properties.linkedEnvironmentMetadata.platformSku -ne 'Lite') {
        $scope = New-Object System.Collections.Generic.List[string]
        $scope.Add("$apiUrl/.default")

        $token = $pubApp.AcquireTokenSilent($scope, $account).ExecuteAsync().GetAwaiter().GetResult()

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
            $response = $null
            try {
                $response = Invoke-RestMethod -Uri "$apiUrl/api/data/$verStr/organizations?`$select=organizationid,isauditenabled,auditretentionperiodv2,isuseraccessauditenabled,isreadauditenabled" -Headers $headers
            } catch {
                Write-Warning -Message "Failed to retrieve data from Dataverse associated with $($instance.DisplayName)."
                Write-Error $_
            }
            if ($response) {
                $OrgID = $response.value.organizationid

                $result += [pscustomobject] [ordered] @{
                    EnvDisplayName = $instance.DisplayName
                    OrgID = $OrgID
                    EnvState = $instance.Internal.properties.linkedEnvironmentMetadata.instanceState
                    IsAuditEnabled = $response.value.isauditenabled
                    IsAccessAuditEnabled = $response.value.isuseraccessauditenabled
                    IsReadLogsEnabled = $response.value.isreadauditenabled
                    RetentionPeriod = $response.value.auditretentionperiodv2
                }
                # Adding these properties even if auditing won't be updated for the environment is so the first object contains all
                # possible properties so export to CSV will include all columns.
                if ($EnableAuditing) {
                    $result[$result.Count - 1] | Add-Member -MemberType NoteProperty -Name 'AuditSettingsUpdated' -Value $null
                    $result[$result.Count - 1] | Add-Member -MemberType NoteProperty -Name 'RetentionPeriodUpdated' -Value $null
                }

                if ($EnableAuditing -and ($response.value.isauditenabled -ne $true -or $response.value.isuseraccessauditenabled -ne $true -or $response.value.isreadauditenabled -ne $true)) {
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
                    if ($response.value.isauditenabled -ne $true -and $response.value.isuseraccessauditenabled -ne $true -and $response.value.isreadauditenabled -ne $true) {
                        $Body | Add-Member -MemberType NoteProperty -Name 'auditretentionperiodv2' -Value $RetentionPeriod
                        $updatePeriod = $true
                    } else {
                        $updatePeriod = $false
                    }
                    Write-Verbose -Message "Setting Auditing for $apiUrl"
                    Write-Verbose -Message "OrgID: $OrgID"

                    try {
                        Invoke-RestMethod -Method Patch -Uri "$apiUrl/api/data/$verStr/organizations($OrgID)" -Headers $Headers -Body ($body | ConvertTo-Json)
                        $result[$result.Count - 1].AuditSettingsUpdated =  $true
                        $result[$result.Count - 1].RetentionPeriodUpdated = $(if($updatePeriod){$true}else{$false})
                    } catch {
                        Write-Warning "Error while making PATCH request to $apiUrl"
                        $result[$result.Count - 1].AuditSettingsUpdated = $false
                    }
                }
            }
        }
    }
}
Write-Progress -Activity "$verbs Dataverse audit settings" -Completed

If ($AsJson) {
    $Result | ConvertTo-Json -Depth 10 | Out-File -FilePath "Dataverse-Auditing.json"
    Write-Host "Results saved to Dataverse-Auditing.json"
} Else {
    $Result | Export-Csv -NoTypeInformation -Path "Dataverse-Auditing.csv"
    Write-Host "Results saved to Dataverse-Auditing.csv"
}
