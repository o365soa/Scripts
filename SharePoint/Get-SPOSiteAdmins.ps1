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

#Requires -Modules Microsoft.Online.SharePoint.PowerShell

<#
	.Synopsis
		Get the list of site admins for every site.
	.Description
		Saves to CSV the site admins for every site, differentiating between users and groups.
        Group membership will be expanded if that option is used.
	.Parameter OutputDir
		Directory to save the output file.  Default is the current directory.
	.Parameter SPOAdmin
		UPN of the admin account that has, or will be granted, site admin permission (unless
        SPOPermissionOptOut is True).
	.Parameter SPOTenantName
		Name of the tenant, which is the subdomain value of the .onmicrosoft.com domain, e.g.,
		"contoso" for contoso.onmicrosoft.com.  If not already connected to SPO and the value is not
		entered via the parameter, you will be asked to enter it in order to continue.
	.Parameter SPOPermissionOptOut
		If the SPOAdmin needs to be added as a site admin to be able to retrieve the list of site admins,
		use this switch if you don't want to add the account as a site admin (and, therefore, skip
		checking any site that would need the account added in order get the list of admins).  If added,
		the permission is removed after the site is checked.
    .Parameter IncludeOneDriveSites
        Include OneDrive for Business sites.
    .Parameter O365EnvironmentName
        When using ExpandGroups, changes endpoints to use for Microsoft Graph for tenants in sovereign clouds for group expansion. 
        The accepted values are Commercial [Default],USGovGCC, USGovGCCHigh, USGovDoD, Germany, China.
    .Parameter ExpandGroups
        Get the membership of a group (including recursion) assigned site admin. The Azure AD Preview module and the SOA Azure AD
        application are required.
	.Notes
		Version: 2.2
		Date: April 6, 2023
#>

Param(
    [CmdletBinding()]
    [Parameter(Mandatory=$false)][string]$OutputDir = (Get-Location).Path,
    [Parameter(Mandatory=$true)][string]$SPOAdmin,
    [string]$SPOTenantName,
    [switch]$SPOPermissionOptOut,
    [switch]$IncludeOneDriveSites,
    [ValidateSet("Commercial", "USGovGCC", "USGovGCCHigh", "USGovDoD", "Germany", "China")][string]$O365EnvironmentName="Commercial",
    [switch]$ExpandGroups
)

if (-not([System.Management.Automation.PSTypeName]'SiteCollectionAdminState').Type){
	Add-Type -TypeDefinition @"
	   public enum SiteCollectionAdminState
	   {
	        Needed,
	        NotNeeded,
	        Skip
	   }
"@
}

function Grant-SiteAdmin
{
    Param(
        [Parameter(Mandatory=$True)]
        [Microsoft.Online.SharePoint.PowerShell.SPOSite]$Site
    )

    [SiteCollectionAdminState]$adminState = [SiteCollectionAdminState]::NotNeeded

    # Skip this site collection if the current user does not have permissions and
    # permission changes should not be made ($SPOPermissionOptOut)
    #if ($needsAdmin -and $SPOPermissionOptOut)
    if ($SPOPermissionOptOut) {
        Write-Verbose "$(Get-Date) Grant-SiteAdmin: Skipping $($Site.URL) PermissionOptOut $SPOPermissionOptOut"
        [SiteCollectionAdminState]$adminState = [SiteCollectionAdminState]::Skip
    }
    # Grant access to the site collection, if required
    else {
        try{
            Write-Verbose "$(Get-Date) Grant-SiteAdmin: Adding $SPOAdmin to $($Site.URL)"
            Set-SPOUser -site $Site -LoginName $SPOAdmin -IsSiteCollectionAdmin $True -ErrorAction Stop | Out-Null
            [SiteCollectionAdminState]$adminState = [SiteCollectionAdminState]::Needed
        }
        catch{
            Write-Verbose "$(Get-Date) Failed to add site admin to site  $($Site.URL)"
        }
    }

    Write-Verbose "$(Get-Date) Grant-SiteAdmin: Finished"

    return $adminState
}

function Revoke-SiteAdmin
{
    Param(
        [Parameter(Mandatory=$True)]
        [Microsoft.Online.SharePoint.PowerShell.SPOSite]$Site,
        [Parameter(Mandatory=$True)]
        [SiteCollectionAdminState]$AdminState
    )

    # Cleanup permission changes, if any
    if ($AdminState -eq [SiteCollectionAdminState]::Needed)
    {
        Trap [System.Net.WebException] {
            Write-Verbose "$(Get-Date) $($error[0].Exception.Message)" -Verbose
            Write-Verbose "$(Get-Date) A terminating error was caught by a Trap statement"

            # Add throttling to next attempt. Try to read the value in the Retry-After header, otherwise we just sleep a static amount
            Try { $seconds = 0 ; $seconds = $Error[0].Exception.Response.Headers["Retry-After"] } Catch { $seconds = 5 }
            Write-Host "$(Get-Date) Requests are being throttled. Sleeping for $seconds seconds" -ForegroundColor Yellow
            Start-Sleep -Seconds $seconds
            continue 
        }
        Try {
            Write-Verbose "$(Get-Date) Revoke-SiteAdmin: $($site.url) Revoking $SPOAdminUPN"
            Set-SPOUser -site $Site -LoginName $SPOAdmin -IsSiteCollectionAdmin $False | Out-Null
        } Catch [System.Net.WebException] {
            Write-Verbose "$(Get-Date) $($error[0].Exception.Message)" -Verbose
            Write-Verbose "$(Get-Date) A non-terminating error was caught by a Catch statement"

            # Add throttling to next attempt. Try to read the value in the Retry-After header, otherwise we just sleep a static amount
            Try { $seconds = 0 ; $seconds = $Error[0].Exception.Response.Headers["Retry-After"] } Catch { $seconds = 5 }
            Write-Host "$(Get-Date) Requests are being throttled. Sleeping for $seconds seconds" -ForegroundColor Yellow
            Start-Sleep -Seconds $seconds
        } Catch {
            Write-Verbose "$(Get-Date) $($error[0].Exception.Message)" -Verbose
        }

    }
    
    Write-Verbose "$(Get-Date) Revoke-SiteAdmin: Finished"
}

function Get-GraphToken {
    Param(
        $ClientID,
        $ClientSecret,
        $TenantDomain,
        $Resource="https://graph.microsoft.com"
    )

    switch ($O365EnvironmentName) {
        "Commercial"   {$Resource = "https://graph.microsoft.com";break}
        "USGovGCC"     {$Resource = "https://graph.microsoft.com";break}
        "USGovGCCHigh" {$Resource = "https://graph.microsoft.us";break}
        "USGovDoD"     {$Resource = "https://dod-graph.microsoft.us";break}
        "Germany"      {$Resource = "https://graph.microsoft.com";break}
        "China"        {$Resource = "https://microsoftgraph.chinacloudapi.cn"}
    }

    # Get an Oauth 2 access token based on client id, secret and tenant domain
    # Get-MSALAccessToken is in SOA module
    $Result = Get-MSALAccessToken -TenantName $TenantDomain -ClientID $ClientID -Secret $ClientSecret -Resource $Resource -O365EnvironmentName $O365EnvironmentName

    if ($null -ne $Result.AccessToken) {
        Return $Result
    } Else {
        Write-Error "Failed to get token."
        Return $False
    }
}
function Get-MgGraphData {
    <#

        Preferred to use this function for querying data from the Graph API as it leverages the native Invoke-MgGraphRequest cmdlet
    
    #>
    Param (
        [Switch]$Beta,
        [String]$Endpoint,
        $ClientID,
        $ClientSecret,
        $TenantDomain,
        [string]$Query
    )

    switch ($O365EnvironmentName) {
        "Commercial"   {$Base = "https://graph.microsoft.com";break}
        "USGovGCC"     {$Base = "https://graph.microsoft.com";break}
        "USGovGCCHigh" {$Base = "https://graph.microsoft.us";break}
        "USGovDoD"     {$Base = "https://dod-graph.microsoft.us";break}
        "Germany"      {$Base = "https://graph.microsoft.com";break}
        "China"        {$Base = "https://microsoftgraph.chinacloudapi.cn";break}
    }

    # Request endpoints found at http://aka.ms/GraphApiRef
    If($Beta) {
        $Uri = "$Base/beta/$Endpoint"
    } Else {
        $Uri = "$Base/v1.0/$Endpoint"
    }

    $ApiUrl = $Uri+$Query

    # Get a new access token with every request to avoid expiration, until v2 of the Graph SDK modules release which support client secret credentials using Connect-MgGraph
    $Token = Get-GraphToken -ClientID $ClientID -ClientSecret $ClientSecret -TenantDomain $TenantDomain


    # Skip call if access token was not obtained
    if ($null -ne $Token) {
        Try {
            $MgToken = $Token.AccessToken | ConvertTo-SecureString -AsPlainText -Force
            Write-Verbose "Running Invoke-MgGraphRequest to $ApiUrl"
            $result = Invoke-MgGraphRequest -Method GET $ApiUrl -Authentication UserProvidedToken -Token $MgToken
        } Catch {
            Write-Warning "Unable to run Invoke-MgGraphRequest to $ApiUrl"
            Write-Verbose $error[0]
        }
    }
    else {
        Write-Warning -Message "$(Get-Date) Skipping this Graph API call because an access token was not obtained."
    }

    return ($result | ConvertTo-Json -Depth 10)
}
function Get-AadRoleMembers {
    param (
        $RoleId
    )
    
    $members = @()
    $roleMembers = (Get-MgGraphData -Endpoint "directoryRoles" -Query "/$Id/members" -ClientID $AzureADApp.AppId -ClientSecret $AzureADAppCred -TenantDomain $AppDomain | ConvertFrom-Json).Value

    foreach ($rm in $roleMembers) {
        # Member can be a user/SP or group
        switch ($rm."@odata.type") {
            '#microsoft.graph.user' {$members += $rm.userPrincipalName}
            '#microsoft.graph.servicePrincipal' {$members += $rm.appDisplayName}
            '#microsoft.graph.group' {
                Write-Verbose "$(Get-Date) Group: $($rm.displayName) in role $id"
                $groupMembers = Get-GroupMembers -Id $rm.Id -DisplayName $rm.displayName
                if ($groupMembers) {
                    $members += $groupMembers
                }
            }
            default {$members += $rm.Id}
        }
    }
    return $members
}
function Get-GroupMembers {
    param (
        $Id,
        $DisplayName
    )

    # Perform one-time lookup of AAD roles to get the object ID of each
    if (-not($script:aadRoleIds)) {
        Write-Verbose "$(Get-Date) Getting Azure AD roles"
        $aadRoles = (Get-MgGraphData -Beta -Endpoint directoryRoles -ClientID $AzureADApp.AppId -ClientSecret $AzureADAppCred -TenantDomain $AppDomain | ConvertFrom-Json).value
        if ($aadRoles) {
            $script:aadRoleIds = $aadRoles.Id
            Write-Verbose "$(Get-Date) $($script:aadRoleIds.Count) role IDs added to collection of roles."
        }
        else {
            Write-Verbose "$(Get-Date) No role IDs added to collection of roles. Was there an error getting the roles?"
        }
    }
    
    # If the group has already been looked up, return the membership from the hash table
    if ($script:groupMembers.ContainsKey($Id)) {
        Write-Verbose -Message "$(Get-Date) Group with ID $Id has been retrieved previously. Returning membership from hash table."
        return $script:groupMembers[$Id]
    }
    else {
        if ($script:aadRoleIds -contains $Id) {
                # Group is an AAD role
                $groupMembers = Get-AadRoleMembers -RoleId $Id
                [array]$memberIds = $groupMembers
                Write-Verbose -Message "$(Get-Date) $DisplayName contains $($memberIds.Count) transitive members."
            }
        else {
            # Beta endpoint is used because service principals are not included in response in v1.0
            # 100 members are returned by default. Can use $top to get up to 999
            $groupMembers = (Get-MgGraphData -Beta -Endpoint groups -Query "/$Id/transitiveMembers?`$top=999" -ClientID $AzureADApp.AppId -ClientSecret $AzureADAppCred -TenantDomain $AppDomain | ConvertFrom-Json).value
            $memberIds = @()
            foreach ($member in ($groupMembers | Where-Object {$_."@odata.type" -ne '#microsoft.graph.group'})) {
                switch ($member."@odata.type") {
                    '#microsoft.graph.user' {$memberIds += $member.userPrincipalName}
                    '#microsoft.graph.servicePrincipal' {$memberIds += $member.appDisplayName}
                    default {$memberIds += $member.Id}
                }
            }
            Write-Verbose -Message "$(Get-Date) $DisplayName contains $($memberIds.Count) transitive members."
        }
        # Add the group and membership to hash table
        $script:groupMembers.Add($Id,$memberIds)
        return $memberIds
    }
}

function Get-SPOAdminsList
{
    if (-not(Get-SpoTenant -ErrorAction SilentlyContinue)) {
    	if (-not($SPOTenantName)){
			$SPOTenantName = Read-Host -Prompt "Please enter the tenant name (without the .onmicrosoft.com suffix)"
		}
		$siteUrl = "https://$SPOTenantName-admin.sharepoint.com"
		Write-Verbose "$(Get-Date) Connect to SPO tenant admin URL: $($siteUrl)"
		Connect-SPOService -Url $siteUrl
	}

    #Get list of sites
    if ($IncludeOneDriveSites -eq $true) {
        [array]$sites = Get-SPOSite -Limit All -IncludePersonalSite $true
    }
    else {
        [array]$sites = Get-SPOSite -Limit All
    }
    $admins = @()
    
    # Get domains for SPO/ODfB sites to support custom domain in sites' FQDN
    $nonOdSite = $sites | Sort-Object -Property Url -Descending | Where-Object {$_.Url.Substring($_.Url.IndexOf('.')-3,3) -ne '-my'} | Select-Object -First 1
    $sPOSitesDomain = $nonOdSite.Url.Substring(8,$nonOdSite.Url.IndexOf('/',8)-8)
    $domainParts = $sPOSitesDomain.Split('.')
    $domainParts[0] = $domainParts[0] + '-my'
    $oDSitesDomain = $domainParts -join '.'

    # Exclude root site because it can take a very long time to get the list of users for the site
    $validSites = $sites | Where-Object {$_.url -match [regex]::Escape($SpoSitesDomain)+"\/(sites|teams)"}
    if ($IncludeOneDriveSites -eq $true) {
        $validSites += $sites | Where-Object {$_.url -match [regex]::Escape($OdSitesDomain)+"\/personal"}
    }
	
    # Create hash tables for group membership
    $script:groupMembers = @{}

	$j = 1
    foreach ($site in $validSites) {
        $siteUsers = $null
        Write-Progress -Activity "SharePoint Collection" -Status "Site Admins ($j of $($validSites.Count))" -CurrentOperation "$($site.Url)" -PercentComplete ($j/$validSites.Count * 100)        
        Write-Verbose "$(Get-Date) Getting site $($site.Url)"
            
        [SiteCollectionAdminState]$adminState = [SiteCollectionAdminState]::NotNeeded
        # Check if SiteAdmins count can be determined, if not, suppress errors. Trap statement also needed to sometimes catch Terminating errors
        Trap [System.Net.WebException] {
            Write-Verbose "$(Get-Date) $($error[0].Exception.Message)" -Verbose
            Write-Verbose "$(Get-Date) A terminating error was caught by a Trap statement"

            # Add throttling to next attempt. Try to read the value in the Retry-After header, otherwise we just sleep a static amount
            Try { $retry = $Error[0].Exception.Response.Headers["Retry-After"] } Catch {}
            if ($retry) { 
                $seconds = [math]::Ceiling($retry) 
            }

            If (-Not $Seconds) { $seconds = 15 }
            Write-Host "$(Get-Date) Requests are being throttled. Sleeping for $seconds seconds" -ForegroundColor Yellow
            Start-Sleep -Seconds $seconds
            
            continue 
        }
        try {
            Write-Verbose "$(Get-Date) Get-SPOSiteProperties: Processing $($site.Url)"
            # Try and get the list of site users
            $siteUsers = Get-SPOUser -Site $site -Limit ALL -ErrorAction Stop
        }
        catch [Microsoft.SharePoint.Client.ServerUnauthorizedAccessException] { 
            Write-Verbose "$(Get-Date) Access is denied to site $($Site.Url)"
            # Grant permission to the site, if needed and allowed
            [SiteCollectionAdminState]$adminState = Grant-SiteAdmin -Site $site
    
            # Skip this site collection if permission is not granted
            if ($adminState -ne [SiteCollectionAdminState]::Skip) {
                #Try to get site users again
                try {
                    $siteUsers = Get-SPOUser -Site $site -Limit ALL -ErrorAction Stop
                }
                catch {}
            }
        }
        catch { 
            Write-Verbose "$(Get-Date) $($error[0].Exception.Message)" -Verbose
            Write-Verbose "$(Get-Date) A non-terminating error was caught by a Catch statement"

            # Add throttling to next attempt. Try to read the value in the Retry-After header, otherwise we just sleep a static amount
            Try { $retry = $Error[0].Exception.Response.Headers["Retry-After"] } Catch {}
            if ($retry) { 
                $seconds = [math]::Ceiling($retry) 
            }
            
            If (-Not $Seconds) { $seconds = 15 }
            Write-Host "$(Get-Date) Requests are being throttled. Sleeping for $seconds seconds" -ForegroundColor Yellow
            Start-Sleep -Seconds $seconds
        }
        finally {
            $siteAdminEntries = $siteUsers | Where-Object { $_.IsSiteAdmin -eq $true}
            $siteAdminUsers = @()
            $siteAdminGroups = @()
            if ($siteAdminEntries) {
                if ($ExpandGroups) {
                    # Loop through site admins, expanding groups
                    foreach ($entry in $siteAdminEntries) {
                        if ($entry.IsGroup -eq $true) {
                            Write-Verbose -Message "$(Get-Date) Site admin entry `"$($entry.DisplayName)`" is a group. Expanding members."
                            # Drop _o from end of GUID for site owner group
                            if ($entry.LoginName.Length -eq 38) {
                                $id = $entry.LoginName.Substring(0,36)
                            }
                            else {
                                $id = $entry.LoginName
                            }
                            [array]$groupMembers = Get-GroupMembers -Id $id -DisplayName $entry.DisplayName 
                            $siteAdminGroups += $entry.DisplayName + ' (' + $($groupMembers -join ' ') + ')'
                        }
                        else {
                            $siteAdminUsers += $entry.LoginName
                        }
                    }
                    # Do not include admin when the script added the site admin     
                    if ($AdminState -eq [SiteCollectionAdminState]::Needed) {
                        [array]$siteAdminUsers = $siteAdminUsers | Where-Object {$_ -ne $SpoAdmin}
                    }
                }
                else {
                    for ($i=0; $i -lt $siteAdminEntries.Count; $i++) {
                        # Skip admin if listed only because of "admin needed"
                        if (-not($siteAdminEntries[$i].LoginName -eq $SpoAdmin -and $adminState -eq [SiteCollectionAdminState]::Needed)) {
                            # Include entry type
                            if ($siteAdminEntries[$i].IsGroup) {
                                $siteAdminGroups += $siteAdminEntries[$i].DisplayName
                            }
                            else {
                                $siteAdminUsers += $siteAdminEntries[$i].LoginName
                            }
                        }
                    }
                }
                # Cleanup permission changes
                if ($AdminState -eq [SiteCollectionAdminState]::Needed) {
                   Revoke-SiteAdmin -Site $site -AdminState $adminState
                }
            }
            elseif ($siteUsers) {
                # Collection of users returned, but none are listed as a site admin
                $siteAdminUsers = "[No site admins]"
            }
            else {
                # No collection returned due to no permission to site
                $siteAdminUsers = "[Access denied to site]"
            }
        }
            
        $admins += New-Object PSObject -Property @{
            Url = $site.Url
            AdminUsers = $siteAdminUsers -join ','
            AdminGroups = $siteAdminGroups -join ','
        }

        $j++
    }
    $admins | Select-Object -Property Url,AdminUsers,AdminGroups | Export-Csv "$OutputDir\SPOSiteAdmins.csv" -NoTypeInformation
    Write-Host "$(Get-Date) Output saved to $OutputDir\SPOSiteAdmins.csv"

}

if ($ExpandGroups) {
    if (-not(Get-Module -Name SOA -ListAvailable)){
        throw "The SOA module is required when ExpandGroups is True. Run `"Install-Module SOA`" first."
    }
    
    # Connect to Azure AD. This is required to create the client secret used by the SOA application.    
    if (Get-Module -Name AzureADPreview -ListAvailable) {
        Import-Module AzureADPreview
        Write-Host -ForegroundColor Green "$(Get-Date) Connecting to Azure AD. Use an administrator account with owner permission to the SOA application."
        Connect-AzureAD -AccountId $SPOAdmin
    }
    else {
        throw "The AzureADPreview module is required when ExpandGroups is True."
    }

    $AppDomain = (Get-AzureADTenantDetail | Select-Object -ExpandProperty VerifiedDomains | Where-Object { $_.Initial }).Name

    $AzureADApp = Get-AzureADApplication -Filter "displayName eq 'Office 365 Security Optimization Assessment'"
    if ($AzureADApp) {
        Write-Host -ForegroundColor Green "$(Get-Date) Creating a new client secret for the SOA application."
        try {
            # Reset-SOAAppSecret is in SOA module
            $AzureADAppCred = Reset-SOAAppSecret -App $AzureADApp -Task "Get Site Admins"
        }
        catch {
            throw "Unable to create client secret. Verify the signed in user has permission to manage the application."
        }
        Write-Host -ForegroundColor Green "$(Get-Date) Sleeping for 60 seconds for replication of the client secret."
        Start-sleep 60
    
    }
    else {
        throw "The SOA Azure AD application does not exist and is required when ExpandGroups is True. Run `"Install-SOAPrerequisites -AzureADAppOnly`"."
    }
}
else {
    $tag = "(without expanding groups)"
}

#region Collect - SPO Site Admins
Write-Host "$(Get-Date) Getting SPO site admins $tag" -ForegroundColor Green

Get-SPOAdminsList 
#endregion

if ($AzureADAppCred) {
    # Remove-SOAAppSecret is in SOA module
    Remove-SOAAppSecret -app $AzureADApp
}
