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
        Group membership will be expanded if not opted out.
        Group expansion requires the Microsoft.Graph.Authentication module and the least-privileged 
            overlapping scope for getting directory roles and group membership: Directory.Read.All
	.Parameter OutputDir
		Directory to save the output file.  Default is the current directory.
	.Parameter SPOAdmin
		UPN of the admin account that has, or will be temporarily granted, site admin permission (unless
        SPOPermissionOptOut is True).
    .PARAMETER SPOAdminDomain
        If not already connected to SharePoint Online and a custom (vanity) domain is used to connect
        (such as MTE customers), this is the FQDN used to connect to the SPO administrative endpoint.
	.Parameter SPOTenantName
		If not already connected to SharePoint Online, this is the name of the tenant, which is the 
        subdomain value of the onmicrosoft domain, e.g., "contoso" for contoso.onmicrosoft.com.
	.Parameter SPOPermissionOptOut
		If the SPOAdmin needs to be added as a site admin to be able to retrieve the list of site admins,
		use this switch if you don't want to add the account as a site admin (and, therefore, skip
		checking any site that would need the account added in order get the list of admins).  If added,
		the permission is removed after the site is checked.
    .Parameter IncludeOneDriveSites
        Switch to include OneDrive for Business sites.
    .Parameter CloudEnvironment
        The cloud instance that hosts the tenant. Used to set the endpoints for authentication and connection.
        Value can be Commercial, USGovGCC, USGovGCCHigh, USGovDoD, China. Default is Commercial.
    .Parameter DoNotExpandGroups
        Do not get the recursive membership of groups assigned site admin.
	.Notes
		Version: 2.5
		Date: January 13, 2025
#>
[CmdletBinding()]
Param(
    [Parameter(Mandatory=$false)][string]$OutputDir = (Get-Location).Path,
    [Parameter(Mandatory=$true)][string]$SPOAdmin,
    [ValidateScript({if (Resolve-DnsName -Name $PSItem) {$true} else {throw "SPO admin domain does not resolve.  Verify you entered a valid fully qualified domain name."}})]
        [ValidateNotNullOrEmpty()][string]$SPOAdminDomain,
    [string]$SPOTenantName,
    [switch]$SPOPermissionOptOut,
    [switch]$IncludeOneDriveSites,
    [ValidateSet("Commercial", "USGovGCC", "USGovGCCHigh", "USGovDoD", "China")][string]$CloudEnvironment="Commercial",
    [switch]$DoNotExpandGroups
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
    param(
        [Parameter(Mandatory=$True)]
        [Microsoft.Online.SharePoint.PowerShell.SPOSite]$Site
    )

    [SiteCollectionAdminState]$adminState = [SiteCollectionAdminState]::NotNeeded

    # Skip this site collection if the current user does not have permissions and
    # permission changes should not be made ($SPOPermissionOptOut)
    if ($SPOPermissionOptOut) {
        Write-Verbose "$(Get-Date) Grant-SiteAdmin: Skipping $($Site.URL) PermissionOptOut $SPOPermissionOptOut"
        [SiteCollectionAdminState]$adminState = [SiteCollectionAdminState]::Skip
    } else {
        # Grant access to the site collection, if required
        try {
            Write-Verbose "$(Get-Date) Grant-SiteAdmin: Adding $SPOAdmin to $($Site.URL)"
            Set-SPOUser -site $Site -LoginName $SPOAdmin -IsSiteCollectionAdmin $True -ErrorAction Stop | Out-Null
            [SiteCollectionAdminState]$adminState = [SiteCollectionAdminState]::Needed
        } catch {
            Write-Verbose "$(Get-Date) Failed to add site admin to site  $($Site.URL)"
        }
    }

    Write-Verbose "$(Get-Date) Grant-SiteAdmin: Finished"

    return $adminState
}

function Revoke-SiteAdmin
{
    param(
        [Parameter(Mandatory=$True)]
        [Microsoft.Online.SharePoint.PowerShell.SPOSite]$Site,
        [Parameter(Mandatory=$True)]
        [SiteCollectionAdminState]$AdminState
    )

    # Cleanup permission changes, if any
    if ($AdminState -eq [SiteCollectionAdminState]::Needed) {
        trap [System.Net.WebException] {
            Write-Verbose "$(Get-Date) $($error[0].Exception.Message)" -Verbose
            Write-Verbose "$(Get-Date) A terminating error was caught by a Trap statement"

            # Add throttling to next attempt. Try to read the value in the Retry-After header, otherwise we just sleep a static amount
            Try { $seconds = 0 ; $seconds = $Error[0].Exception.Response.Headers["Retry-After"] } Catch { $seconds = 5 }
            Write-Host "$(Get-Date) Requests are being throttled. Sleeping for $seconds seconds" -ForegroundColor Yellow
            Start-Sleep -Seconds $seconds
            continue 
        }
        try {
            Write-Verbose "$(Get-Date) Revoke-SiteAdmin: $($site.url) Revoking $SPOAdminUPN"
            Set-SPOUser -site $Site -LoginName $SPOAdmin -IsSiteCollectionAdmin $False | Out-Null
        } catch [System.Net.WebException] {
            Write-Verbose "$(Get-Date) $($error[0].Exception.Message)" -Verbose
            Write-Verbose "$(Get-Date) A non-terminating error was caught by a Catch statement"

            # Add throttling to next attempt. Try to read the value in the Retry-After header, otherwise we just sleep a static amount
            try {
                $seconds = 0 ; $seconds = $Error[0].Exception.Response.Headers["Retry-After"]
            } catch {
                $seconds = 5
            }
            Write-Host "$(Get-Date) Requests are being throttled. Sleeping for $seconds seconds" -ForegroundColor Yellow
            Start-Sleep -Seconds $seconds
        } catch {
            Write-Verbose "$(Get-Date) $($error[0].Exception.Message)" -Verbose
        }
    }
    
    Write-Verbose "$(Get-Date) Revoke-SiteAdmin: Finished"
}

function Get-MgGraphData {
    Param (
        [Switch]$Beta,
        [String]$Endpoint,
        [string]$Query
    )

    if ($Beta) {
        $Uri = "/beta/$Endpoint"
    } else {
        $Uri = "/v1.0/$Endpoint"
    }

    $ApiUrl = $Uri+$Query
    $result = @()
    try {
        Write-Verbose "Running Invoke-MgGraphRequest to $ApiUrl"
        do {
            # Get data via Graph and continue paging until complete
            $response = Invoke-MgGraphRequest -Method GET $apiUrl -OutputType PSObject
            $apiUrl = $($response."@odata.nextLink")
            if ($apiUrl) { Write-Verbose "@odata.nextLink: $apiUrl" }
            if ($response.Value) {
                $result += $response.Value
            } else {
                $result += $response
            }
        }
        until ($null -eq $response."@odata.nextLink")
    } 
    catch {
        Write-Warning "Unable to run Invoke-MgGraphRequest to $ApiUrl"
        Write-Verbose $error[0]
    }

    return $result
}

function Get-AadRoleMembers {
    param (
        $RoleId
    )
    
    $members = @()
    $roleMembers = Get-MgGraphData -Endpoint "directoryRoles" -Query "/$Id/members"

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

    # Perform one-time lookup of Entra roles to get the object ID of each
    if (-not($script:aadRoleIds)) {
        Write-Verbose "$(Get-Date) Getting Microsoft Entra roles"
        $aadRoles = Get-MgGraphData -Beta -Endpoint directoryRoles
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
                # Group is an Entra role
                $groupMembers = Get-AadRoleMembers -RoleId $Id
                [array]$memberIds = $groupMembers
                Write-Verbose -Message "$(Get-Date) $DisplayName contains $($memberIds.Count) transitive members."
            }
        else {
            # Beta endpoint is used because service principals are not included in response in v1.0
            # 100 members are returned by default. Can use $top to get up to 999
            # If M365 Group, drop _o from end of GUID for site owner group and get only the owners
            if ($id.Length -eq 38) {
                $lookupId = $id.Substring(0,36)
                $groupMembers = Get-MgGraphData -Beta -Endpoint groups -Query "/$lookupId/owners?`$top=999"
            }
            else {
                $groupMembers = Get-MgGraphData -Beta -Endpoint groups -Query "/$Id/transitiveMembers?`$top=999"
            }
            $memberIds = @()
            foreach ($member in ($groupMembers | Where-Object {$_."@odata.type" -ne '#microsoft.graph.group'})) {
                switch ($member."@odata.type") {
                    '#microsoft.graph.user' {$memberIds += $member.userPrincipalName}
                    '#microsoft.graph.servicePrincipal' {$memberIds += $member.appDisplayName}
                    default {$memberIds += $member.Id}
                }
            }
            Write-Verbose -Message "$(Get-Date) $DisplayName contains $($memberIds.Count) owners or transitive members."
        }
        # Add the group and membership to hash table
        $script:groupMembers.Add($Id,$memberIds)
        return $memberIds
    }
}

function Get-SPOAdminsList {
    try {
        Get-SPOTenant -ErrorAction Stop | Out-Null
    } catch {
        if ($SPOAdminDomain) {
            Connect-SPOService -Url $SPOAdminDomain | Out-Null
        } else {
            if (-not $SPOTenantName) {
                $SPOTenantName = Read-Host -Prompt "Enter the tenant name (without the onmicrosoft domain suffix)"
            }
            switch ($CloudEnvironment) {
                "Commercial"   {Connect-SPOService -Url "https://$SPOTenantName-admin.sharepoint.com" | Out-Null}
                "USGovGCC"     {Connect-SPOService -Url "https://$SPOTenantName-admin.sharepoint.com" | Out-Null}
                "USGovGCCHigh" {Connect-SPOService -Url "https://$SPOTenantName-admin.sharepoint.us" -Region ITAR | Out-Null}
                "USGovDoD"     {Connect-SPOService -Url "https://$SPOTenantName-admin.dps.mil" -Region ITAR | Out-Null}
                "China"        {Connect-SPOService -Url "https://$SPOTenantName-admin.sharepoint.cn" -Region China | Out-Null}
            }
        }
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
        trap [System.Net.WebException] {
            Write-Verbose "$(Get-Date) $($error[0].Exception.Message)" -Verbose
            Write-Verbose "$(Get-Date) A terminating error was caught by a Trap statement"

            # Add throttling to next attempt. Try to read the value in the Retry-After header, otherwise we just sleep a static amount
            try { $retry = $Error[0].Exception.Response.Headers["Retry-After"] } Catch {}
            if ($retry) { 
                $seconds = [math]::Ceiling($retry) 
            }

            if (-Not $Seconds) { $seconds = 15 }
            Write-Host "$(Get-Date) Requests are being throttled. Sleeping for $seconds seconds" -ForegroundColor Yellow
            Start-Sleep -Seconds $seconds
            
            continue 
        }
        try {
            Write-Verbose "$(Get-Date) Get-SPOSiteProperties: Processing $($site.Url)"
            # Try and get the list of site users
            $siteUsers = Get-SPOUser -Site $site -Limit ALL -ErrorAction Stop
        } catch [Microsoft.SharePoint.Client.ServerUnauthorizedAccessException] { 
            Write-Verbose "$(Get-Date) Access is denied to site $($Site.Url)"
            # Grant permission to the site, if needed and allowed
            [SiteCollectionAdminState]$adminState = Grant-SiteAdmin -Site $site
    
            # Skip this site collection if permission is not granted
            if ($adminState -ne [SiteCollectionAdminState]::Skip) {
                #Try to get site users again
                try {
                    $siteUsers = Get-SPOUser -Site $site -Limit ALL -ErrorAction Stop
                } catch {}
            }
        } catch { 
            Write-Verbose "$(Get-Date) $($error[0].Exception.Message)" -Verbose
            Write-Verbose "$(Get-Date) A non-terminating error was caught by a Catch statement"

            # Add throttling to next attempt. Try to read the value in the Retry-After header, otherwise we just sleep a static amount
            try { $retry = $Error[0].Exception.Response.Headers["Retry-After"] } Catch {}
            if ($retry) { 
                $seconds = [math]::Ceiling($retry) 
            }
            
            if (-not $Seconds) { $seconds = 15 }
            Write-Host "$(Get-Date) Requests are being throttled. Sleeping for $seconds seconds" -ForegroundColor Yellow
            Start-Sleep -Seconds $seconds
        } finally {
            $siteAdminEntries = $siteUsers | Where-Object { $_.IsSiteAdmin -eq $true}
            $siteAdminUsers = @()
            $siteAdminGroups = @()
            if ($siteAdminEntries) {
                if ($DoNotExpandGroups -eq $false) {
                    # Loop through site admins, expanding groups
                    foreach ($entry in $siteAdminEntries) {
                        if ($entry.IsGroup -eq $true) {
                            Write-Verbose -Message "$(Get-Date) Site admin entry `"$($entry.DisplayName)`" is a group. Expanding members."
                            [array]$groupMembers = Get-GroupMembers -Id $entry.LoginName -DisplayName $entry.DisplayName 
                            $siteAdminGroups += $entry.DisplayName + ' (' + $($groupMembers -join ' ') + ')'
                        } else {
                            $siteAdminUsers += $entry.LoginName
                        }
                    }
                    # Do not include admin when the script added the site admin     
                    if ($AdminState -eq [SiteCollectionAdminState]::Needed) {
                        [array]$siteAdminUsers = $siteAdminUsers | Where-Object {$_ -ne $SpoAdmin}
                    }
                } else {
                    for ($i=0; $i -lt $siteAdminEntries.Count; $i++) {
                        # Skip admin if listed only because of "admin needed"
                        if (-not($siteAdminEntries[$i].LoginName -eq $SpoAdmin -and $adminState -eq [SiteCollectionAdminState]::Needed)) {
                            # Include entry type
                            if ($siteAdminEntries[$i].IsGroup) {
                                $siteAdminGroups += $siteAdminEntries[$i].DisplayName
                            } else {
                                $siteAdminUsers += $siteAdminEntries[$i].LoginName
                            }
                        }
                    }
                }
                # Cleanup permission changes
                if ($AdminState -eq [SiteCollectionAdminState]::Needed) {
                   Revoke-SiteAdmin -Site $site -AdminState $adminState
                }
            } elseif ($siteUsers) {
                # Collection of users returned, but none are listed as a site admin
                $siteAdminUsers = "[No site admins]"
            } else {
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

if ($DoNotExpandGroups -eq $false) {
    if (-not((Get-Module -Name Microsoft.Graph.Authentication -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1).Version.Major -ge 2)){
        throw "The Microsoft.Graph.Authentication module v2 is required when group expansion is included. (Use -DoNotExpandGroups to opt out.)"
    }
    if ((Get-Module -Name Microsoft.Graph.Authentication).Version.Major -lt 2) {
        Remove-Module Microsoft.Graph.Authentication
    }
    Import-Module -Name Microsoft.Graph.Authentication -MinimumVersion 2.0.0

    switch ($CloudEnvironment) {
        "Commercial"   {$cloud = 'Global'}
        "USGovGCC"     {$cloud = 'Global'}
        "USGovGCCHigh" {$cloud = 'USGov'}
        "USGovDoD"     {$cloud = 'USGovDoD'}
        "China"        {$cloud = 'China'}
    }
    Write-Host -ForegroundColor Green "$(Get-Date) Connecting to Microsoft Graph with delegated authentication..."
    if ((Get-MgContext).Scopes -notcontains 'Directory.Read.All') {
        try {
            Connect-MgGraph -Scopes 'Directory.Read.All' -Environment $cloud -ContextScope CurrentUser -ErrorAction Stop | Out-Null
        }
        catch {
            throw $_
        }
    }
}
else {
    $tag = "(without expanding groups)"
}

Write-Host "$(Get-Date) Getting SPO site admins $tag" -ForegroundColor Green

Get-SPOAdminsList 
#endregion

