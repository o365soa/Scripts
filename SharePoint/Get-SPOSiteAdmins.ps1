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
		Help synposis
	.Description
		Help description
	.Parameter OutputDir
		Directory to save the output file.  Default is the current directory.
	.Parameter SPOAdmin
		UPN of the admin account that has, or will be granted, site admin permission (the latter
		only if SPOPermissionOptOut is not True).
	.Parameter SPOTenantName
		Name of the tenant, which is the subdomain value of the .onmicrosoft.com domain, e.g.,
		"contoso" for contoso.onmicrosoft.com.  If not already connected to SPO and the value is not
		entered via the parameter, you will be asked to enter it in order to continue.
	.Parameter SPOPermissionOptOut
		If the SPOAdmin is needed to be added as a site admin to retrieve the list of site admins,
		use this switch if you don't want to add the account as a site admin (and, therefore, skip
		checking any site that would need the account added in order to perform the check).  If added,
		the permission is removed after the site is checked.
	.Notes
		Version: 2.0
		Date: June 22, 2021
#>

Param(
    [CmdletBinding()]
    [Parameter(Mandatory=$false)]
    [String]$OutputDir = (Get-Location).Path,
    [Parameter(Mandatory=$true)][String]$SPOAdmin,
    [String]$SPOTenantName,
    [Switch]$SPOPermissionOptOut
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

function Grant-SiteCollectionAdmin
{
    Param(
        [Parameter(Mandatory=$True)]
        [Microsoft.Online.SharePoint.PowerShell.SPOSite]$Site
    )

    [SiteCollectionAdminState]$adminState = [SiteCollectionAdminState]::NotNeeded

    # Determine if admin rights need to be granted
    try {
        $adminUser = Get-SPOUser -site $Site -LoginName $SPOAdmin -ErrorAction:SilentlyContinue
        $needsAdmin = ($false -eq $adminUser.IsSiteAdmin)
    }
    catch {
        $needsAdmin = $true
    }

    # Skip this site collection if the current user does not have permissions and
    # permission changes should not be made ($SPOPermissionOptOut)
    if ($needsAdmin -and $SPOPermissionOptOut -eq $true)
    {
        Write-Verbose "$(Get-Date) Grant-SiteCollectionAdmin Skipping $($Site.URL) Needs Admin $needsAdmin PermissionOptOut $SPOPermissionOptOut"
        [SiteCollectionAdminState]$adminState = [SiteCollectionAdminState]::Skip
    }
    # Grant access to the site collection, if required
    elseif ($needsAdmin)
    {
        try{
            Write-Verbose "$(Get-Date) Grant-SiteCollectionAdmin Adding $($SPOAdmin) $($Site.URL) Needs Admin $needsAdmin PermissionOptOut $SPOPermissionOptOut"
            Set-SPOUser -site $Site -LoginName $SPOAdmin -IsSiteCollectionAdmin $True | Out-Null
 
            [SiteCollectionAdminState]$adminState = [SiteCollectionAdminState]::Needed
        }
        catch{
            Write-Verbose "Cannot assign permissions to Site Collection $($Site.URL)"
        }
    }

    Write-Verbose "$(Get-Date) Grant-SiteCollectionAdmin Finished"

    return $adminState
}

function Revoke-SiteCollectionAdmin
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
        Write-Verbose "$(Get-Date) Revoke-SiteCollectionAdmin $($site.url) Revoking $SPOAdmin"
        Set-SPOUser -site $Site -LoginName $SPOAdmin -IsSiteCollectionAdmin $False | Out-Null
    }
    
    Write-Verbose "$(Get-Date) Revoke-SiteCollectionAdmin Finished"
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
    [array]$sites = Get-SPOSite -Limit All -IncludePersonalSite $true
    $admins = @()
    $siteAdmins = @()
    
    # Exclude root site because it can take a very long time to get the list of users for the site
    # Get domain for sites to support custom domain in sites' FQDN
    $sitesDomain = $sites[0].Url.Substring(8,$sites[0].Url.IndexOf('/',8)-8)
    $validSites = $sites | Where-Object {$_.url -match [regex]::Escape($sitesDomain)+"\/(sites|teams)"}
	
	$j = 1
    foreach ($site in $validSites)
    {
        Write-Progress -Activity "SharePoint Collection" -Status "Site Collection Properties ($j of $($validSites.Count))" -CurrentOperation "$($site.Url)" -PercentComplete ($j/$validSites.Count * 100)        
		#if($site.Template -ne "GROUP#0")
		#{
            Write-Verbose "$(Get-Date) connecting to site $($site.Url)"
            Write-Verbose "$(Get-Date) Get-SPOAdminList Processing $($site.Url)"
            # Grant permission to the site collection, if needed AND allowed
            [SiteCollectionAdminState]$adminState = Grant-SiteCollectionAdmin -Site $site
            # Skip this site collection if permission is not granted
            if ($adminState -eq [SiteCollectionAdminState]::Skip)
            {
                continue
            }

            $siteAdmins = Get-SPOUser -Site $site -Limit All | Where-Object { $_.IsSiteAdmin -eq $true}
            $count = $siteAdmins.Count
            
            $AdminList = @()

            for($i=0; $i -lt $count; $i++)
            {
				# Skip admin if listed only because of "admin needed"
				if (-not($siteAdmins[$i].LoginName -eq $SpoAdmin -and $adminState -eq [SiteCollectionAdminState]::Needed)) {
					# Include entry type
					if ($siteAdmins[$i].IsGroup) {
						$type = "Group"
					}
					else {
						$type = "User"
					}
	                $AdminList += $siteAdmins[$i].DisplayName+" ($type)"
				}
            }
            
            $admins += New-Object PSObject -Property @{
                Url=$($site.Url)
                Admins=$(@($AdminList) -join ',')
            }
			
			Write-Host "$(Get-Date) Site: $($site.Url) SiteAdminCount: $($AdminList.Count)"
                        
            # Cleanup permission changes, if any
            Revoke-SiteCollectionAdmin -Site $site -AdminState $adminState
        #}
	$j++
    }
    $admins | Export-Csv "$OutputDir\SPOSiteAdmins.csv" -NoTypeInformation

}

#region Collect - SPO Site Admins
Write-Host "$(Get-Date) Getting SPO Site Admins"

Get-SPOAdminsList 
#endregion

