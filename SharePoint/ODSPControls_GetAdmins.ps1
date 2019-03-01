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

Param(
    [CmdletBinding()]
    [Parameter(Mandatory=$false)]
    [String]$OutputDir,
    [String]$SPOAdmin,
    [String]$SPOTenantName,
    [Switch]$SPOPermissionOptIn
)

Add-Type -TypeDefinition @"
   public enum SiteCollectionAdminState
   {
        Needed,
        NotNeeded,
        Skip
   }
"@

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
    if ($needsAdmin -and $SPOPermissionOptIn -eq $false)
    {
        Write-Verbose "$(Get-Date) Grant-SiteCollectionAdmin Skipping $($Site.URL) Needs Admin $needsAdmin PermissionOptIn $SPOPermissionOptIn"
        [SiteCollectionAdminState]$adminState = [SiteCollectionAdminState]::Skip
    }
    # Grant access to the site collection, if required
    elseif ($needsAdmin)
    {
        try{
            Write-Verbose "$(Get-Date) Grant-SiteCollectionAdmin Adding $($SPOAdmin) $($Site.URL) Needs Admin $needsAdmin PermissionOptIn $SPOPermissionOptIn"
            Set-SPOUser -site $Site -LoginName $SPOAdmin -IsSiteCollectionAdmin $True | Out-Null
    
            # Workaround for a race condition that has PnP connect to SPO before the permission access is committed
            Start-Sleep -Seconds 1
    
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
    $siteUrl = "https://$SPOTenantName-admin.sharepoint.com"
    Write-Verbose "$(Get-Date) connect to tenant admin $($siteUrl)"
    Connect-SPOService -Url $siteUrl

    #Get list of admins per site
    $sites = Get-SPOSite -Limit All -IncludePersonalSite $true
    $admins = @()
    $siteAdmins = @()
    
    # Isolate the valid sites - Matches *.sharepoint.com/sites/*, *.sharepoint.com/teams/*, *.sharepoint.com
    $validSites = $sites | `
    Where-Object { $_.Url -match '((\.sharepoint\.com\/(sites|teams))|(^https:\/\/.+(?<!-my)\.sharepoint\.com\/?$))'}

    foreach ($site in $validSites)
    {
        if($site.Template -ne "GROUP#0")
        {
            Write-Verbose "$(Get-Date) connecting to site $($site.Url)"
            Write-Verbose "$(Get-Date) Get-SPOAdminList Processing $($site.Url)"
            # Grant permission to the site collection, if needed AND allowed
            [SiteCollectionAdminState]$adminState = Grant-SiteCollectionAdmin -Site $site
            # Skip this site collection if permission is not granted
            if ($adminState -eq [SiteCollectionAdminState]::Skip)
            {
                continue
            }

            $siteAdmins = Get-SPOUser -Site $site -Limit ALL | Where-Object { $_.IsSiteAdmin -eq $true}
            Write-Host "$(Get-Date) Get-SPOSiteCollectionProperties $site SiteAdmins $($siteAdmins.Count)"
            $count = $siteAdmins.Count
            
            $AdminList = @()

            for($i=0; $i -lt $count; $i++)
            {
                $AdminList += $siteAdmins[$i].DisplayName
            }
            
            $admins += New-Object psobject -Property @{
                Url=$($site.Url)
                Admins=$(@($AdminList) -join ',')
            }
                        
            # Cleanup permission changes, if any
            Revoke-SiteCollectionAdmin -Site $site -AdminState $adminState
        }

    }
    $admins | Export-Csv "$OutputDir\admins_CSV.csv"

}

#region Collect - SPO Site Admins
Write-Host "$(Get-Date) Getting SPO Site Admins"

Get-SPOAdminsList 
#endregion

