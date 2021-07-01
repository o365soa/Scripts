This folder contains sample scripts to help you get additional information from your SharePoint Online tenant. These scripts are provided for reference and example purposes only. Please ensure that you thoroughly review each script before executing it within your environment.

### Get Admins for Site Collections

Get-SPOSiteAdmins.ps1

This script iterates through all the SharePoint Online site collections with a /sites or /teams path and gets the list of admin entries for each site. If the user running the script does not have site admin permission to the site collection, the script will add the user as a site admin to the site collection and then remove the user after the site has been checked.

# Disclaimer

The sample scripts are not supported under any Microsoft standard support program or service. The sample scripts are provided AS IS without warranty of any kind. Microsoft further disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. In no event shall Microsoft, its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out of the use of or inability to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.

