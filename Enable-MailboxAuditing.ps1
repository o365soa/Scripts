<#
	.SYNOPSIS
		Enable Mailbox Auditing

	.DESCRIPTION

        This script enables Mailbox Auditing either for all mailboxes or for a subset of mailboxes
        using a CSV file.

        You need to be connected to Exchange Online PowerShell in order to perform these activities.
        Go to http://aka.ms/EXOPSPreview in order to download the Exchange Online PowerShell
        module.

    .PARAMETER  All
        This All switch performs against all mailboxes by first doing a Get-Mailbox
    
    .PARAMETER CSV
        Path to a CSV file containing the identities of mailboxes you want to set the auditing for
        This CSV should have one column labelled Identity

	.LINK
		about_functions_advanced

#>
Param (
    [Parameter(ParameterSetName='AllMailboxes')]
        [Switch]$All,
    [Parameter(ParameterSetName='Selective')]
        [String]$CSV
)

Write-Host "$(Get-Date) Please note, the service will have mailbox auditing enabled by default at some point. Ensure you keep track of the notifications in the message center." -ForegroundColor Green
Sleep 5

# Determine if CSV Used, or ALL
If($All) {

    # Get all users with the Get-Mailbox command
    Write-Host "$(Get-Date) Getting all mailbox users, this may take some time.." -ForegroundColor Green
    $Users = Get-Mailbox -ResultSize:Unlimited | Select 

} Else {

    If($CSV) {

        # Import CSV
        $Users = Import-CSV $CSV

        # Format check CSV
        If(!$Users[0].Identity) {
            Write-Error "CSV file looks malformed. Ensure there is an Identity column for the mailboxes."
            Exit
        }

    } Else {

        # No CSV specified, error out.
        Write-Error "Use either -All switch, or specify a CSV file to include with the -CSV Param"
        Exit

    }

}
	
ForEach ($User in $Users)
{
    Write-Host "$(Get-Date) Setting audit enabled for user $($User.Identity)" -ForegroundColor Green
    Set-Mailbox -Identity $User.Identity -AuditLogAgeLimit 90 -AuditEnabled $true -AuditAdmin UpdateCalendarDelegation,UpdateFolderPermissions,UpdateInboxRules,Update,Move,MoveToDeletedItems,SoftDelete,HardDelete,FolderBind,SendAs,SendOnBehalf,Create,Copy,MailItemsAccessed -AuditDelegate UpdateFolderPermissions,UpdateInboxRules,Update,Move,MoveToDeletedItems,SoftDelete,HardDelete,FolderBind,SendAs,SendOnBehalf,Create -AuditOwner UpdateCalendarDelegation,UpdateFolderPermissions,UpdateInboxRules,Update,MoveToDeletedItems,Move,SoftDelete,HardDelete,Create,MailboxLogin
}
