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
	.Parameter UserPrincipalName
		UPN of the user whose configuration to report
	.Parameter StartDate
		Beginning of search window in audit logs. Defaults to seven days (and current time) ago.
	.Parameter EndDate
		End of search window in audit logs.  Defaults to current date/time.
	.Example
		.\Get-PersistenceConfiguration.ps1 johndoe@contoso.com
	.Example
		.\Get-PersistenceConfiguration.ps1 johndoe@contoso.com -StartDate (Get-Date).AddDays(-14)
	.Notes
		Version: 2.0
		Date: November 4, 2025
#>
[CmdletBinding()]
param (
	[parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)]
		[ValidateScript({ForEach-Object{if ((Get-Command -Name Get-Mailbox) -and (Get-Mailbox -Identity $_)) {$true}
		else {throw "Either you are not connected to Exchange Online or $_ is not a valid mailbox identity."}}
		})]
		[Alias("Identity")][string]$UserPrincipalName,
	[parameter(Mandatory=$false)][datetime]$StartDate = (Get-Date).AddDays(-7),
	[parameter(Mandatory=$false)][datetime]$EndDate = (Get-Date).AddDays(1)
	)

#requires -Modules ExchangeOnlineManagement

function Write-ProgressHelper ($activity) {
	Write-Progress -Activity "Checking for persistence configured by account $UserPrincipalName" -Status 'Overall Progress' `
		-CurrentOperation $activity -PercentComplete ((($script:step++)/$totalSteps)*100) -Id 1 
}

function Get-UALData {
	Write-ProgressHelper -Activity 'Searching Unified Audit Log for relevant activities'
	# The certificates and secrets operation uses Unicode character 2013 (not 002d) for the dash and has a trailing space
	$operations = @('Update application – Certificates and secrets management ','New-InboxRule','Set-InboxRule',
		'UpdateInboxRules','Set-Mailbox','AddFolderPermissions','ModifyFolderPermissions','RemoveFolderPermissions',
		'Set-MailboxCalendarFolder','CompanyLinkCreated', 'SecureLinkCreated', 'AnonymousLinkCreated',
		'AnonymousLinkUpdated','AddedToSecureLink','AddedToSharingLink','MemberAdded','AddMemberToGroup.',
		'Add owner to group.','EditFlow','CreateFlow','Set-MailboxJunkEmailConfiguration','Consent to application.',
		'PowerAppPermissionEdited','PublishPowerApp','New-JournalRule','Set-JournalRule','New-TransportRule',
		'Set-TransportRule','Add-MailboxPermission')
	$results = @()
	$sId = "Persistence" + (-join ((65..90) | Get-Random -Count 5 | ForEach-Object {[char]$_}))
	do {
		# Query the UAL using back-end paging until all results are returned
		$rCount = $results.Count
		$results += Search-UnifiedAuditLog -UserIds $UserPrincipalName -Operations $operations -StartDate $StartDate -EndDate $EndDate -Formatted -SessionCommand ReturnLargeSet -SessionId $sId -ResultSize 500
	} until ($rCount -eq $results.Count)
	return $results
}

#region Persistence Check Functions
function Get-MailboxRules {
	Write-ProgressHelper -Activity 'Checking Inbox rules'
	# Get inbox rules that forward/redirect or delete
	$rules = Get-InboxRule -Mailbox $UserPrincipalName | Where-Object {($_.ForwardTo -or $_.ForwardAsAttachmentTo -or $_.RedirectTo -or $_.DeleteMessage -or $_.SoftDeleteMessage)}
	if ($rules) {
		# Determine if rules were created/modified during search window
		$auditLogOwaRules = $ualResults | Where-Object {$_.Operations -in @('New-InboxRule','Set-InboxRule')} | Sort-Object -Property Identity -Unique
		$auditLogMapiRules = $ualResults | Where-Object {$_.Operations -in @('UpdateInboxRules')} | Sort-Object -Property Identity -Unique
		foreach ($ruleEntry in $rules) {
			$ruleAction = @()
			if ($ruleEntry.ForwardTo) {
				$ruleAction += 'ForwardTo:' + $ruleEntry.ForwardTo[0].Substring(0,$ruleEntry.ForwardTo[0].IndexOf(' ')).Replace('"','')
			}
			if ($ruleEntry.RedirectTo) {
				$ruleAction += 'RedirectTo:' + $ruleEntry.RedirectTo[0].Substring(0,$ruleEntry.RedirectTo[0].IndexOf(' ')).Replace('"','')
			}
			if ($ruleEntry.ForwardAsAttachmentTo) {
				$ruleAction += 'ForwardAsAttachmentTo:' + $ruleEntry.ForwardAsAttachmentTo[0].Substring(0,$ruleEntry.ForwardAsAttachmentTo[0].IndexOf(' ')).Replace('"','')
			}
			if ($ruleEntry.DeleteMessage) {
				$ruleAction += 'DeleteMessage'
			}
			if ($ruleEntry.SoftDeleteMessage) {
				$ruleAction += 'SoftDeleteMessage'
			}
			if ($auditLogOwaRules) { # Rules edited in OWA
				foreach ($logEntry in $auditLogOwaRules) {
					$runDate = $null
					# Check if matching rule name exists in audit log
					$auditData = $logEntry.AuditData | ConvertFrom-Json
					$ruleName = $auditData.ObjectId.Substring($auditData.ObjectId.LastIndexOf('\') + 1)
					if ($ruleEntry.Name -eq $ruleName) {
						$runDate = ([datetime]$logEntry.CreationDate).ToLocalTime() # CreationDate is stored in UTC
						break
					}
				}
			}
			elseif ($auditLogMapiRules) { # Rules edited in Outlook MAPI
				foreach ($logEntry in $auditLogMapiRules) {
					$runDate = $null
					$auditData = $logEntry.AuditData | ConvertFrom-Json
					$ruleName = $auditData.OperationProperties | Where-Object {$_.Name -eq 'RuleName'} | Select-Object -ExpandProperty Value
					if ($ruleEntry.Name -eq $ruleName) {
						$runDate = ([datetime]$auditData.CreationTime).ToLocalTime() # CreationTime is stored in UTC
						break
					}
				}
			}
			[PSCustomObject]@{
            	Check = 'Inbox Rule'
				User = $UserPrincipalName
            	RuleName = $ruleEntry.Name
				RuleAction = $ruleAction
				Enabled = $ruleEntry.Enabled
				Date = if ($runDate) {$runDate} else {'OutsideOfSearchWindow'}
			}
        }
	}
}

function Get-OWAForwarding {
	Write-ProgressHelper -Activity 'Checking OWA (SMTP) forwarding'
	# Check for mail forwarding via ForwardingSMTPAddress
	$mb = Get-Mailbox -Identity $UserPrincipalName
	if ($mb.ForwardingSmtpAddress) {
		$checkForwardOutput = [PSCustomObject] @{
			Check = 'SMTP Forwarding'
			User = $UserPrincipalName
			ForwardingAddress = $mb.ForwardingSmtpAddress
			Date = $null
		}
		# Search audit log to see if forwarding was set during search window
		$auditLogForward = $ualResults | Where-Object {$_.Operations -eq 'Set-Mailbox'} | Sort-Object -Property Identity -Unique | Sort-Object -Property CreationDate -Descending
		if ($auditLogForward) {
			foreach ($logEntry in $auditLogForward) {
				$auditData = $logEntry.AuditData | ConvertFrom-Json
				$propValue = $auditData.Parameters | Where-Object {$_.Name -eq 'ForwardingSmtpAddress'} | Select-Object -ExpandProperty Value
				if ($propValue -eq $mb.ForwardingSmtpAddress) {
					$runDate = ([datetime]$logEntry.CreationDate).ToLocalTime() # CreationDate is stored in UTC
					break
				}
			}
		}
		$checkForwardOutput.Date = if ($runDate) {$runDate} else {'OutsideOfSearchWindow'}
		$checkForwardOutput
	}
}

function Get-UserConsents {
	Write-ProgressHelper -Activity 'Checking user-based application consents'
	$operations = @('Consent to application.')
	$auditConsent = $ualResults | Where-Object {$_.Operations -in $operations} | Sort-Object -Property Identity -Unique
	if ($auditConsent) {
		foreach ($consent in $auditConsent) {
			$auditData = $consent.AuditData | ConvertFrom-Json
			if (($auditData.ModifiedProperties | Where-Object {$_.Name -eq 'ConsentContext.IsAdminConsent'} | Select-Object -ExpandProperty newValue) -eq $false) {
				[PSCustomObject] @{
					Check = 'User Consent'
					User = $UserPrincipalName
					ApplicationName = $auditData.Target | Where-Object {$_.Type -eq 1} | Select-Object -ExpandProperty ID # Type 1 is the app display name
					ApplicationId = $auditData.ObjectId
					Date = ([datetime]$consent.CreationDate).ToLocalTime() # CreationDate is stored in UTC
				}
			}
		}
	}
}

function Get-ClientSecretsAdded {
	Write-ProgressHelper -Activity 'Checking for certificates and client secrets added to applications'
	$operations = @('Update application – Certificates and secrets management ')
	foreach ($update in ($ualResults | Where-Object {$_.Operations -in $operations} | Sort-Object -Property Identity -Unique)) {
		$auditData = $update.AuditData | ConvertFrom-Json
		# To know if a key has been added or removed, compare the old and new key values
		$keys = $auditData.ModifiedProperties | Where-Object {$_.Name -eq 'KeyDescription'}
		# Convert string of keys into array and filter out non-key items
		$oldKeys = ($keys | Select-Object -ExpandProperty oldValue).Trim('[', ']').Trim('"').Split('][') | Where-Object {$_ -like 'Key*'}
		$newKeys = ($keys | Select-Object -ExpandProperty newValue).Trim('[', ']').Trim('"').Split('][') | Where-Object {$_ -like 'Key*'}
		$changedKeys = Compare-Object -ReferenceObject $oldKeys -DifferenceObject $newKeys | Where-Object {$_.SideIndicator -eq '=>'}
		foreach ($key in $changedKeys.InputObject) {
			# Determine type of key that was added
			switch (($key.Split(',') | Where-Object {$_ -like "KeyType*"}).Split('=')[1]) {
				Password {$keyType = 'ClientSecret'}
				AsymmetricX509Cert {$keyType = 'Certificate'}
			}
			[PSCustomObject] @{
				Check = 'App Cert/Secret Added'
				User = $UserPrincipalName
				ApplicationName = $auditData.Target | Where-Object {$_.Type -eq 1} | Select-Object -ExpandProperty ID
				AppObjectId = ($auditData.Target | Where-Object {$_.ID -like 'Application_*'} | Select-Object -ExpandProperty ID).Substring(12)
				KeyType = $keyType
				KeyName = ($key.Split(',') | Where-Object {$_ -like "DisplayName*"}).Split('=')[1]
				Date = ([datetime]$update.CreationDate).ToLocalTime() # CreationDate is stored in UTC
			}
		}
	}
}

function Get-FolderPermissionChanges {
	Write-ProgressHelper -Activity 'Checking mailbox folder permissions'
	$mailbox = Get-Mailbox -Identity $UserPrincipalName
	# Verify mailbox auditing requirements have been met
	$orgAuditState = (Get-OrganizationConfig).AuditDisabled
	$mailboxOwnerAudit = $mailbox.AuditOwner -contains 'UpdateFolderPermissions'
	$userAuditBypass = (Get-MailboxAuditBypassAssociation -Identity $UserPrincipalName).AuditBypassEnabled
	Write-Verbose -Message "Org-level mailbox auditing disabled: $orgAuditState; Owner auditing includes UpdateFolderPermissions: $mailboxOwnerAudit; User audit bypass enabled: $mailboxAuditBypass"
	if ($orgAuditState -eq $false -and $mailboxOwnerAudit -and $userAuditBypass -eq $false) {
		# Checking folder permission additions includes calendar delegate additions,
		# so there is no need to check for UpdateCalendarDelegation operation
		$operations = @('AddFolderPermissions','ModifyFolderPermissions','RemoveFolderPermissions')
		$folderAuditLogs = $ualResults | Where-Object {$_.Operations -in $operations} | Sort-Object -Property Identity -Unique
		if ($folderAuditLogs) {
			foreach ($logEntry in $folderAuditLogs) {
				$auditData = $logEntry.AuditData | ConvertFrom-Json
				[PSCustomObject] @{
					Check = 'Mailbox Folder Permission Change'
					User = $UserPrincipalName
					FolderName = "$($auditData.Item.ParentFolder.Name) ($($auditData.Item.ParentFolder.Path.Replace('\\','\')))"
					Action = $logEntry.Operations
					Assignee = $auditData.Item.ParentFolder.MemberUpn
					Permission = $auditData.Item.ParentFolder.MemberRights
					Date = ([datetime]$logEntry.CreationDate).ToLocalTime() # CreationDate is stored in UTC
				}
			}
		}
	}
	else {
		if ($orgAuditState -eq $true) {
			Write-Warning -Message 'Changes to mailbox folder permissions was skipped because organization-level mailbox auditing is disabled.'
		}
		if ($mailboxOwnerAudit -eq $false) {
			Write-Warning -Message 'Changes to mailbox folder permissions was skipped because mailbox owner auditing of folder permission changes is not enabled.'
		} 
		if ($userAuditBypass -eq $true) {
			Write-Warning -Message "Changes to mailbox folder permissions was skipped because mailbox audit bypass is enabled for $UserPrincipalName."
		}
	}
}

function Get-CalendarPublishing {
	Write-ProgressHelper -Activity 'Checking anonymous calendar publishing'
	# Get localized folder name
	$primaryCalendarPath = Get-MailboxFolderStatistics -Identity $UserPrincipalName -FolderScope Calendar | 
		Where-Object {$_.FolderType -eq 'Calendar'} | Select-Object -ExpandProperty FolderPath
	$calendarPublishing = Get-MailboxCalendarFolder -Identity "$($UserPrincipalName):\$($primaryCalendarPath.Substring(1))"
	if ($calendarPublishing.PublishEnabled) {
		$checkCalPubOutput = [PSCustomObject] @{
			Check = 'Calendar Publishing'
			User = $UserPrincipalName
			PublishEnabled = $calendarPublishing.PublishEnabled
			Date = $null
		}
		# Determine if publishing enabled by owner during search window
		$operations = @('Set-MailboxCalendarFolder')
		$auditLogPublish = $ualResults | Where-Object {$_.Operations -in $operations} | Sort-Object -Property Identity -Unique
		if ($auditLogPublish) {
			foreach ($logEntry in $auditLogPublish) {
				$auditData = $logEntry.AuditData | ConvertFrom-Json
				if (($auditData.Parameters | Where-Object {$_.Name -eq 'PublishEnabled'} | Select-Object -ExpandProperty Value) -eq $true) {
					$checkCalPubOutput.Date = ([datetime]$logEntry.CreationDate).ToLocalTime() # CreationDate is returned in UTC
					break
				}
				else {
					$checkCalPubOutput.Date = 'OutsideOfSearchWindow'
				}
			}
		}
		else {
			$checkCalPubOutput.Date = 'OutsideOfSearchWindow'
		}
		$checkCalPubOutput
	}
}

function Get-MobileDevices {
	Write-ProgressHelper -Activity 'Checking for new/active mobile device partnerships'
	$mobileDevices = Get-MobileDeviceStatistics -Mailbox $UserPrincipalName
	if ($mobileDevices) {
		foreach ($device in $mobileDevices) {
			if ($device.FirstSyncTime -gt $StartDate -or $device.LastSyncTime -gt $StartDate) {
				[PSCustomObject] @{
					Check = 'Mobile Device'
					User = $UserPrincipalName
					DeviceName = $device.DeviceFriendlyName
					DeviceAgent = $device.DeviceUserAgent
					FirstSync = $device.FirstSyncTime
					LastSync = $device.LastSyncTime
				}
			}
		}
	}
}

function Get-FileSharing {
	Write-ProgressHelper -Activity 'Checking file sharing'
	$operations = @('CompanyLinkCreated', 'SecureLinkCreated', 'AnonymousLinkCreated', 'AnonymousLinkUpdated', 'AddedToSecureLink','AddedToSharingLink')
	$auditLinks = $ualResults | Where-Object {$_.Operations -in $operations} | Sort-Object -Property Identity -Unique | Sort-Object -Property CreationDate
	if ($auditLinks) {
		foreach ($link in $auditLinks) {
			$auditData = $link.AuditData | ConvertFrom-Json
			if ($auditData.Operation -eq 'AddedToSecureLink' -or $auditData.Operation -eq 'AddedToSharingLink') {
				if ($auditData.TargetUserOrGroupName -like '*#EXT#*') {
					$recipient = $auditData.TargetUserOrGroupName.Substring(0,$auditData.TargetUserOrGroupName.IndexOf('#EXT#',[System.StringComparison]::InvariantCultureIgnoreCase)) -replace ('_','@')
				}
				else {
					$recipient = $auditData.TargetUserOrGroupName
				}
			} else {
				$recipient = $null
			}

			[PSCustomObject] @{
				Check = 'File Sharing'
				User = $UserPrincipalName
				Operation = $auditData.Operation
				FilePath = $auditData.ObjectId
				Recipient = $recipient
				Date = ([datetime]$link.CreationDate).ToLocalTime() # CreationDate is stored in UTC
			}
		}
	}
}

function Get-TeamMemberAdded {
	Write-ProgressHelper -Activity 'Checking Team member additions'
	$operations = @('MemberAdded')
	$auditMembers = $ualResults | Where-Object {$_.Operations -in $operations} | Sort-Object -Property Identity -Unique
	if ($auditMembers) {
		foreach ($entry in $auditMembers) {
			$auditData = $entry.AuditData | ConvertFrom-Json
			foreach ($user in $auditData.Members) {
				[PSCustomObject] @{
					Check = 'Team Member Add'
					User = $UserPrincipalName
					Operation = $entry.Operations
					Member = if ($user.UPN -like '*#EXT#*') {$user.UPN.Substring(0,$user.UPN.IndexOf('#EXT#',[System.StringComparison]::InvariantCultureIgnoreCase)) -replace ('_','@')} else {$user.UPN}
					Team = $auditData.TeamName
					Date = ([datetime]$entry.CreationDate).ToLocalTime() # CreationDate is stored in UTC
				}
			}
		}
	}
}
	
function Get-GroupMemberAdded {
	Write-ProgressHelper -Activity 'Checking Microsoft 365 Group member additions'
	$operations = @('AddMemberToGroup.','Add owner to group.')
	$addedMemberLogs = $ualResults | Where-Object {$_.Operations -in $operations} | Sort-Object -Property Identity -Unique | Sort-Object -Property CreationDate
	if ($addedMemberLogs) {
		foreach ($entry in $addedMemberLogs) {
			$auditData = $entry.AuditData | ConvertFrom-Json
			[PSCustomObject] @{
				Check = 'Group Member Add'
				User = $UserPrincipalName
				Operation = $entry.Operations
				Member = if ($auditData.ObjectId -like '*#EXT#*') {$auditData.ObjectId.Substring(0,$auditData.ObjectId.IndexOf('#EXT#',[System.StringComparison]::InvariantCultureIgnoreCase)) -replace ('_','@')} else {$auditData.ObjectId}
				Group = $auditData.ModifiedProperties | Where-Object {$_.Name -eq 'Group.DisplayName'} | Select-Object -ExpandProperty newValue
				Date = ([datetime]$entry.CreationDate).ToLocalTime() # CreationDate is stored in UTC
			}
		}
	}
}
	
function Get-UpdatedFlows {
	# Details of workflows via PowerShell or Graph is non-existent, but an entry from the audit log
	# includes the admin URL of the workflow (for manual further review) as well as the connectors being used
	Write-ProgressHelper -Activity 'Checking Power Automate workflows'
	$operations = @('EditFlow','CreateFlow')
	$auditFlows = $ualResults | Where-Object {$_.Operations -in $operations} | Sort-Object -Property Identity -Unique | Sort-Object -Property CreationDate
	if ($auditFlows) {
		foreach ($entry in $auditFlows)	{
			[PSCustomObject] @{
				Check = 'Power Automate'
				User = $UserPrincipalName
				Operation = $auditData.Operation
				FlowUrl = $auditData.FlowDetailsUrl
				Connectors = $auditData.FlowConnectorNames
				Date = ([datetime]$entry.CreationDate).ToLocalTime() # CreationDate is stored in UTC
			}
		}
	}
}

function Get-OutlookAddIns {
	Write-ProgressHelper -Activity 'Checking for custom Mail add-ins'
	# Add-ins that have same properties as side-loaded but are default add-ins
	# Poll/Polls,Share to Teams,OneNote/Send to OneNote
	$ignoreAddIn = @('afde34e6-58a4-4122-8a52-ef402180a878','545d8236-721a-468f-85d8-254eca7cb0da','6b47614e-0125-454b-9f76-bd5aef85ac7b')
	# Side-loaded add-ins have a Type of Private and Scope of User
	$addins = Get-App -Mailbox $UserPrincipalName | Where-Object {$_.Type -eq 'Private' -and $_.Scope -eq 'User' -and $_.AppId -notin $ignoreAddIn}
	if ($addins) {
		foreach ($addin in $addins) {
			[PSCustomObject] @{
				Check = 'Custom Mail Add-In'
				User = $UserPrincipalName
				AddInName = $addin.DisplayName
				Permission = $addin.Requirements
			}
		}
	}
}

function Get-SafeSenderList {
	Write-ProgressHelper -Activity 'Checking for modified Safe Senders List'
	$operations = @('Set-MailboxJunkEmailConfiguration')
	$auditSafeSenders = $ualResults | Where-Object {$_.Operations -in $operations} | Sort-Object -Property Identity -Unique
	if ($auditSafeSenders) {
		foreach ($entry in $auditSafeSenders) {
			$auditData = $entry.AuditData | ConvertFrom-Json
			if ($auditData.Parameters.Name -contains 'TrustedSendersAndDomains') {
				[PSCustomObject] @{
					Check = 'Safe Senders List'
					User = $UserPrincipalName
					Operation = $entry.Operations
					SafeSenders = $auditData.Parameters.Value[[array]::IndexOf($auditData.Parameters.Name,'TrustedSendersAndDomains')] # List is returned as semi-colon-separated string
					Date = ([datetime]$entry.CreationDate).ToLocalTime() # CreationDate is stored in UTC
				}
			}
		}
	}
}
	
function Get-PowerApps {
	Write-ProgressHelper -Activity 'Checking PowerApps'
	$operations = @('PublishPowerApp','PowerAppPermissionEdited')
	$auditPowerApps = $ualResults | Where-Object {$_.Operations -in $operations} | Sort-Object -Property Identity -Unique | Sort-Object -Property CreationDate
	if ($auditPowerApps) {
		foreach ($entry in $auditPowerApps) {
			$auditData = ($entry.AuditData | ConvertFrom-Json).JsonPropertiesCollection | ConvertFrom-Json
			if ($entry.Operations -eq 'PowerAppPermissionEdited') {
				[PSCustomObject] @{
					Check = 'Power Apps'
					User = $UserPrincipalName
					Operation = $entry.Operations
					AppId = $auditData.'powerplatform.analytics.resource.power_app.id'.SubString($auditData.'powerplatform.analytics.resource.power_app.id'.LastIndexOf('/')+1)
					Assignee = $auditData.'targetuser.id'
					PermissionType = $auditData.'powerplatform.analytics.resource.permission_type'
					Date = ([datetime]$entry.CreationDate).ToLocalTime() # CreationDate is stored in UTC
				}
			} elseif ($entry.Operations -eq 'PublishPowerApp') {
				[PSCustomObject] @{
					Check = 'Power Apps'
					User = $UserPrincipalName
					Operation = $entry.Operations
					AppId = $auditData.'powerplatform.analytics.resource.power_app.id'
					AppName = $auditData.'powerplatform.analytics.resource.power_app.display_name'
					EnvironmentName = $auditData.'powerplatform.analytics.resource.environment.name'
					Date = ([datetime]$entry.CreationDate).ToLocalTime() # CreationDate is stored in UTC
				}
			}
		}
	}
}

function Get-ExAdminPersistence {
	Write-ProgressHelper -Activity 'Checking for persistence configuration by Exchange administrator'
	$operations = @('New-JournalRule','Set-JournalRule','New-TransportRule','Set-TransportRule','Add-MailboxPermission')
	$auditExAdmin = $ualResults | Where-Object {$_.Operations -in $operations} | Sort-Object -Property Identity -Unique
	if ($auditExAdmin) {
		foreach ($entry in $auditExAdmin) {
			$auditData = $entry.AuditData | ConvertFrom-Json
			if ($entry.Operations -in @('New-JournalRule','Set-JournalRule')) {
				[PSCustomObject] @{
					Check = 'Exchange Admin Persistence'
					User = $UserPrincipalName
					Operation = $entry.Operations
					RuleName = $auditData.ObjectId
					JournalRecipient = $auditData.Parameters | Where-Object {$_.Name -eq 'JournalEmailAddress'} | Select-Object -ExpandProperty Value
					Date = ([datetime]$entry.CreationDate).ToLocalTime() # CreationDate is stored in UTC
				}
			}
			if ($entry.Operations -in @('New-TransportRule','Set-TransportRule')) {
				[PSCustomObject] @{
					Check = 'Exchange Admin Persistence'
					User = $UserPrincipalName
					Operation = $entry.Operations
					RuleName = $auditData.ObjectId
					Date = ([datetime]$entry.CreationDate).ToLocalTime() # CreationDate is stored in UTC
				}
			}
			if ($entry.Operations -eq 'Add-MailboxPermission') {
				[PSCustomObject] @{
					Check = 'Exchange Admin Persistence'
					User = $UserPrincipalName
					Operation = $entry.Operations
					Mailbox = $auditData.ObjectId
					AssigneeId = $auditData.Parameters | Where-Object {$_.Name -eq 'User'} | Select-Object -ExpandProperty Value # This is the Entra object ID
					Date = ([datetime]$entry.CreationDate).ToLocalTime() # CreationDate is stored in UTC
				}
			}
		}
	}
}

function Get-AdminConsentGranted {
	Write-ProgressHelper -Activity 'Checking for admin consent grants'
	$operations = @('Consent to application.')
	$auditConsent = $ualResults | Where-Object {$_.Operations -in $operations} | Sort-Object -Property Identity -Unique
	if ($auditConsent) {
		foreach ($consent in $auditConsent) {
			$auditData = $consent.AuditData | ConvertFrom-Json
			if (($auditData.ModifiedProperties | Where-Object {$_.Name -eq 'ConsentContext.IsAdminConsent'} | Select-Object -ExpandProperty newValue) -eq $true) {
				[PSCustomObject] @{
					Check = 'Admin Consent'
					User = $UserPrincipalName
					ApplicationName = $auditData.Target | Where-Object {$_.Type -eq 1} | Select-Object -ExpandProperty ID # Type 1 is the app display name
					ApplicationId = $auditData.ObjectId
					Date = ([datetime]$consent.CreationDate).ToLocalTime() # CreationDate is stored in UTC
				}
			}
		}
	}
}
#endregion Persistence Check Functions

#region Start

$step = 0
$totalSteps = 16 # Includes getting logs (but excluding mobile devices)

$ualResults = Get-UALData
Get-MailboxRules
Get-OWAForwarding
Get-FolderPermissionChanges
Get-CalendarPublishing
Get-OutlookAddIns
Get-SafeSenderList
#Get-MobileDevices
Get-UserConsents
Get-FileSharing
Get-TeamMemberAdded
Get-GroupMemberAdded
Get-UpdatedFlows
Get-PowerApps
Get-ExAdminPersistence
Get-ClientSecretsAdded
Get-AdminConsentGranted

#endregion Start