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

#Requires -Version 4
# This script will require registration of a Web Application in Application Registration Portal (see https://apps.dev.microsoft.com/)

<#
	.SYNOPSIS
		Script to parse DMARC reports from an Exchange Online mailbox, denormalize and write them to a SQL database.
        This script is intended for small scale DMARC reporting usage. It has not been tested in production.

	.DESCRIPTION
		Reads config file,
		Requests 1-hour access token from AAD,
		Access the Exchange Online mailbox using Rest API,
		Loops through all messages in the inbox
			Emails with no attachments are moved to the SkippedNoAttachment folder
			If the attachments is not a ZIP file the email is moved to the Skipped folder
			If the attachment is ZIP, all xml files are parsed and denormalized
		The Geo-IP location is looked up using freegeoip.net API and the data added.
		Data is saved on an Azure SQL table or displayed on the console.

		See GitHub Readme at 

        Requires Microsoft.Identity.Client (MSAL) libraries in the script root.

	.EXAMPLE
		PS C:\> .\DMARCParser.ps1

	.OUTPUTS
		System.String

	.NOTES
		Andres Canello
		Senior Engineer - Microsoft
		andres.canello@microsoft.com

		Cam Murray
		Field Engineer - Microsoft
		cam.murray@microsoft.com
		
		For updates, and more scripts, visit https://github.com/O365AES/Scripts
		
		Last update: 29 March 2017

	.LINK
		about_functions_advanced

#>


[System.Reflection.Assembly]::LoadWithPartialName('System.IO.compression') | Out-Null

#region Functions

Function Read-ConfigFile {
	try {
        $tmp = Get-Content $PSScriptRoot\DMARCParser.config -ErrorAction Stop -Raw | ConvertFrom-Json
    }
    catch {
        Write-Host 'Error reading DMARCParser.config file'
        exit
    }
   return $tmp
}

Function Get-FolderByName {
	Param(
		[string]$FolderName
	)
	$folder = $null
	$folders = Invoke-RestMethod ($apiUrl + '/MailFolders/Inbox/childfolders') -Headers $headers
	$folder = $folders.value | Where-Object {$_.DisplayName -eq $FolderName}
	return $folder   
}

Function Get-FolderByNameOrCreate {
	Param(
		[string]$FolderName
	)

    $folders = Invoke-RestMethod ($apiUrl + '/MailFolders/Inbox/childfolders') -Headers $headers
	$folder = $folders.value | Where-Object {$_.DisplayName -eq $FolderName}

    if(!$folder) {
        # Folder doesnt exist, so create it
        $body = @{DisplayName=$FolderName} | ConvertTo-Json
        $folder = Invoke-RestMethod -Method Post -uri ($apiUrl + '/MailFolders/Inbox/childfolders') -Headers $headers -Body $body  -ContentType 'application/json'
    }

	return $folder   
}

Function New-Folder {
	Param(
		[parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$FolderName
	)
	$body = @{DisplayName=$FolderName} | ConvertTo-Json

    $NewFolder = Invoke-RestMethod -Method Post -uri ($apiUrl + '/MailFolders/Inbox/childfolders') -Headers $headers -Body $body  -ContentType 'application/json'
	return $NewFolder
}

Function Get-Messages {
	Param(
		[parameter(Mandatory=$true)]
        [ValidateSet('TopTenNoAttachment', 'TopTen')]
		[string]$filter
	)
    switch ($filter) {       
					'TopTenNoAttachment' {
                    $response = Invoke-RestMethod ($apiUrl + '/MailFolders/Inbox/messages?$filter=HasAttachments%20eq%20false&$top=3&$select=HasAttachments,Subject') -Headers $headers
                    return $response
					}
    		        'TopTen' {
                        $response = Invoke-RestMethod ($apiUrl + '/MailFolders/Inbox/messages?top=10') -Headers $headers
                        return $response
                    }   
	}
}

Function Move-Message {
	Param(
		[parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
		[string]$MessageId,
		
		[parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
		[string]$TargetFolderId
	)
    $fld = @{DestinationId=$TargetFolderId} | ConvertTo-Json
	$response = Invoke-RestMethod -Method Post -uri "$apiUrl/messages/$MessageId/move" -Body $fld -Headers $headers -ContentType 'application/json'
	return $response
}

Function Get-Attachments {
	Param(
		[parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]		
		[string]$MessageId
    )
	$response = Invoke-RestMethod -uri "$apiUrl/MailFolders/Inbox/messages/$MessageId/attachments" -Headers $headers
	return $response
}

Function DecompressZip {
	Param(
		$ContentBytes
	)
	$tmp = [system.convert]::FromBase64String($ContentBytes)
	$stream = new-object System.IO.MemoryStream($tmp,$false)
	$zip = new-object System.IO.Compression.ziparchive($stream,[System.IO.Compression.CompressionMode]::Decompress)
	return $zip.Entries
}

Function Get-XMLFilesFromZIP {
    Param(
		$files
	)
    $XMLs = @()
    foreach ($file in $files) {        
		if ($file.fullname -ilike '*.xml') {
			$XMLStream = $File.Open()
			$reader = new-object System.IO.StreamReader($XMLStream)
			$string = $reader.ReadToEnd()
			[xml]$xml = $string
			$XMLs += $xml
		}
	}
    return $XMLs
}

Function Get-XMLFiles {
	Param(
		$attachments
    )
    $AllXMLs = @()
    foreach ($file in $attachments) {
		if ($file.ContentType -in ('application/zip','application/x-zip-compressed')) {
			$files = DecompressZip -ContentBytes $file.ContentBytes
			$XMLs = Get-XMLFilesFromZIP -files $files
			$AllXMLs += $XMLs
		}
	}
    return $AllXMLs   
}

Function New-DMARCEntrySchema {
	New-Object -TypeName psobject -Property @{
		ReportReceivedDate=''
		ReportMetadataOrgName=''
		ReportMetadataEmail=''
		ReportMetadataReportId=''
		ReportMetadataDateRangeBegin=''
		ReportMetadataDateRangeEnd=''
		PolicyPublishedDomain=''
		PolicyPublishedDKIM=''
		PolicyPublishedSPF=''
		PolicyPublishedP=''
		PolicyPublishedSP=''
		PolicyPublishedPCT=''
		SourceIP=''
		SourceIPCountry=''
		SourceIPRegion=''
		SourceIPCity=''
		SourceIPLatitude=''
		SourceIPLongitude=''
		Count=''
		PolicyEvaluatedDisposition=''
		PolicyEvaluatedDKIM=''
		PolicyEvaluatedSPF=''
		IdentifiersHeaderFrom=''
		AuthResultsSPFDomain=''
		AuthResultsSPFResult=''
		AuthResultsDKIMDomain=''
		AuthResultsDKIMResult=''
	}
}

Function Read-DMARCReport {
	Param(
		[parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]	
		[xml]$report,
		
		[ValidateNotNullOrEmpty()]
		$reportReceivedDate
	)
	$tmpEntries = @()
	foreach ($record in $report.feedback.record) {
		foreach ($row in $record.row) {
			# Get IP Location info
			$IPLocation = Get-IPLocation -IP $row.source_ip
			# Get an empty entry
			$entry = New-DMARCEntrySchema
			# Populate the data
			$entry.ReportReceivedDate=$reportReceivedDate
			$entry.ReportMetadataOrgName=$report.feedback.report_metadata.org_name
			$entry.ReportMetadataEmail=$report.feedback.report_metadata.email
			$entry.ReportMetadataReportId=$report.feedback.report_metadata.report_id
			$entry.ReportMetadataDateRangeBegin=(Convert-UNIXTime -UNIXTime $report.feedback.report_metadata.date_range.begin)
			$entry.ReportMetadataDateRangeEnd=(Convert-UNIXTime -UNIXTime $report.feedback.report_metadata.date_range.end)
			$entry.PolicyPublishedDomain=$report.feedback.policy_published.domain
			$entry.PolicyPublishedDKIM=$report.feedback.policy_published.adkim
			$entry.PolicyPublishedSPF=$report.feedback.policy_published.aspf
			$entry.PolicyPublishedP=$report.feedback.policy_published.p
			$entry.PolicyPublishedSP=$report.feedback.policy_published.sp
			$entry.PolicyPublishedPCT=$report.feedback.policy_published.pct
			$entry.SourceIP=$row.source_ip
			$entry.SourceIPCountry=$IPLocation.country_name
			$entry.SourceIPRegion=$IPLocation.region_name
			$entry.SourceIPCity=$IPLocation.city
			$entry.SourceIPLatitude=$IPLocation.latitude
			$entry.SourceIPLongitude=$IPLocation.longitude
			$entry.Count=$row.count
			$entry.PolicyEvaluatedDisposition=$row.policy_evaluated.disposition
			$entry.PolicyEvaluatedDKIM=$row.policy_evaluated.dkim
			$entry.PolicyEvaluatedSPF=$row.policy_evaluated.spf
			$entry.IdentifiersHeaderFrom=$record.identifiers.header_from
			$entry.AuthResultsSPFDomain=$record.auth_results.spf.domain
			$entry.AuthResultsSPFResult=$record.auth_results.spf.result
			$entry.AuthResultsDKIMDomain=$record.auth_results.dkim.domain
			$entry.AuthResultsDKIMResult=$record.auth_results.dkim.result
			$tmpEntries += $entry
		}
	}
	return $tmpEntries
}

Function Invoke-SQLQuery {
	Param(
		[parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]	
		[string]$Query
	)

    # Construct the SQL Connection String
    $ConnectionString = "server=" + $($config.SqlHostname) + ",1433; uid=" + $($config.SqlUsername) + "; pwd=" + $($config.SqlPass) + "; database=" + $($config.SqlDb)

    # Attempt to write to SQL
    Try {
		$Connection = New-Object System.Data.SqlClient.SqlConnection
		$Connection.ConnectionString = $ConnectionString
		$Connection.Open()
		$Command = New-Object System.Data.SqlClient.SqlCommand($Query, $Connection)
		$DataAdapter = New-Object System.Data.SqlClient.SqlDataAdapter($Command)
		$DataSet = New-Object System.Data.DataSet
		$RecordCount = $dataAdapter.Fill($dataSet, 'data')
		$DataSet.Tables[0]
	} Catch {
        Write-Error "Unable to run query : $query"
	} Finally {
		$Connection.Close()
	}
}

Function Write-SQL {
	Param(
		[parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]	
		$entries
	)
	#Make sure we've got tables
	$TableCreateQuery = "if not exists (select * from sysobjects where name ='dmarc' and xtype='U') CREATE TABLE dmarc (ReportReceivedDate DATETIME, ReportMetadataOrgName TEXT, ReportMetadataEmail TEXT, ReportMetadataReportId TEXT, ReportMetadataDateRangeBegin DATETIME, ReportMetadataDateRangeEnd DATETIME, PolicyPublishedDomain TEXT, PolicyPublishedDKIM TEXT, PolicyPublishedSPF TEXT, PolicyPublishedP TEXT, PolicyPublishedSP TEXT, PolicyPublishedPCT INT, SourceIP TEXT, SourceIPCountry TEXT, SourceIPRegion TEXT, SourceIPCity TEXT, SourceIPLatitude TEXT, SourceIPLongitude TEXT, Count INT, PolicyEvaluatedDisposition TEXT, PolicyEvaluatedDKIM TEXT, PolicyEvaluatedSPF TEXT, IdentifiersHeaderFrom TEXT, AuthResultsSPFDomain TEXT, AuthResultsSPFResult TEXT, AuthResultsDKIMDomain TEXT, AuthResultsDKIMResult TEXT, ID BIGINT NOT NULL IDENTITY (1,1))"
	Invoke-SQLQuery -Query $TableCreateQuery

    #Loop entries to add to SQL
	foreach ($entry in $entries) {
		#Construct the query to a valid string, then execute
		#Need to handle the parameters object by converting to a string.

		$InsertQuery = "BEGIN INSERT INTO dmarc (ReportReceivedDate, ReportMetadataOrgName, ReportMetadataEmail, ReportMetadataReportId, ReportMetadataDateRangeBegin, ReportMetadataDateRangeEnd, PolicyPublishedDomain, PolicyPublishedDKIM, PolicyPublishedSPF, PolicyPublishedP, PolicyPublishedSP, PolicyPublishedPCT, SourceIP, SourceIPCountry, SourceIPRegion, SourceIPCity, SourceIPLatitude, SourceIPLongitude, Count, PolicyEvaluatedDisposition, PolicyEvaluatedDKIM, PolicyEvaluatedSPF, IdentifiersHeaderFrom, AuthResultsSPFDomain, AuthResultsSPFResult, AuthResultsDKIMDomain, AuthResultsDKIMResult) VALUES ('" + $entry.ReportReceivedDate + "', '" + $entry.ReportMetadataOrgName + "', '" + $entry.ReportMetadataEmail + "', '" + $entry.ReportMetadataReportId + "', '" + $entry.ReportMetadataDateRangeBegin + "', '" + $entry.ReportMetadataDateRangeEnd + "', '" + $entry.PolicyPublishedDomain + "', '" + $entry.PolicyPublishedDKIM + "', '" + $entry.PolicyPublishedSPF + "', '" + $entry.PolicyPublishedP + "', '" + $entry.PolicyPublishedSP + "', '" + $entry.PolicyPublishedPCT + "', '" + $entry.SourceIP + "', '" + $entry.SourceIPCountry + "', '" + $entry.SourceIPRegion + "', '" + $entry.SourceIPCity + "', '" + $entry.SourceIPLatitude + "', '" + $entry.SourceIPLongitude + "', '" + $entry.Count + "', '" + $entry.PolicyEvaluatedDisposition + "', '" + $entry.PolicyEvaluatedDKIM + "', '" + $entry.PolicyEvaluatedSPF + "', '" + $entry.IdentifiersHeaderFrom + "', '" + $entry.AuthResultsSPFDomain + "', '" + $entry.AuthResultsSPFResult + "', '" + $entry.AuthResultsDKIMDomain + "', '" + $entry.AuthResultsDKIMResult + "') END"
		Invoke-SQLQuery -Query $InsertQuery
	}
}

Function Convert-UNIXTime {
	Param(
		[parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]	
		$UNIXTime
	)
	$date = get-date '1/1/1970'
	return $date.AddSeconds($UNIXTime)
}

Function Get-IPLocation {
	Param(
		[parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
		$IP
	)
	$response = Invoke-RestMethod -Method Get -Uri "http://freegeoip.net/json/$IP"
	return $response
}

#endregion

#region Main

$config = Read-ConfigFile

# MSAL Authentication
Add-Type -path "$PSScriptRoot\Microsoft.Identity.Client.dll"

# Obtain a larger scope for shared mailboxes if we are reading from another mailbox.
# Adjust the apiUrl for different context if accessing shared mailbox

if($config.Mailbox -ne $config.Account) {
    $scope = 'https://outlook.office.com/mail.readwrite','https://outlook.office.com/mail.readwrite.shared'
    $apiUrl = "https://outlook.office.com/api/v2.0/users/$($config.Mailbox)"
} else {
    $scope = 'https://outlook.office.com/mail.readwrite'
    $apiUrl = "https://outlook.office.com/api/v2.0/me"
}

$pca = New-Object "Microsoft.Identity.Client.PublicClientApplication" -ArgumentList $config.ClientID

# Load saved token cache
if(Test-Path("$PSScriptRoot\tokencache.dat")) {
    $tokenbytes = [io.file]::ReadAllBytes("$PSScriptRoot\tokencache.dat");
    $pca.UserTokenCache.Deserialize($tokenbytes);
}

# Select tokens which have been created in the life of the refresh token, if none exist, we can't do a silent acquisition of token
# and require user input

if(($pca.UserTokenCache.ReadItems($config.ClientID) | where-object {$_.ExpiresOn.AddDays(14) -gt (Get-Date)})) {
    $result = $pca.AcquireTokenSilentAsync($scope)
} else {
    $result = $pca.AcquireTokenAsync($scope)
}

$token = $result.Result

# Prepare header to be used in requests
$headers = New-Object 'System.Collections.Generic.Dictionary[[String],[String]]'
$headers.Add('Accept', 'application/json')
$headers.Add('Authorization', "Bearer $($token.Token)")
$headers.Add('X-AnchorMailbox', $config.mailbox)

# Check if the folders Inbox\Skipped and Inbox\Processed are there, otherwise create them
$skipped = Get-FolderByNameOrCreate -FolderName 'Skipped'
$skippedNoAttachment = Get-FolderByNameOrCreate -FolderName 'SkippedNoAttachment'
$processed = Get-FolderByNameOrCreate -FolderName 'Processed'

# Get all the messages in the Inbox with no attachments and move them to the Inbox\SkippedNoAttachment folder
$emails = Get-Messages -filter 'TopTenNoAttachment'
While ($emails.value -ne $null) {
	foreach ($email in $emails.value) {
		$moveResult = Move-Message -MessageId $email.id -TargetFolder $skippedNoAttachment.Id
    }
    $emails = Get-Messages -filter 'TopTenNoAttachment'
}

# Get messages in the Inbox with attachments
$emails = Get-Messages -filter 'TopTen'
$entries = @()
While ($emails.value -ne $null) {
	foreach ($email in $emails.value) {
		$attachments = Get-Attachments -MessageId $email.Id
		$XMLs = Get-XMLFiles -Attachments $attachments.value
		if ($XMLs) { 
			foreach ($xml in $XMLs) {
				$entries += Read-DMARCReport -report $xml -reportReceivedDate $email.ReceivedDateTime
			}
		$moveResult = Move-Message -MessageId $email.Id -TargetFolderId $processed.Id
		} else {
		$moveResult = Move-Message -MessageId $email.Id -TargetFolderId $skipped.Id
		write-host 'No XML files found in the attachments'
		}
	}
	$emails = Get-Messages -filter 'TopTen'
}

# Save token cache
[io.file]::WriteAllBytes("$PSScriptRoot\tokencache.dat", $pca.UserTokenCache.Serialize());

#This will export the data in the API to an Azure SQL instance
#write each entry, only move message to processed if all the entries were written
if (($config.Output -eq "SQL") -and $entries) { Write-Verbose "Writing to SQL"; Write-SQL $entries; }
return $entries

#endregion
