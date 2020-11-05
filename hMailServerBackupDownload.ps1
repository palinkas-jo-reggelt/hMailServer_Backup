<#

.SYNOPSIS
	hMailServer Backup Download Utility

.DESCRIPTION
	hMailServer Backup Download Utility

.FUNCTIONALITY
	Looks for last backup at LetsUpload and downloads for restoration

.PARAMETER 

	
.NOTES
	Fill in user variables below - not dependent on hMailServerBackupConfig.ps1
	
.EXAMPLE


#>

<###   FOLDER LOCATIONS   ###>
$BackupLocation = "X:\HMS-BACKUP"       # Local folder that will contain downloaded files

<###   LETSUPLOAD API VARIABLES   ###>
$APIKey1 = "1QFMyGCDgCH7BKG6ZKhxmUvAl98abP4bYiJ16iJTtLYZopqycRZJpndpca6ZgByT"
$APIKey2 = "Fky8b24HpzuYhPeXmZO8m1pe6vqcxluodasRtF1C6dnShutYkpguAlJYAWd7JgiB"

<###   BEGIN SCRIPT   ###>
$StartScript = Get-Date

<#  Load required files  #>
Try {
	.("$PSScriptRoot\hMailServerBackupFunctions.ps1")
}
Catch {
	Write-Output "$(Get-Date) -f G) : ERROR : Unable to load supporting PowerShell Scripts" | Out-File "$PSScriptRoot\PSError.log" -Append
	Write-Output "$(Get-Date) -f G) : ERROR : $Error" | Out-File "$PSScriptRoot\PSError.log" -Append
	Exit
}

<#  Clear out error variable  #>
$Error.Clear()

<#  Validate backup folder  #>
$BackupLocation = $BackupLocation -Replace('\\$','')
ValidateFolders $BackupLocation
$DownloadFolder = "$BackupLocation\hMailServer-Restoration"
If (Test-Path $DownloadFolder) {Remove-Item -Force -Path $DownloadFolder -Recurse}
New-Item -Path $DownloadFolder -ItemType Directory

<#  Delete old debug file and create new  #>
$EmailBody = "$PSScriptRoot\EmailBody.log"
If (Test-Path $EmailBody) {Remove-Item -Force -Path $EmailBody}
New-Item $EmailBody
$DebugLog = "$BackupLocation\hMailServerDownloadRestore.log"
If (Test-Path $DebugLog) {Remove-Item -Force -Path $DebugLog}
New-Item $DebugLog
Write-Output "::: hMailServer Download Offsite Backup Routine $(Get-Date -f D) :::" | Out-File $DebugLog -Encoding ASCII -Append
Write-Output " " | Out-File $DebugLog -Encoding ASCII -Append
If ($UseHTML) {
	Write-Output "
	<!DOCTYPE html PUBLIC `"-//W3C//DTD XHTML 1.0 Transitional//EN`" `"https://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd`">
	<html xmlns=`"https://www.w3.org/1999/xhtml`">
	<head>
	<title>hMailServer Backup & Offsite Upload</title>
	<meta name=`"viewport`" content=`"width=device-width, initial-scale=1.0 `" />
	</head>
	<body style=`"font-family:Arial Narrow`">
	<table>
	" | Out-File $EmailBody -Encoding ASCII -Append
}

<#  Authorize and get access token  #>
Debug "Getting access token from LetsUpload"
$URIAuth = "https://letsupload.io/api/v2/authorize"
$AuthBody = @{
	'key1' = $APIKey1;
	'key2' = $APIKey2;
}
Try{
	$Auth = Invoke-RestMethod -Method GET $URIAuth -Body $AuthBody -ContentType 'application/json; charset=utf-8' 
}
Catch {
	Debug "[ERROR] LetsUpload Authentication : $Error"
	Email "[ERROR] LetsUpload Authentication : Check Debug Log"
	EmailResults
	Exit
}
$AccessToken = $Auth.data.access_token
$AccountID = $Auth.data.account_id
Debug "Access Token : $AccessToken"
Debug "Account ID   : $AccountID"

<#  Get folder_id of last upload  #>
$URIFolderListing = "https://letsupload.io/api/v2/folder/listing"
$FLBody = @{
	'access_token' = $AccessToken;
	'account_id' = $AccountID;
	'parent_folder_id' = "";
}
Try{
	$FolderListing = Invoke-RestMethod -Method GET $URIFolderListing -Body $FLBody -ContentType 'application/json; charset=utf-8'
}
Catch {
	Debug "[ERROR] obtaining backup folder ID : $Error"
	Email "[ERROR] obtaining backup folder ID : Check Debug Log"
	EmailResults
	Exit
}
$NewestBackup = $FolderListing.data.folders | Sort-Object date_added -Descending | Where {$_.folderName -match "hMailServer"} | Select -First 1
$FolderID = $NewestBackup.id
$FolderName = $NewestBackup.folderName
Debug "Folder Name: $FolderName"
Debug "Folder ID  : $FolderID"
Email "[OK] Latest backup : $FolderName"

<#  Get file listing within latest backup folder  #>
$URIFolderListing = "https://letsupload.io/api/v2/folder/listing"
$FLBody = @{
	'access_token' = $AccessToken;
	'account_id' = $AccountID;
	'parent_folder_id' = $FolderID;
}
Try{
	$FileListing = Invoke-RestMethod -Method GET $URIFolderListing -Body $FLBody -ContentType 'application/json; charset=utf-8'
}
Catch {
	Debug "[ERROR] obtaining backup file listing : $Error"
	Email "[ERROR] obtaining backup file listing : Check Debug Log"
	EmailResults
	Exit
}
$DownloadCount = ($FileListing.data.files).Count
$DownloadNumber = 1
Debug "File count: $DownloadCount"
Debug "Starting file download"
Email "[OK] $DownloadCount files to download"

<#  Loop through results and download files  #>
$FileListing.data.files | ForEach {
	$FileID = $_.id
	$FileName = $_.filename
	$FileURL = $_.url_file
	$RemoteFileSize = $_.fileSize
	$RemoteFileSizeFormatted = "{0:N2}" -f (($RemoteFileSize)/1MB)
	Debug "----------------------------"
	Debug "File $DownloadNumber of $DownloadCount"
	Debug "File ID     : $FileID"
	Debug "File Name   : $FileName"
	Debug "File Size   : $RemoteFileSizeFormatted mb"

	$URIDownload = "https://letsupload.io/api/v2/file/download"
	$DLBody = @{
		'access_token' = $AccessToken;
		'account_id' = $AccountID;
		'file_id' = $FileID;
	}

	<#  Get download URL  #>
	$GetURLTry = 1
	Do {
		$Error.Clear()
		Try{
			$FileDownload = Invoke-RestMethod -Method GET $URIDownload -Body $DLBody -ContentType 'application/json; charset=utf-8'
			If ($FileDownload._status -match "success") {
				$URLSuccess = $True
			} Else {
				Throw "[ERROR] Getting download URL on Try $GetURLTry"
			}
			$DownloadURL = $FileDownload.data.download_url
			Debug "Download URL: $DownloadURL"
		}
		Catch {
			Debug "[ERROR] Getting download URL : $Error"
			$URLSuccess = $False
		}
		$GetURLTry++
	} Until (($GetURLTry -eq 5) -or ($URLSuccess -eq $True))

	<#  If get download URL success, then download file  #>
	If ($URLSuccess -eq $False) {
		Debug "[ERROR] obtaining download URL : Tried $GetURLTry times : Giving up"
		Debug "[ERROR] obtaining download URL : $Error"
		Email "[ERROR] obtaining download URL : Check Debug Log"
		EmailResults
		Exit
	} Else {
		<#  Download file using BITS  #>
		$DownloadTries = 1
		Do {
			$Error.Clear()
			Try {
				Debug "Download Try $DownloadTries"
				$BeginDL = Get-Date
				Import-Module BitsTransfer
				Start-BitsTransfer -Source $DownloadURL -Destination "$DownloadFolder\$FileName"
				$LocalFileSize = (Get-Item "$DownloadFolder\$FileName").Length
				If ($RemoteFileSize -ne $LocalFileSize) {
					Throw "[ERROR] Remote and local file sizes do not match"
				}
				Debug "File $DownloadNumber downloaded in $(ElapsedTime $BeginDL)"
				$DownloadFileSuccess = $True
			}
			Catch {
				Debug "[ERROR] BITS downloading file $DownloadNumber of $FileCount : $Error"
				Debug "[ERROR] Remote file size: $RemoteFileSize"
				Debug "[ERROR] Local file size : $LocalFileSize"
				$DownloadFileSuccess = $False
			}
			$DownloadTries++
		} Until (($DownloadFileSuccess -eq $True) -or ($DownloadTries -eq 5))
		If ($DownloadFileSuccess -eq $False) {
			Debug "[ERROR] Downloading File : Tried $GetURLTry times : Giving up"
			Debug "[ERROR] Downloading File : $Error"
			Email "[ERROR] Downloading File : Check Debug Log"
			EmailResults
			Exit
		}
	}
	$DownloadNumber++
}

<#  Count and compare remote to local files  #>
If ($DownloadCount -eq ((Get-ChildItem $DownloadFolder).Count)) {
	Debug "----------------------------"
	Debug "Download successful. $DownloadCount files downloaded to $DownloadFolder"
	Email "[OK] Download successful. $DownloadCount files downloaded to $DownloadFolder"
} Else {
	Debug "[ERROR] Download unsuccessful : Remote and local file counts do not match!"
	Email "[ERROR] Download unsuccessful : Remote and local file counts do not match!"
}

<#  Finish up and email results  #>
Debug "----------------------------"
Debug "Script completed in $(ElapsedTime $StartScript)"
Email " "
Email "Script completed in $(ElapsedTime $StartScript)"
Debug "Sending Email"
If (($AttachDebugLog) -and (Test-Path $DebugLog) -and (((Get-Item $DebugLog).length/1MB) -gt $MaxAttachmentSize)){
	Email "Debug log size exceeds maximum attachment size. Please see log file in script folder"
}
If ($UseHTML) {Write-Output "</table></body></html>" | Out-File $EmailBody -Encoding ASCII -Append}
EmailResults