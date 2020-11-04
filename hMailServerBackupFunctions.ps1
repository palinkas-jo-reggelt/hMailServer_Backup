<#

.SYNOPSIS
	hMailServer Backup

.DESCRIPTION
	Common Code for hMailServer Backup

.FUNCTIONALITY
	Backs up hMailServer, compresses backup and uploads to LetsUpload

.PARAMETER 

	
.NOTES
	7-Zip required - install and place in system path
	Run at 12:58PM from task scheduler
	
	
.EXAMPLE


#>

<#  Miscellaneous Functions  #>

Function Debug ($DebugOutput) {
	If ($VerboseFile) {Write-Output "$(Get-Date -f G) : $DebugOutput" | Out-File $DebugLog -Encoding ASCII -Append}
	If ($VerboseConsole) {Write-Host "$(Get-Date -f G) : $DebugOutput"}
}

Function Email ($EmailOutput) {
	If ($UseHTML){
		If ($EmailOutput -match "\[OK\]") {$EmailOutput = $EmailOutput -Replace "\[OK\]","<span style=`"background-color:green;color:white;font-weight:bold;font-family:Courier New;`">[OK]</span>"}
		If ($EmailOutput -match "\[INFO\]") {$EmailOutput = $EmailOutput -Replace "\[INFO\]","<span style=`"background-color:yellow;font-weight:bold;font-family:Courier New;`">[INFO]</span>"}
		If ($EmailOutput -match "\[ERROR\]") {$EmailOutput = $EmailOutput -Replace "\[ERROR\]","<span style=`"background-color:red;color:white;font-weight:bold;font-family:Courier New;`">[ERROR]</span>"}
		If ($EmailOutput -match "^\s$") {$EmailOutput = $EmailOutput -Replace "\s","&nbsp;"}
		Write-Output "<tr><td>$EmailOutput</td></tr>" | Out-File $EmailBody -Encoding ASCII -Append
	} Else {
		Write-Output $EmailOutput | Out-File $EmailBody -Encoding ASCII -Append
	}	
}

Function Plural ($Integer) {
	If ($Integer -eq 1) {$S = ""} Else {$S = "s"}
	Return $S
}

Function EmailResults {
	Try {
		$Body = (Get-Content -Path $EmailBody | Out-String )
		If (($AttachDebugLog) -and (Test-Path $DebugLog) -and (((Get-Item $DebugLog).length/1MB) -lt $MaxAttachmentSize)){$Attachment = New-Object System.Net.Mail.Attachment $DebugLog}
		$Message = New-Object System.Net.Mail.Mailmessage $EmailFrom, $EmailTo, $Subject, $Body
		$Message.IsBodyHTML = $UseHTML
		If (($AttachDebugLog) -and (Test-Path $DebugLog) -and (((Get-Item $DebugLog).length/1MB) -lt $MaxAttachmentSize)){$Message.Attachments.Add($DebugLog)}
		$SMTP = New-Object System.Net.Mail.SMTPClient $SMTPServer,$SMTPPort
		$SMTP.EnableSsl = $SSL
		$SMTP.Credentials = New-Object System.Net.NetworkCredential($SMTPAuthUser, $SMTPAuthPass); 
		$SMTP.Send($Message)
	}
	Catch {
		Debug "Email ERROR : $Error"
	}
}

Function GmailResults ($GBody){
	Try {
		$Subject = "hMailServer Backup Problem"
		$Message = New-Object System.Net.Mail.Mailmessage $GmailUser, $GmailTo, $Subject, $GBody
		$Message.IsBodyHTML = $False
		$SMTP = New-Object System.Net.Mail.SMTPClient("smtp.gmail.com", 587)
		$SMTP.EnableSsl = $True
		$SMTP.Credentials = New-Object System.Net.NetworkCredential($GmailUser, $GmailPass); 
		$SMTP.Send($Message)
	}
	Catch {
		Debug "Gmail ERROR : $Error"
	}
}

Function ValidateFolders ($Folder) {
	If (-not(Test-Path $Folder)) {
		Debug "Error : Folder location $Folder does not exist : Quitting script"
		Email "[ERROR] Folder location $Folder does not exist : Quitting script"
		EmailResults
		Exit
	}
}
 
Function ElapsedTime ($EndTime) {
	$TimeSpan = New-Timespan $EndTime
	If (([int]($TimeSpan).Hours) -eq 0) {$Hours = ""} ElseIf (([int]($TimeSpan).Hours) -eq 1) {$Hours = "1 hour "} Else {$Hours = "$([int]($TimeSpan).Hours) hours "}
	If (([int]($TimeSpan).Minutes) -eq 0) {$Minutes = ""} ElseIf (([int]($TimeSpan).Minutes) -eq 1) {$Minutes = "1 minute "} Else {$Minutes = "$([int]($TimeSpan).Minutes) minutes "}
	If (([int]($TimeSpan).Seconds) -eq 1) {$Seconds = "1 second"} Else {$Seconds = "$([int]($TimeSpan).Seconds) seconds"}
	
	If (($TimeSpan).TotalSeconds -lt 1) {
		$Return = "less than 1 second"
	} Else {
		$Return = "$Hours$Minutes$Seconds"
	}
	Return $Return
}

<#  Service start and stop functions  #>

Function ServiceStart ($ServiceName) {
	<#  Check to see if already running  #>
	If ($ServiceName -eq $hMSServiceName) {$ServiceDescription = "hMailServer"}
	If ($ServiceName -eq $SAServiceName) {$ServiceDescription = "SpamAssassin"}
	$BeginStartupRoutine = Get-Date
	Debug "----------------------------"
	Debug "Start $ServiceDescription"
	$ServiceStopped = $False
	(Get-Service $ServiceName).Refresh()
	If ((Get-Service $ServiceName).Status -eq 'Running'){
		Debug "$ServiceDescription already RUNNING. Nothing to start."
		Email "[ERROR] $ServiceDescription : service already RUNNING. Check event logs."
	} Else {
		Debug "$ServiceDescription not running. Preparing to start service."
		$ServiceStopped = $True
	}

	<#  Start service routine  #>
	If ($ServiceStopped) {
		Debug "$ServiceDescription starting up"
		$BeginStartup = Get-Date
		Do {
			Start-Service $ServiceName
			# Start-Sleep -Seconds 60
			(Get-Service $ServiceName).Refresh()
			$ServiceStatus = (Get-Service $ServiceName).Status
		} Until (((New-Timespan -Start $BeginStartup -End (Get-Date)).TotalMinutes -gt $ServiceTimeout) -or ($ServiceStatus -eq "Running"))

		If ($ServiceStatus -ne "Running"){
			Debug "$ServiceDescription failed to start"
			GmailResults "$ServiceDescription failed to start during backup routine! Check status NOW!"
			Break
		} Else {
			Debug "$ServiceDescription successfully started in $(ElapsedTime $BeginStartupRoutine)"
			Email "[OK] $ServiceDescription started"
		}
	}
}

Function ServiceStop ($ServiceName) {
	<#  Check to see if already stopped  #>
	If ($ServiceName -eq $hMSServiceName) {$ServiceDescription = "hMailServer"}
	If ($ServiceName -eq $SAServiceName) {$ServiceDescription = "SpamAssassin"}
	$BeginShutdownRoutine = Get-Date
	Debug "----------------------------"
	Debug "Stop $ServiceDescription"
	$ServiceRunning = $False
	(Get-Service $ServiceName).Refresh()
	If ((Get-Service $ServiceName).Status -eq 'Stopped'){
		Debug "$ServiceDescription already STOPPED. Nothing to stop. Check event logs."
		Email "[ERROR] $ServiceDescription : service already STOPPED. Check event logs."
	} Else {
		Debug "$ServiceDescription running. Preparing to stop service."
		$ServiceRunning = $True
	}

	<#  Stop service routine  #>
	If ($ServiceRunning) {
		Debug "$ServiceDescription shutting down."
		$BeginShutdown = Get-Date
		Do {
			Stop-Service $ServiceName
			# Start-Sleep -Seconds 60
			(Get-Service $ServiceName).Refresh()
			$ServiceStatus = (Get-Service $ServiceName).Status
		} Until (((New-Timespan $BeginShutdown).TotalMinutes -gt $ServiceTimeout) -or ($ServiceStatus -eq "Stopped"))

		If ($ServiceStatus -ne "Stopped"){
			Debug "$ServiceDescription failed to stop."
			GmailResults "$ServiceDescription failed to stop during backup routine! Check status NOW!"
			Break
		} Else {
			Debug "$ServiceDescription successfully stopped in $(ElapsedTime $BeginShutdownRoutine)"
			Email "[OK] $ServiceDescription stopped"
		}
	}
}

<#  7-zip archive creation function  #>

Function MakeArchive {
	$StartArchive = Get-Date
	Debug "----------------------------"
	Debug "Create archive : $BackupName"
	Debug "Archive folder : $BackupTempDir"
	$VolumeSwitch = "-v$VolumeSize"
	$PWSwitch = "-p$ArchivePassword"
	Try {
		$SevenZip = & cmd /c 7z a $VolumeSwitch -t7z -m0=lzma2 -mx=9 -mfb=64 -md=32m -ms=on -mhe=on $PWSwitch "$BackupLocation\$BackupName\$BackupName.7z" "$BackupTempDir\*" | Out-String
		Debug $SevenZip
		Debug "Archive creation finished in $(ElapsedTime $StartArchive)"
		Debug "Wait a few seconds to make sure archive is finished"
		Email "[OK] 7z archive created"
		Start-Sleep -Seconds 3
	}
	Catch {
		Debug "Archive Creation ERROR : $Error"
		Email "[ERROR] Archive Creation : Check Debug Log"
		Email "[ERROR] Archive Creation : $Error"
		EmailResults
		Exit
	}
}

<#  Prune Messages Functions  #> 

Set-Variable -Name TotalDeletedMessages -Value 0 -Option AllScope
Set-Variable -Name TotalDeletedFolders -Value 0 -Option AllScope
Set-Variable -Name DeleteMessageErrors -Value 0 -Option AllScope
Set-Variable -Name DeleteFolderErrors -Value 0 -Option AllScope

Function GetSubFolders ($Folder) {
	$IterateFolder = 0
	$ArrayDeletedFolders = @()
	If ($Folder.SubFolders.Count -gt 0) {
		Do {
			$SubFolder = $Folder.SubFolders.Item($IterateFolder)
			$SubFolderName = $SubFolder.Name
			$SubFolderID = $SubFolder.ID
			If ($SubFolder.Subfolders.Count -gt 0) {GetSubFolders $SubFolder} 
			If ($SubFolder.Messages.Count -gt 0) {
				If ($PruneSubFolders) {GetMessages $SubFolder}
			} Else {
				If ($DeleteEmptySubFolders) {$ArrayDeletedFolders += $SubFolderID}
			} 
			$IterateFolder++
		} Until ($IterateFolder -eq $Folder.SubFolders.Count)
	}
	If ($DeleteEmptySubFolders) {
		$ArrayDeletedFolders | ForEach {
			$CheckFolder = $Folder.SubFolders.ItemByDBID($_)
			$FolderName = $CheckFolder.Name
			If (SubFoldersEmpty $CheckFolder) {
				Try {
					If ($DoDelete) {$Folder.SubFolders.DeleteByDBID($_)}
					$TotalDeletedFolders++
					Debug "Deleted empty subfolder $FolderName in $AccountAddress"
				}
				Catch {
					$DeleteFolderErrors++
					Debug "[ERROR] Deleting empty subfolder $FolderName in $AccountAddress"
					Debug "[ERROR] : $Error"
				}
				$Error.Clear()
			}
		}
	}
	$ArrayDeletedFolders.Clear()
}

Function SubFoldersEmpty ($Folder) {
	$IterateFolder = 0
	If ($Folder.SubFolders.Count -gt 0) {
		Do {
			$SubFolder = $Folder.SubFolders.Item($IterateFolder)
			If ($SubFolder.Messages.Count -gt 0) {
				Return $False
				Break
			}
			If ($SubFolder.SubFolders.Count -gt 0) {
				SubFoldersEmpty $SubFolder
			}
			$IterateFolder++
		} Until ($IterateFolder -eq $Folder.SubFolders.Count)
	} Else {
		Return $True
	}
}

Function GetMatchFolders ($Folder) {
	$IterateFolder = 0
	If ($Folder.SubFolders.Count -gt 0) {
		Do {
			$SubFolder = $Folder.SubFolders.Item($IterateFolder)
			$SubFolderName = $SubFolder.Name
			If ($SubFolderName -match [regex]$PruneFolders) {
				GetSubFolders $SubFolder
				GetMessages $SubFolder
			} Else {
				GetMatchFolders $SubFolder
			}
			$IterateFolder++
		} Until ($IterateFolder -eq $Folder.SubFolders.Count)
	}
}

Function GetMessages ($Folder) {
	$IterateMessage = 0
	$ArrayMessagesToDelete = @()
	$DeletedMessages = 0
	If ($Folder.Messages.Count -gt 0) {
		Do {
			$Message = $Folder.Messages.Item($IterateMessage)
			If ($Message.InternalDate -lt ((Get-Date).AddDays(-$DaysBeforeDelete))) {
				$ArrayMessagesToDelete += $Message.ID
			}
			$IterateMessage++
		} Until ($IterateMessage -eq $Folder.Messages.Count)
	}
	$ArrayMessagesToDelete | ForEach {
		$AFolderName = $Folder.Name
		Try {
			If ($DoDelete) {$Folder.Messages.DeleteByDBID($_)}
			$DeletedMessages++
			$TotalDeletedMessages++
		}
		Catch {
			$DeleteMessageErrors++
			Debug "[ERROR] Deleting messages from folder $AFolderName in $AccountAddress"
			Debug "[ERROR] $Error"
		}
		$Error.Clear()
	}
	If ($DeletedMessages -gt 0) {
		Debug "Deleted $DeletedMessages message$(Plural $DeletedMessages) from $AFolderName in $AccountAddress"
	}
	$ArrayMessagesToDelete.Clear()
}

Function PruneMessages {
	
	$Error.Clear()
	$BeginDeletingOldMessages = Get-Date
	Debug "----------------------------"
	Debug "Begin deleting messages older than $DaysBeforeDelete days"
	If (-not($DoDelete)) {
		Debug "Delete disabled - Test Run ONLY"
	}

	<#  Authenticate hMailServer COM  #>
	$hMS = New-Object -COMObject hMailServer.Application
	$hMS.Authenticate("Administrator", $hMSAdminPass) | Out-Null
	
	$IterateDomains = 0
	Do {
		$hMSDomain = $hMS.Domains.Item($IterateDomains)
		If ($hMSDomain.Active) {
			$IterateAccounts = 0
			Do {
				$hMSAccount = $hMSDomain.Accounts.Item($IterateAccounts)
				If ($hMSAccount.Active) {
					$AccountAddress = $hMSAccount.Address
					$IterateIMAPFolders = 0
					If ($hMSAccount.IMAPFolders.Count -gt 0) {
						Do {
							$hMSIMAPFolder = $hMSAccount.IMAPFolders.Item($IterateIMAPFolders)
							If ($hMSIMAPFolder.Name -match [regex]$PruneFolders) {
								If ($hMSIMAPFolder.SubFolders.Count -gt 0) {
									GetSubFolders $hMSIMAPFolder
								} # IF SUBFOLDER COUNT > 0
								GetMessages $hMSIMAPFolder
							} # IF FOLDERNAME MATCH REGEX
							Else {
								GetMatchFolders $hMSIMAPFolder
							} # IF NOT FOLDERNAME MATCH REGEX
						$IterateIMAPFolders++
						} Until ($IterateIMAPFolders -eq $hMSAccount.IMAPFolders.Count)
					} # IF IMAPFOLDER COUNT > 0
				} #IF ACCOUNT ACTIVE
				$IterateAccounts++
			} Until ($IterateAccounts -eq $hMSDomain.Accounts.Count)
		} # IF DOMAIN ACTIVE
		$IterateDomains++
	} Until ($IterateDomains -eq $hMS.Domains.Count)

	If ($DeleteMessageErrors -gt 0) {
		Debug "Finished Message Pruning : $DeleteMessageErrors Errors present"
		Email "[ERROR] Message Pruning : $DeleteMessageErrors Errors present : Check debug log"
	} Else {
		If ($TotalDeletedMessages -gt 0) {
			Debug "Finished pruning $TotalDeletedMessages message$(Plural $TotalDeletedMessages) in $(ElapsedTime $BeginDeletingOldMessages)"
			Email "[OK] Pruned $TotalDeletedMessages message$(Plural $TotalDeletedMessages)"
		} Else {
			Debug "No messages older than $DaysBeforeDelete days to prune"
			Email "[OK] No messages older than $DaysBeforeDelete days to prune"
		}
	}
	If ($DeleteFolderErrors -gt 0) {
		Debug "Deleting Empty Folders : $DeleteFolderErrors Error$(Plural $DeleteFolderErrors) present"
		Email "[ERROR] Deleting Empty Folders : $DeleteFolderErrors Error$(Plural $DeleteFolderErrors) present : Check debug log"
	} Else {
		If ($TotalDeletedFolders -gt 0) {
			Debug "Finished deleting $TotalDeletedFolders empty subfolder$(Plural $TotalDeletedFolders)"
			Email "[OK] Deleted $TotalDeletedFolders empty subfolder$(Plural $TotalDeletedFolders)"
		} Else {
			Debug "No empty subfolders deleted"
			Email "[OK] No empty subfolders to delete"
		}
	}
}

<#  Offsite upload function  #>

Function OffsiteUpload {

	$BeginOffsiteUpload = Get-Date
	Debug "----------------------------"
	Debug "Begin offsite upload process"

	<#  Authorize and get access token  #>
	Debug "Getting access token from LetsUpload"
	$URIAuth = "https://letsupload.io/api/v2/authorize"
	$AuthBody = @{
		'key1' = $APIKey1;
		'key2' = $APIKey2;
	}
	Try{
		$Auth = Invoke-RestMethod -Method GET $URIAuth -Body $AuthBody -ContentType 'application/json; charset=utf-8' 
		$AccessToken = $Auth.data.access_token
		$AccountID = $Auth.data.account_id
		Debug "Access Token : $AccessToken"
		Debug "Account ID   : $AccountID"
	}
	Catch {
		Debug "LetsUpload Authentication ERROR : $Error"
		Email "[ERROR] LetsUpload Authentication : Check Debug Log"
		Email "[ERROR] LetsUpload Authentication : $Error"
		EmailResults
		Exit
	}

	<#  Create Folder  #>
	Debug "----------------------------"
	Debug "Creating Folder $BackupName at LetsUpload"
	$URICF = "https://letsupload.io/api/v2/folder/create"
	$CFBody = @{
		'access_token' = $AccessToken;
		'account_id' = $AccountID;
		'folder_name' = $BackupName;
		'is_public' = $IsPublic;
	}
	Try {
		$CreateFolder = Invoke-RestMethod -Method GET $URICF -Body $CFBody -ContentType 'application/json; charset=utf-8' 
		$CFResponse = $CreateFolder.response
		$FolderID = $CreateFolder.data.id
		$FolderURL = $CreateFolder.data.url_folder
		Debug "Response   : $CFResponse"
		Debug "Folder ID  : $FolderID"
		Debug "Folder URL : $FolderURL"
	}
	Catch {
		Debug "LetsUpload Folder Creation ERROR : $Error"
		Email "[ERROR] LetsUpload Folder Creation : Check Debug Log"
		Email "[ERROR] LetsUpload Folder Creation : $Error"
		EmailResults
		Exit
	}

	<#  Upload Files  #>
	$StartUpload = Get-Date
	Debug "----------------------------"
	Debug "Begin uploading files to LetsUpload"
	$CountArchVol = (Get-ChildItem "$BackupLocation\$BackupName").Count
	Debug "There are $CountArchVol files to upload"
	$UploadCounter = 1

	Get-ChildItem "$BackupLocation\$BackupName" | ForEach {

		$FileName = $_.Name;
		$FilePath = $_.FullName;
		$FileSize = $_.Length;
		
		$UploadURI = "https://letsupload.io/api/v2/file/upload";
		Debug "----------------------------"
		Debug "Encoding file $FileName"
		$BeginEnc = Get-Date
		Try {
			$FileBytes = [System.IO.File]::ReadAllBytes($FilePath);
			$FileEnc = [System.Text.Encoding]::GetEncoding('ISO-8859-1').GetString($FileBytes);
		}
		Catch {
				Debug "Error in encoding file $UploadCounter."
				Debug "$Error"
				Debug " "
		}
		Debug "Finished encoding file in $(ElapsedTime $BeginEnc)";
		$Boundary = [System.Guid]::NewGuid().ToString(); 
		$LF = "`r`n";

		$BodyLines = (
			"--$Boundary",
			"Content-Disposition: form-data; name=`"access_token`"",
			'',
			$AccessToken,
			"--$Boundary",
			"Content-Disposition: form-data; name=`"account_id`"",
			'',
			$AccountID,
			"--$Boundary",
			"Content-Disposition: form-data; name=`"folder_id`"",
			'',
			$FolderID,
			"--$Boundary",
			"Content-Disposition: form-data; name=`"upload_file`"; filename=`"$FileName`"",
			"Content-Type: application/json",
			'',
			$FileEnc,
			"--$Boundary--"
		) -join $LF
			
		Debug "Uploading $FileName - $UploadCounter of $CountArchVol"
		$UploadTries = 1
		$BeginUpload = Get-Date
		Do {
			$Error.Clear()
			$Upload = $UResponse = $UURL = $USize = $UStatus = $NULL
			Try {
				$Upload = Invoke-RestMethod -Uri $UploadURI -Method POST -ContentType "multipart/form-data; boundary=`"$boundary`"" -Body $BodyLines
				$UResponse = $Upload.response
				$UURL = $Upload.data.url
				$USize = $Upload.data.size
				$USizeFormatted = "{0:N2}" -f (($USize)/1MB)
				$UStatus = $Upload._status
				$UFileID = $upload.data.file_id
				If ($USize -ne $FileSize) {Throw "Local and remote filesizes do not match!"}
				Debug "Upload try $UploadTries"
				Debug "Response : $UResponse"
				Debug "File ID  : $UFileID"
				Debug "URL      : $UURL"
				Debug "Size     : $USizeFormatted mb"
				Debug "Status   : $UStatus"
				Debug "Finished uploading file in $(ElapsedTime $BeginUpload)"
			} 
			Catch {
				Debug "Upload try $UploadTries"
				Debug "[ERROR]  : $Error"
				If (($USize -gt 0) -and ($UFileID -match '\d+')) {
					Debug "Deleting file due to size mismatch"
					$URIDF = "https://letsupload.io/api/v2/file/delete"
					$DFBody = @{
						'access_token' = $AccessToken;
						'account_id' = $AccountID;
						'file_id' = $UFileID;
					}
					Try {
						$DeleteFile = Invoke-RestMethod -Method GET $URIDF -Body $DFBody -ContentType 'application/json; charset=utf-8' 
					}
					Catch {
						Debug "File delete ERROR : $Error"
					}
				}
			}
			$UploadTries++
		} Until (($UploadTries -eq ($MaxUploadTries + 1)) -or ($UStatus -match "success"))

		If (-not($UStatus -Match "success")) {
			Debug "Error in uploading file number $UploadCounter. Check the log for errors."
			Email "[ERROR] in uploading file number $UploadCounter. Check the log for errors."
			EmailResults
			Exit
		}
		$UploadCounter++
	}
	Debug "----------------------------"
	Debug "Finished offsite upload process in $(ElapsedTime $BeginOffsiteUpload)"
	
	<#  Count remote files  #>
	Debug "----------------------------"
	Debug "Counting uploaded files at LetsUpload"
	$URIFL = "https://letsupload.io/api/v2/folder/listing"
	$FLBody = @{
		'access_token' = $AccessToken;
		'account_id' = $AccountID;
		'parent_folder_id' = $FolderID;
	}
	Try {
		$FolderListing = Invoke-RestMethod -Method GET $URIFL -Body $FLBody -ContentType 'application/json; charset=utf-8' 
	}
	Catch {
		Debug "LetsUpload Folder Listing ERROR : $Error"
		Email "[ERROR] LetsUpload Folder Listing : Check Debug Log"
		Email "[ERROR] LetsUpload Folder Listing : $Error"
	}
	$FolderListingStatus = $FolderListing._status
	$RemoteFileCount = ($FolderListing.data.files.id).Count
	
	<#  Report results  #>
	If ($FolderListingStatus -match "success") {
		Debug "There are $RemoteFileCount file$(Plural $RemoteFileCount) in the remote folder"
		If ($RemoteFileCount -eq $CountArchVol) {
			Debug "----------------------------"
			Debug "Finished uploading $CountArchVol file$(Plural $CountArchVol) in $(ElapsedTime $StartUpload)"
			Debug "Upload sucessful. $CountArchVol file$(Plural $CountArchVol) uploaded to $FolderURL"
			Email "[OK] Offsite backup upload:"
			Email "[OK] $CountArchVol file$(Plural $CountArchVol) uploaded to $FolderURL"
		} Else {
			Debug "----------------------------"
			Debug "Finished uploading in $(ElapsedTime $StartUpload)"
			Debug "[ERROR] Number of archive files uploaded does not match count in remote folder"
			Debug "[ERROR] Archive volumes   : $CountArchVol"
			Debug "[ERROR] Remote file count : $RemoteFileCount"
			Email "[ERROR] Number of archive files uploaded does not match count in remote folder - see debug log"
		}
	} Else {
		Debug "----------------------------"
		Debug "Error : Unable to obtain file count from remote folder"
		Email "[ERROR] Unable to obtain uploaded file count from remote folder - see debug log"
	}

}