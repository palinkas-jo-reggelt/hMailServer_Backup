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
		Debug "Email ERROR : $($Error[0])"
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
		Debug "Gmail ERROR : $($Error[0])"
	}
}

Function ValidateFolders ($Folder) {
	If (-not(Test-Path $Folder)) {
		Debug "[ERROR] Folder location $Folder does not exist : Quitting script"
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

Function ServiceStop ($ServiceName) {
	<#  Check to see if already stopped  #>
	$BeginShutdownRoutine = Get-Date
	Debug "----------------------------"
	Debug "Stop $ServiceName"
	$ServiceRunning = $False
	(Get-Service $ServiceName).Refresh()
	If ((Get-Service $ServiceName).Status -eq 'Stopped'){
		Debug "$ServiceName already STOPPED. Nothing to stop. Check event logs."
		Email "[INFO] $ServiceName : service already STOPPED. Check event logs."
	} Else {
		Debug "$ServiceName running. Preparing to stop service."
		$ServiceRunning = $True
	}

	<#  Stop service routine  #>
	If ($ServiceRunning) {
		Debug "$ServiceName shutting down."
		$BeginShutdown = Get-Date
		Do {
			Stop-Service $ServiceName
			# Start-Sleep -Seconds 60
			(Get-Service $ServiceName).Refresh()
			$ServiceStatus = (Get-Service $ServiceName).Status
		} Until (((New-Timespan $BeginShutdown).TotalMinutes -gt $ServiceTimeout) -or ($ServiceStatus -eq "Stopped"))

		If ($ServiceStatus -ne "Stopped"){
			Debug "$ServiceName failed to stop."
			GmailResults "$ServiceName failed to stop during backup routine! Check status NOW!"
			Break
		} Else {
			Debug "$ServiceName successfully stopped in $(ElapsedTime $BeginShutdownRoutine)"
			Email "[OK] $ServiceName stopped"
		}
	}
}

Function ServiceStart ($ServiceName) {
	<#  Check to see if already running  #>
	$BeginStartupRoutine = Get-Date
	Debug "----------------------------"
	Debug "Start $ServiceName"
	$ServiceStopped = $False
	(Get-Service $ServiceName).Refresh()
	If ((Get-Service $ServiceName).Status -eq 'Running'){
		Debug "$ServiceName already RUNNING. Nothing to start."
		Email "[INFO] $ServiceName : service already RUNNING. Check event logs."
	} Else {
		Debug "$ServiceName not running. Preparing to start service."
		$ServiceStopped = $True
	}

	<#  Start service routine  #>
	If ($ServiceStopped) {
		Debug "$ServiceName starting up"
		$BeginStartup = Get-Date
		Do {
			Start-Service $ServiceName
			# Start-Sleep -Seconds 60
			(Get-Service $ServiceName).Refresh()
			$ServiceStatus = (Get-Service $ServiceName).Status
		} Until (((New-Timespan $BeginStartup).TotalMinutes -gt $ServiceTimeout) -or ($ServiceStatus -eq "Running"))

		If ($ServiceStatus -ne "Running"){
			Debug "$ServiceName failed to start"
			GmailResults "$ServiceName failed to start during backup routine! Check status NOW!"
			Break
		} Else {
			Debug "$ServiceName successfully started in $(ElapsedTime $BeginStartupRoutine)"
			Email "[OK] $ServiceName started"
		}
	}
}

<#  7-zip archive creation function  #>

Function MakeArchive {
	$StartArchive = Get-Date
	Debug "----------------------------"
	Debug "Create archive : $BackupName"
	Debug "Archive folder : $BackupTempDir"
	$SevenZipExe = "$hMSDir\Bin\7za.exe"
	$VolumeSwitch = "-v$VolumeSize"
	$PWSwitch = "-p$ArchivePassword"
	Try {
		If ($PWProtectedArchive) {
			$SevenZip = & cmd /c $SevenZipExe a $VolumeSwitch -t7z -m0=lzma2 -mx=9 -mfb=64 -md=32m -ms=on -mhe=on $PWSwitch "$BackupLocation\$BackupName\$BackupName.7z" "$BackupTempDir\*" | Out-String
		} Else {
			$SevenZip = & cmd /c $SevenZipExe a -tzip "$BackupLocation\$BackupName\$BackupName.zip" "$BackupTempDir\*" | Out-String
		}
		Debug $SevenZip
		Debug "Archive creation finished in $(ElapsedTime $StartArchive)"
		Debug "Wait a few seconds to make sure archive is finished"
		Email "[OK] 7z archive created"
		Start-Sleep -Seconds 3
	}
	Catch {
		Debug "[ERROR] Archive Creation : $($Error[0])"
		Email "[ERROR] Archive Creation : Check Debug Log"
		Email "[ERROR] Archive Creation : $($Error[0])"
		EmailResults
		Exit
	}
}


<#  Cycle Logs  #>

Function CycleLogs {
	Debug "----------------------------"
	Debug "Cycling Logs"
	$LogsToCycle | ForEach {
		$FullName = (Get-Item $_).FullName
		$BaseName = (Get-Item $_).BaseName
		$FileExt = (Get-Item $_).Extension
		If (Test-Path $FullName) {
			$NewLogName = $BaseName+"_"+$((Get-Date).ToString('yyyy-MM-dd'))+$FileExt
			Try {
				Rename-Item $FullName $NewLogName -ErrorAction Stop
				Debug "Cylcled $NewLogName"
				} 
			Catch {
				Debug "[ERROR] Log cylcling ERROR : $($Error[0])"
			}
		} Else {
			Debug "[ERROR] $FullName not found"
		}
	}
}

<#  Prune hMailServer logs  #>

Function PruneLogs {
	$FilesToDel = Get-ChildItem -Path "$hMSDir\Logs" | Where-Object {$_.LastWriteTime -lt ((Get-Date).AddDays(-$DaysToKeepLogs))}
	$CountDelLogs = $FilesToDel.Count
	If ($CountDelLogs -gt 0){
		Debug "----------------------------"
		Debug "Begin pruning hMailServer logs older than $DaysToKeepLogs days"
		$EnumCountDelLogs = 0
		Try {
			$FilesToDel | ForEach {
				$FullName = $_.FullName
				$Name = $_.Name
				If (Test-Path $_.FullName -PathType Leaf) {
					Remove-Item -Force -Path $FullName
					Debug "Deleted file  : $Name"
					$EnumCountDelLogs++
				}
			}
			If ($CountDelLogs -eq $EnumCountDelLogs) {
				Debug "Successfully pruned $CountDelLogs hMailServer log$(Plural $CountDelLogs)"
				Email "[OK] Pruned hMailServer logs older than $DaysToKeepLogs days"
			} Else {
				Debug "[ERROR] Prune hMailServer logs : Filecount does not match delete count"
				Email "[ERROR] Prune hMailServer logs : Check Debug Log"
			}
		}
		Catch {
			Debug "[ERROR] Prune hMailServer logs : $($Error[0])"
			Email "[ERROR] Prune hMailServer logs : Check Debug Log"
		}
	}
}

<#  Prune Backups Function  #>

Function PruneBackups {
	$FilesToDel = Get-ChildItem -Path $BackupLocation  | Where-Object {$_.LastWriteTime -lt ((Get-Date).AddDays(-$DaysToKeepBackups))}
	$CountDel = $FilesToDel.Count
	If ($CountDel -gt 0) {
		Debug "----------------------------"
		Debug "Begin pruning local backups older than $DaysToKeepBackups days"
		$EnumCountDel = 0
		Try {
			$FilesToDel | ForEach {
				$FullName = $_.FullName
				$Name = $_.Name
				If (Test-Path $_.FullName -PathType Container) {
					Remove-Item -Force -Recurse -Path $FullName
					Debug "Deleted folder: $Name"
					$EnumCountDel++
				}
				If (Test-Path $_.FullName -PathType Leaf) {
					Remove-Item -Force -Path $FullName
					Debug "Deleted file  : $Name"
					$EnumCountDel++
				}
			}
			If ($CountDel -eq $EnumCountDel) {
				Debug "Successfully pruned $CountDel item$(Plural $CountDel)"
				Email "[OK] Pruned backups older than $DaysToKeepBackups days"
			} Else {
				Debug "[ERROR] Prune backups : Filecount does not match delete count"
				Email "[ERROR] Prune backups : Check Debug Log"
			}
		}
		Catch {
			Debug "[ERROR] Prune backups : $($Error[0])"
			Email "[ERROR] Prune backups : Check Debug Log"
		}
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
				If ($PruneEmptySubFolders) {$ArrayDeletedFolders += $SubFolderID}
			} 
			$IterateFolder++
		} Until ($IterateFolder -eq $Folder.SubFolders.Count)
	}
	If ($PruneEmptySubFolders) {
		$ArrayDeletedFolders | ForEach {
			$CheckFolder = $Folder.SubFolders.ItemByDBID($_)
			If (SubFoldersEmpty $CheckFolder) {
				Try {
					If ($DoDelete) {$Folder.SubFolders.DeleteByDBID($_)}
					$TotalDeletedFolders++
					Debug "Deleted empty subfolder $($CheckFolder.Name) in $($hMSAccount.Address)"
				}
				Catch {
					$DeleteFolderErrors++
					Debug "[ERROR] Deleting empty subfolder $($CheckFolder.Name) in $($hMSAccount.Address)"
					Debug "[ERROR] : $($Error[0])"
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
			If ($SubFolderName -match $PruneFolders) {
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
		Try {
			If ($DoDelete) {$Folder.Messages.DeleteByDBID($_)}
			$DeletedMessages++
			$TotalDeletedMessages++
		}
		Catch {
			$DeleteMessageErrors++
			Debug "[ERROR] Deleting messages from folder $($Folder.Name) in $($hMSAccount.Address)"
			Debug "[ERROR] $($Error[0])"
		}
		$Error.Clear()
	}
	If ($DeletedMessages -gt 0) {
		Debug "Deleted $DeletedMessages message$(Plural $DeletedMessages) from $($Folder.Name) in $($hMSAccount.Address)"
	}
	$ArrayMessagesToDelete.Clear()
}

Function PruneMessages {
	
	$Error.Clear()
	$BeginDeletingOldMessages = Get-Date
	Debug "----------------------------"
	Debug "Begin pruning messages older than $DaysBeforeDelete days"
	If (-not($DoDelete)) {
		Debug "Delete disabled - Test Run ONLY"
	}

	<#  Authenticate hMailServer COM  #>
	$hMS = New-Object -COMObject hMailServer.Application
	$hMS.Authenticate("Administrator", $hMSAdminPass) | Out-Null
	
	If ($hMS.Domains.Count -gt 0) {
		$IterateDomains = 0
		Do {
			$hMSDomain = $hMS.Domains.Item($IterateDomains)
			If (($hMSDomain.Active) -and ($hMSDomain.Name -notmatch [regex]$SkipDomainPruning) -and ($hMSDomain.Accounts.Count -gt 0)) {
				$IterateAccounts = 0
				Do {
					$hMSAccount = $hMSDomain.Accounts.Item($IterateAccounts)
					If (($hMSAccount.Active) -and ($hMSAccount.Address -notmatch [regex]$SkipAccountPruning) -and ($hMSAccount.IMAPFolders.Count -gt 0)) {
						$IterateIMAPFolders = 0
						Do {
							$hMSIMAPFolder = $hMSAccount.IMAPFolders.Item($IterateIMAPFolders)
							If ($hMSIMAPFolder.Name -match $PruneFolders) {
								If ($hMSIMAPFolder.SubFolders.Count -gt 0) {
									GetSubFolders $hMSIMAPFolder
								}
								GetMessages $hMSIMAPFolder
							} Else {
								GetMatchFolders $hMSIMAPFolder
							}
						$IterateIMAPFolders++
						} Until ($IterateIMAPFolders -eq $hMSAccount.IMAPFolders.Count)
					}
					$IterateAccounts++
				} Until ($IterateAccounts -eq $hMSDomain.Accounts.Count)
			}
			$IterateDomains++
		} Until ($IterateDomains -eq $hMS.Domains.Count)
	}

	<#  Report message pruning  #>
	If ($DeleteMessageErrors -gt 0) {
		Debug "Finished Message Pruning : $DeleteMessageErrors Errors present"
		Email "[ERROR] Message Pruning : $DeleteMessageErrors Errors present : Check debug log"
	}
	If ($TotalDeletedMessages -gt 0) {
		Debug "Finished pruning $TotalDeletedMessages messages in $(ElapsedTime $BeginDeletingOldMessages)"
		Email "[OK] Pruned $TotalDeletedMessages messages older than $DaysBeforeDelete days"
	} Else {
		Debug "No messages older than $DaysBeforeDelete days to prune"
		Email "[OK] No messages older than $DaysBeforeDelete days to prune"
	}

	<#  Report folder pruning  #>
	If ($DeleteFolderErrors -gt 0) {
		Debug "Deleting Empty Folders : $DeleteFolderErrors Errors present"
		Email "[ERROR] Deleting Empty Folders : $DeleteFolderErrors Errors present : Check debug log"
	}
	If ($TotalDeletedFolders -gt 0) {
		Debug "Deleted $TotalDeletedFolders empty subfolders"
		Email "[OK] Deleted $TotalDeletedFolders empty subfolders"
	} Else {
		Debug "No empty subfolders deleted"
		# Email "[OK] No empty subfolders deleted"
	}
}

<#  Feed Bayes  #>

Set-Variable -Name TotalHamFedMessages -Value 0 -Option AllScope
Set-Variable -Name TotalSpamFedMessages -Value 0 -Option AllScope
Set-Variable -Name HamFedMessageErrors -Value 0 -Option AllScope
Set-Variable -Name SpamFedMessageErrors -Value 0 -Option AllScope
Set-Variable -Name LearnedHamMessages -Value 0 -Option AllScope
Set-Variable -Name LearnedSpamMessages -Value 0 -Option AllScope

Function GetBayesSubFolders ($Folder) {
	$IterateFolder = 0
	$ArrayBayesMessages = @()
	If ($Folder.SubFolders.Count -gt 0) {
		Do {
			$SubFolder = $Folder.SubFolders.Item($IterateFolder)
			$SubFolderName = $SubFolder.Name
			$SubFolderID = $SubFolder.ID
			If ($SubFolder.Subfolders.Count -gt 0) {GetBayesSubFolders $SubFolder} 
			If ($SubFolder.Messages.Count -gt 0) {
				If ($BayesSubFolders) {GetBayesMessages $SubFolder}
			} 
			$IterateFolder++
		} Until ($IterateFolder -eq $Folder.SubFolders.Count)
	}
	$ArrayBayesMessages.Clear()
}

Function GetBayesMatchFolders ($Folder) {
	$IterateFolder = 0
	If ($Folder.SubFolders.Count -gt 0) {
		Do {
			$SubFolder = $Folder.SubFolders.Item($IterateFolder)
			$SubFolderName = $SubFolder.Name
			If (($SubFolderName -match $HamFolders) -or ($SubFolderName -match $SpamFolders)) {
				GetBayesSubFolders $SubFolder
				GetBayesMessages $SubFolder
			} Else {
				GetBayesMatchFolders $SubFolder
			}
			$IterateFolder++
		} Until ($IterateFolder -eq $Folder.SubFolders.Count)
	}
}

Function GetBayesMessages ($Folder) {
	$IterateMessage = 0
	$ArrayHamToFeed = @()
	$ArraySpamToFeed = @()
	$HamFedMessages = 0
	$SpamFedMessages = 0
	$LearnedHamMessagesFolder = 0
	$LearnedSpamMessagesFolder = 0
	If ($Folder.Messages.Count -gt 0) {
		If ($Folder.Name -match $HamFolders) {
			Do {
				$Message = $Folder.Messages.Item($IterateMessage)
				If ($Message.InternalDate -gt ((Get-Date).AddDays(-$BayesDays))) {
					$ArrayHamToFeed += $Message.FileName
				}
				$IterateMessage++
			} Until ($IterateMessage -eq $Folder.Messages.Count)
		}
		If ($Folder.Name -match $SpamFolders) {
			Do {
				$Message = $Folder.Messages.Item($IterateMessage)
				If ($Message.InternalDate -gt ((Get-Date).AddDays(-$BayesDays))) {
					$ArraySpamToFeed += $Message.FileName
				}
				$IterateMessage++
			} Until ($IterateMessage -eq $Folder.Messages.Count)
		}
	}
	$ArrayHamToFeed | ForEach {
		$FileName = $_
		Try {
			If ((Get-Item $FileName).Length -lt 512000) {
				If ($DoSpamC) {
					$SpamC = & cmd /c "`"$SADir\spamc.exe`" -d `"$SAHost`" -p `"$SAPort`" -L ham < `"$FileName`""
					$SpamCResult = Out-String -InputObject $SpamC
					If ($SpamCResult -match "Message successfully un/learned") {
						$LearnedHamMessages++
						$LearnedHamMessagesFolder++
					}
					If (($SpamCResult -notmatch "Message successfully un/learned") -and ($SpamCResult -notmatch "Message was already un/learned")) {
						Throw $SpamCResult
					}
				}
				$HamFedMessages++
				$TotalHamFedMessages++
			}
		}
		Catch {
			$HamFedMessageErrors++
			Debug "[ERROR] Feeding HAM message $FileName in $($hMSAccount.Address)"
			Debug "[ERROR] $($Error[0])"
		}
	}
	$ArraySpamToFeed | ForEach {
		$FileName = $_
		Try {
			If ((Get-Item $FileName).Length -lt 512000) {
				If ($DoSpamC) {
					$SpamC = & cmd /c "`"$SADir\spamc.exe`" -d `"$SAHost`" -p `"$SAPort`" -L spam < `"$FileName`""
					$SpamCResult = Out-String -InputObject $SpamC
					If ($SpamCResult -match "Message successfully un/learned") {
						$LearnedSpamMessages++
						$LearnedSpamMessagesFolder++
					}
					If (($SpamCResult -notmatch "Message successfully un/learned") -and ($SpamCResult -notmatch "Message was already un/learned")) {
						Throw $SpamCResult
					}
				}
				$SpamFedMessages++
				$TotalSpamFedMessages++
			}
		}
		Catch {
			$SpamFedMessageErrors++
			Debug "[ERROR] Feeding SPAM message $FileName in $($hMSAccount.Address)"
			Debug "[ERROR] $($Error[0])"
		}
	}
	If ($HamFedMessages -gt 0) {
		Debug "Learned tokens from $LearnedHamMessagesFolder of $HamFedMessages HAM message$(Plural $HamFedMessages) fed from $($Folder.Name) in $($hMSAccount.Address)"
	}
	If ($SpamFedMessages -gt 0) {
		Debug "Learned tokens from $LearnedSpamMessagesFolder of $SpamFedMessages SPAM message$(Plural $SpamFedMessages) fed from $($Folder.Name) in $($hMSAccount.Address)"
	}
	$ArraySpamToFeed.Clear()
}

Function FeedBayes {
	
	$Error.Clear()
	
	$BeginFeedingBayes = Get-Date
	Debug "----------------------------"
	Debug "Begin learning Bayes tokens from messages newer than $BayesDays days"
	If (-not($DoSpamC)) {
		Debug "SpamC disabled - Test Run ONLY"
	}

	<#  Authenticate hMailServer COM  #>
	$hMS = New-Object -COMObject hMailServer.Application
	$hMS.Authenticate("Administrator", $hMSAdminPass) | Out-Null
	
	$SAHost = $hMS.Settings.AntiSpam.SpamAssassinHost
	$SAPort = $hMS.Settings.AntiSpam.SpamAssassinPort
	
	If ($hMS.Domains.Count -gt 0) {
		$IterateDomains = 0
		Do {
			$hMSDomain = $hMS.Domains.Item($IterateDomains)
			If (($hMSDomain.Active) -and ($hMSDomain.Name -notmatch [regex]$SkipDomainBayes) -and ($hMSDomain.Accounts.Count -gt 0)) {
				$IterateAccounts = 0
				Do {
					$hMSAccount = $hMSDomain.Accounts.Item($IterateAccounts)
					If (($hMSAccount.Active) -and ($hMSAccount.Address -notmatch [regex]$SkipAccountBayes) -and ($hMSAccount.IMAPFolders.Count -gt 0)) {
						$IterateIMAPFolders = 0
						Do {
							$hMSIMAPFolder = $hMSAccount.IMAPFolders.Item($IterateIMAPFolders)
							If (($hMSIMAPFolder.Name -match $HamFolders) -or ($hMSIMAPFolder.Name -match $SpamFolders)) {
								If ($hMSIMAPFolder.SubFolders.Count -gt 0) {
									GetBayesSubFolders $hMSIMAPFolder
								}
								GetBayesMessages $hMSIMAPFolder
							} Else {
								GetBayesMatchFolders $hMSIMAPFolder
							}
						$IterateIMAPFolders++
						} Until ($IterateIMAPFolders -eq $hMSAccount.IMAPFolders.Count)
					}
					$IterateAccounts++
				} Until ($IterateAccounts -eq $hMSDomain.Accounts.Count)
			}
			$IterateDomains++
		} Until ($IterateDomains -eq $hMS.Domains.Count)
	}

	Debug "----------------------------"
	Debug "Finished feeding $($TotalHamFedMessages + $TotalSpamFedMessages) messages to Bayes in $(ElapsedTime $BeginFeedingBayes)"
	Debug "----------------------------"
	
	If ($HamFedMessageErrors -gt 0) {
		Debug "Errors feeding HAM to SpamC : $HamFedMessageErrors Error$(Plural $HamFedMessageErrors) present"
		Email "[ERROR] HAM SpamC : $HamFedMessageErrors Errors present : Check debug log"
	}
	If ($TotalHamFedMessages -gt 0) {
		Debug "Bayes learned from $LearnedHamMessages of $TotalHamFedMessages HAM message$(Plural $TotalHamFedMessages) found"
		Email "[OK] Bayes HAM from $LearnedHamMessages of $TotalHamFedMessages message$(Plural $TotalHamFedMessages)"
	} Else {
		Debug "No HAM messages newer than $BayesDays days to feed to Bayes"
		Email "[OK] No HAM messages to feed Bayes"
	}

	If ($SpamFedMessageErrors -gt 0) {
		Debug "Errors feeding SPAM to SpamC : $SpamFedMessageErrors Error$(Plural $SpamFedMessageErrors) present"
		Email "[ERROR] SPAM SpamC : $SpamFedMessageErrors Errors present : Check debug log"
	}
	If ($TotalSpamFedMessages -gt 0) {
		Debug "Bayes learned from $LearnedSpamMessages of $TotalSpamFedMessages SPAM message$(Plural $TotalSpamFedMessages) found"
		Email "[OK] Bayes SPAM from $LearnedSpamMessages of $TotalSpamFedMessages message$(Plural $TotalSpamFedMessages)"
	} Else {
		Debug "No SPAM messages newer than $BayesDays days to feed to Bayes"
		Email "[OK] No SPAM messages to feed Bayes"
	}

	If ($SyncBayesJournal) {
		Debug "----------------------------"
		Try {
			$BayesSync = & cmd /c "`"$SADir\sa-learn.exe`" --sync"
			$BayesSyncResult = Out-String -InputObject $BayesSync
			If ([string]::IsNullOrEmpty($BayesSyncResult)) {
				Throw "Nothing to sync"
			}
			Debug $BayesSyncResult
		}
		Catch {
			Debug "[INFO] Bayes Journal Sync: $($Error[0])"
		}
	}

	If ($BackupBayesDatabase) {
		Debug "----------------------------"
		Try {
			If (-not(Test-Path $BayesBackupLocation)) {
				Throw "Bayes backup file does not exist - Check Path"
			} Else { 
				& cmd /c "`"$SADir\sa-learn.exe`" --backup > `"$BayesBackupLocation`""
				If ((Get-Item -Path $BayesBackupLocation).LastWriteTime -lt ((Get-Date).AddSeconds(-30))) {
					Throw "Unknown Error backing up Bayes database"
				}
				Debug "Successfully backed up Bayes database"
			}
		}
		Catch {
			Debug "[ERROR] backing up Bayes : $($Error[0])"
			Email "[ERROR] backing up Bayes db"
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
	$GetAccessTokenTries = 1
	Do {
		Try{
			$Auth = Invoke-RestMethod -Method GET $URIAuth -Body $AuthBody -ContentType 'application/json; charset=utf-8' 
			$AccessToken = $Auth.data.access_token
			$AccountID = $Auth.data.account_id
			Debug "Access Token : $AccessToken"
			Debug "Account ID   : $AccountID"
		}
		Catch {
			Debug "LetsUpload Authentication ERROR : Try $GetAccessTokenTries : $($Error[0])"
		}
		$GetAccessTokenTries++
	} Until (($GetAccessTokenTries -gt 5) -or ($AccessToken -match "^\w{128}$"))
	If ($GetAccessTokenTries -gt 5) {
		Debug "LetsUpload Authentication ERROR : Giving up"
		Email "[ERROR] LetsUpload Authentication : Check Debug Log"
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
	$CreateFolderTries = 1
	Do {
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
			Debug "LetsUpload Folder Creation ERROR : Try $CreateFolderTries : $($Error[0])"
		}
		$CreateFolderTries++
	} Until (($CreateFolderTries -gt 5) -or ($FolderID -match "^\d+$"))
	If ($CreateFolderTries -gt 5) {
		Debug "LetsUpload Folder Creation ERROR : Giving up"
		Email "[ERROR] LetsUpload Folder Creation Error : Check Debug Log"
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
	$TotalUploadErrors = 0

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
			Debug "$($Error[0])"
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
				If ($USize -ne $FileSize) {Throw "Local and remote filesizes do not match! Local file: $Filesize ::: Remote file: $USize"}
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
				Debug "[ERROR]  : $($Error[0])"
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
						Debug "Mismatched upload deleted. Trying again."
					}
					Catch {
						Debug "Mismatched upload file delete ERROR : $($Error[0])"
						Email "[ERROR] Un-Repairable Upload Mismatch! See debug log. Quitting script."
						EmailResults
						Exit
					}
				}
				$TotalUploadErrors++
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
		Debug "LetsUpload Folder Listing ERROR : $($Error[0])"
		Email "[ERROR] LetsUpload Folder Listing : Check Debug Log"
		Email "[ERROR] LetsUpload Folder Listing : $($Error[0])"
	}
	$FolderListingStatus = $FolderListing._status
	$RemoteFileCount = ($FolderListing.data.files.id).Count
	
	<#  Report results  #>
	If ($FolderListingStatus -match "success") {
		Debug "There are $RemoteFileCount file$(Plural $RemoteFileCount) in the remote folder"
		If ($RemoteFileCount -eq $CountArchVol) {
			Debug "----------------------------"
			Debug "Finished uploading $CountArchVol file$(Plural $CountArchVol) in $(ElapsedTime $BeginOffsiteUpload)"
			If ($TotalUploadErrors -gt 0){Debug "$TotalUploadErrors upload error$(Plural $TotalUploadErrors) successfully resolved"}
			Debug "Upload sucessful. $CountArchVol file$(Plural $CountArchVol) uploaded to $FolderURL"
			If ($UseHTML) {
				Email "[OK] $CountArchVol file$(Plural $CountArchVol) uploaded to <a href=`"$FolderURL`">letsupload.io</a>"
			} Else {
				Email "[OK] $CountArchVol file$(Plural $CountArchVol) uploaded to $FolderURL"
			}
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

<#  Check for updates  #>

Function CheckForUpdates {
	Debug "----------------------------"
	Debug "Checking for script update at GitHub"
	$GitHubVersion = $LocalVersion = $NULL
	$GetGitHubVersion = $GetLocalVersion = $False
	$GitHubVersionTries = 1
	Do {
		Try {
			$GitHubVersion = [decimal](Invoke-WebRequest -UseBasicParsing -Method GET -URI https://raw.githubusercontent.com/palinkas-jo-reggelt/hMailServer_Offsite_Backup/main/version.txt).Content
			$GetGitHubVersion = $True
		}
		Catch {
			Debug "[ERROR] Obtaining GitHub version : Try $GitHubVersionTries : Obtaining version number: $($Error[0])"
		}
		$GitHubVersionTries++
	} Until (($GitHubVersion -gt 0) -or ($GitHubVersionTries -eq 6))
	If (Test-Path "$PSScriptRoot\version.txt") {
		$LocalVersion = [decimal](Get-Content "$PSScriptRoot\version.txt")
		$GetLocalVersion = $True
	}
	If (($GetGitHubVersion) -and ($GetLocalVersion)) {
		If ($LocalVersion -lt $GitHubVersion) {
			Debug "[INFO] Upgrade to version $GitHubVersion available at https://github.com/palinkas-jo-reggelt/hMailServer_Offsite_Backup"
			If ($UseHTML) {
				Email "[INFO] Upgrade to version $GitHubVersion available at <a href=`"https://github.com/palinkas-jo-reggelt/hMailServer_Offsite_Backup`">GitHub</a>"
			} Else {
				Email "[INFO] Upgrade to version $GitHubVersion available at https://github.com/palinkas-jo-reggelt/hMailServer_Offsite_Backup"
			}
		} Else {
			Debug "Backup & Upload script is latest version: $GitHubVersion"
		}
	} Else {
		If ((-not($GetGitHubVersion)) -and (-not($GetLocalVersion))) {
			Debug "[ERROR] Version test failed : Could not obtain either GitHub nor local version information"
			Email "[ERROR] Version check failed"
		} ElseIf (-not($GetGitHubVersion)) {
			Debug "[ERROR] Version test failed : Could not obtain version information from GitHub"
			Email "[ERROR] Version check failed"
		} ElseIf (-not($GetLocalVersion)) {
			Debug "[ERROR] Version test failed : Could not obtain local install version information"
			Email "[ERROR] Version check failed"
		} Else {
			Debug "[ERROR] Version test failed : Unknown reason - file issue at GitHub"
			Email "[ERROR] Version check failed"
		}
	}
}