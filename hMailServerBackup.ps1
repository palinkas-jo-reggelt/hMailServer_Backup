<#

.SYNOPSIS
	hMailServer Backup

.DESCRIPTION
	hMailServer Backup

.FUNCTIONALITY
	Backs up hMailServer, compresses backup and uploads to LetsUpload

.PARAMETER 

	
.NOTES
	7-Zip required - install and place in system path
	Run at 11:58PM from task scheduler in order to properly cycle log files.
	
	
.EXAMPLE


#>

<###   CONFIG   ###>
Try {
	.("$PSScriptRoot\hMailServerBackupConfig.ps1")
	.("$PSScriptRoot\hMailServerBackupFunctions.ps1")
}
Catch {
	Write-Output "$(Get-Date -f G) : ERROR : Unable to load supporting PowerShell Scripts" | Out-File "$PSScriptRoot\PSError.log" -Append
	Write-Output "$(Get-Date -f G) : ERROR : $Error" | Out-File "$PSScriptRoot\PSError.log" -Append
	Exit
}

<###   BEGIN SCRIPT   ###>
$StartScript = Get-Date
$DateString = (Get-Date).ToString("yyyy-MM-dd")
$BackupName = "$DateString-hMailServer"

<#  Clear out error variable  #>
$Error.Clear()

<#  Remove trailing slashes from folder locations  #>
$hMSDir = $hMSDir -Replace('\\$','')
$SevenZipDir = $SevenZipDir -Replace('\\$','')
$MailDataDir = $MailDataDir -Replace('\\$','')
$BackupTempDir = $BackupTempDir -Replace('\\$','')
$BackupLocation = $BackupLocation -Replace('\\$','')
$SADir = $SADir -Replace('\\$','')
$SAConfDir = $SAConfDir -Replace('\\$','')
$MySQLBINdir = $MySQLBINdir -Replace('\\$','')

<#  Validate folders  #>
ValidateFolders $hMSDir
ValidateFolders $SevenZipDir
ValidateFolders $MailDataDir
ValidateFolders $BackupTempDir
ValidateFolders $BackupLocation
If ($UseSA) {
	ValidateFolders $SADir
	ValidateFolders $SAConfDir
}
If ($UseMySQL) {
	ValidateFolders $MySQLBINdir
}

<#  Delete old debug files and create new  #>
$EmailBody = "$PSScriptRoot\EmailBody.log"
If (Test-Path $EmailBody) {Remove-Item -Force -Path $EmailBody}
New-Item $EmailBody
$DebugLog = "$BackupLocation\hMailServerDebug-$DateString.log"
If (Test-Path $DebugLog) {Remove-Item -Force -Path $DebugLog}
New-Item $DebugLog
Write-Output "::: hMailServer Backup Routine $(Get-Date -f D) :::" | Out-File $DebugLog -Encoding ASCII -Append
Write-Output " " | Out-File $DebugLog -Encoding ASCII -Append
If ($UseHTML) {
	Write-Output "
		<!DOCTYPE html><html>
		<head><meta name=`"viewport`" content=`"width=device-width, initial-scale=1.0 `" /></head>
		<body style=`"font-family:Arial Narrow`"><table>
	" | Out-File $EmailBody -Encoding ASCII -Append
}

<#  Authenticate hMailServer COM  #>
$hMS = New-Object -COMObject hMailServer.Application
$hMS.Authenticate("Administrator", $hMSAdminPass) | Out-Null

<#  Get hMailServer Status  #>
$BootTime = [DateTime]::ParseExact((((Get-WmiObject -Class win32_operatingsystem).LastBootUpTime).Split(".")[0]), 'yyyyMMddHHmmss', $null)
$hMSStartTime = $hMS.Status.StartTime
$hMSSpamCount = $hMS.Status.RemovedSpamMessages
$hMSVirusCount = $hMS.Status.RemovedViruses
Debug "Last Reboot Time          : $(($BootTime).ToString('yyyy-MM-dd HH:mm:ss'))"
Debug "HMS Start Time            : $hMSStartTime"
Debug "HMS Daily Spam Reject     : $hMSSpamCount"
Debug "HMS Daily Viruses Removed : $hMSVirusCount"
If ($UseHTML) {
	Email "<center>:::&nbsp;&nbsp;&nbsp;hMailServer Backup Routine&nbsp;&nbsp;&nbsp;:::</center>"
	Email "<center>$(Get-Date -f D)</center>"
	Email " "
	Email "Last Reboot Time: $(($BootTime).ToString('yyyy-MM-dd HH:mm:ss'))"
	Email "HMS Start Time: $hMSStartTime"
	If ($hMSSpamCount -gt 0) {
		Email "HMS Daily Spam Reject count: <span style=`"background-color:red;color:white;font-weight:bold;font-family:Courier New;`">$hMSSpamCount</span>"
	} Else {
		Email "HMS Daily Spam Reject count: <span style=`"background-color:green;color:white;font-weight:bold;font-family:Courier New;`">0</span>"
	}
	If ($hMSVirusCount -gt 0) {
		Email "HMS Daily Viruses Removed count: <span style=`"background-color:red;color:white;font-weight:bold;font-family:Courier New;`">$hMSVirusCount</span>"
	} Else {
		Email "HMS Daily Viruses Removed count: <span style=`"background-color:green;color:white;font-weight:bold;font-family:Courier New;`">0</span>"
	}
	Email " "
} Else {
	Email ":::   hMailServer Backup Routine   :::"
	Email "       $(Get-Date -f D)"
	Email " "
	Email "Last Reboot Time: $(($BootTime).ToString('yyyy-MM-dd HH:mm:ss'))"
	Email "HMS Start Time: $hMSStartTime"
	Email "HMS Daily Spam Reject count: $hMSSpamCount"
	Email "HMS Daily Viruses Removed count: $hMSVirusCount"
	Email " "
}

<#  Stop hMailServer & SpamAssassin services #>
ServiceStop $hMSServiceName
If ($UseSA) {ServiceStop $SAServiceName}

<#  Update SpamAssassin  #>
If ($UseSA) {
	Debug "----------------------------"
	Debug "Updating SpamAssassin"
	$BeginSAUpdate = Get-Date
	$SAUD = "$SADir\sa-update.exe"
	Try {
		$SAUpdate = & $SAUD -v --nogpg --channel updates.spamassassin.org | Out-String
		Debug $SAUpdate
		Debug "Finished updating SpamAssassin in $(ElapsedTime $BeginSAUpdate)"
		Email "[OK] SpamAssassin updated"
		If ($SAUpdate -match "Update finished, no fresh updates were available"){
			Email "[INFO] No fresh updates available"
		}
	}
	Catch {
		$Err = $Error[0]
		Debug "[ERROR] SpamAssassin update : $Err"
		Email "[ERROR] SpamAssassin update : Check Debug Log"
	}
}

<#  Backup files using RoboCopy  #>
$BackupSuccess = 0
Debug "----------------------------"
Debug "Start backing up datadir with RoboCopy"
$BeginRobocopy = Get-Date
Try {
	$RoboCopy = & robocopy $MailDataDir "$BackupTempDir\hMailData" /mir /ndl /r:43200 /np /w:1 | Out-String
	Debug $RoboCopy
	Debug "Finished backing up data dir in $(ElapsedTime $BeginRobocopy)"
	$RoboStats = $RoboCopy.Split([Environment]::NewLine) | Where-Object {$_ -match 'Files\s:\s+\d'} 
	$RoboStats | ConvertFrom-String -Delimiter "\s+" -PropertyNames Nothing, Files, Colon, Total, Copied, Skipped, Mismatch, Failed, Extras | ForEach {
		$Copied = $_.Copied
		$Mismatch = $_.Mismatch
		$Failed = $_.Failed
		$Extras = $_.Extras
	}
	If (($Mismatch -gt 0) -or ($Failed -gt 0)) {
		Throw "Robocopy MISMATCH or FAILED exists"
	}
	$BackupSuccess++
	Debug "Robocopy backup success: $Copied new, $Extras deleted, $Mismatch mismatched, $Failed failed"
	Email "[OK] DataDir backed up: $Copied new, $Extras del"
}
Catch {
	$Err = $Error[0]
	Debug "[ERROR] RoboCopy : $Err"
	Email "[ERROR] RoboCopy : Check Debug Log"
}

<#  Backup database files  #>
$BeginDBBackup = Get-Date
If ($UseMySQL) {
	$Error.Clear()
	Debug "----------------------------"
	Debug "Begin backing up MySQL"
	If (Test-Path "$BackupTempDir\hMailData\MYSQLDump_*.sql") {
		Debug "Deleting old MySQL database dump"
		Try {
			Remove-Item "$BackupTempDir\hMailData\*.sql"
			Debug "Old MySQL database successfully deleted"
		}
		Catch {
			$Err = $Error[0]
			Debug "[ERROR] Old MySQL database delete : $Err"
			Email "[ERROR] Old MySQL database delete : Check Debug Log"
		}
	}
	$MySQLDump = "$MySQLBINdir\mysqldump.exe"
	$MySQLDumpPass = "-p$MySQLPass"
	$MySQLDumpFile = "$BackupTempDir\hMailData\MYSQLDump_$((Get-Date).ToString('yyyy-MM-dd')).sql"
	Try {
		If ($BackupAllMySQLDatbase) {
			& $MySQLDump -u $MySQLUser $MySQLDumpPass --all-databases --result-file=$MySQLDumpFile
		} Else {
			& $MySQLDump -u $MySQLUser $MySQLDumpPass $MySQLDatabase --result-file=$MySQLDumpFile
		}
		$BackupSuccess++
		Debug "MySQL successfully dumped in $(ElapsedTime $BeginDBBackup)"
	}
	Catch {
		$Err = $Error[0]
		Debug "[ERROR] MySQL Dump : $Err"
		Email "[ERROR] MySQL Dump : Check Debug Log"
	}
} Else {
	Debug "----------------------------"
	Debug "Begin backing up internal database"
	Debug "Copy internal database to backup folder"
	Try {
		$RoboCopyIDB = & robocopy "$hMSDir\Database" "$BackupTempDir\hMailData" /mir /ndl /r:43200 /np /w:1 | Out-String
		$BackupSuccess++
		Debug $RoboCopyIDB
		Debug "Internal DB successfully backed up in $(ElapsedTime $BeginDBBackup)"
	}
	Catch {
		$Err = $Error[0]
		Debug "[ERROR] RoboCopy Internal DB : $Err"
		Email "[ERROR] RoboCopy Internal DB : Check Debug Log"
	}
}

<#  Backup Miscellaneous Files  #>
Debug "----------------------------"
Debug "Begin backing up miscellaneous files"
$MiscBackupFiles | ForEach {
	$MBUF = $_
	$MBUFName = Split-Path -Path $MBUF -Leaf
	If (Test-Path "$BackupTempDir\hMailData\$MBUFName") {
		Remove-Item -Force -Path "$BackupTempDir\hMailData\$MBUFName"
		Debug "Previously backed up $MBUFName successfully deleted"
	} 
	If (Test-Path $MBUF) {
		Try {
			Copy-Item -Path $MBUF -Destination "$BackupTempDir\hMailData"
			$BackupSuccess++
			Debug "$MBUFName successfully backed up"
		}
		Catch {
			$Err = $Error[0]
			Debug "[ERROR] $MBUF Backup : $Err"
			Email "[ERROR] Backup $MBUFName : Check Debug Log"
		}
	} Else {
		Debug "$MBUF copy ERROR : File path not validated"
	}
}

<#  Report backup success  #>
If ($BackupSuccess -eq (2 + ($MiscBackupFiles).Count)) {
	Email "[OK] Backed up data dir, db and misc files"
} Else {
	Email "[ERROR] Backup count mismatch : Check Debug Log"
}

<#  Cycle Logs  #>
If ($CycleLogs) {
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
				$Err = $Error[0]
				Debug "[ERROR] Log cylcling ERROR : $Err"
			}
		} Else {
			Debug "[ERROR] $FullName not found"
		}
	}
}

<#  Restart SpamAssassin and hMailServer  #>
If ($UseSA) {ServiceStart $SAServiceName}
ServiceStart $hMSServiceName

<#  Prune hMailServer logs  #>
If ($PruneLogs) {
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
			$Err = $Error[0]
			Debug "[ERROR] Prune hMailServer logs : $Err"
			Email "[ERROR] Prune hMailServer logs : Check Debug Log"
		}
	}
}

<#  Prune backups  #>
If ($PruneBackups) {
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
			$Err = $Error[0]
			Debug "[ERROR] Prune backups : $Err"
			Email "[ERROR] Prune backups : Check Debug Log"
		}
	}
}

<#  Prune messages/empty folders older than N number of days  #>
If ($PruneMessages) {PruneMessages}

<#  Feed Beyesian database  #>
If ($UseSA) {If ($FeedBayes) {FeedBayes}}

<#  Compress backup into 7z archives  #>
If ($UseSevenZip) {MakeArchive}

<#  Upload archive to LetsUpload.io  #>
If ($UseLetsUpload) {OffsiteUpload}

<#  Check for updates  #>
$GitHubVersion = [decimal](Invoke-WebRequest -Method GET -Uri https://raw.githubusercontent.com/palinkas-jo-reggelt/hMailServer_Offsite_Backup/main/version.txt).Content
If (Test-Path "$PSScriptRoot\version.txt") {
	$LocalVersion = [decimal](Get-Content "$PSScriptRoot\version.txt")
} Else {
	$LocalVersion = 0.1
}
If ($LocalVersion -lt $GitHubVersion) {
	Debug "----------------------------"
	Debug "Upgrade available at https://github.com/palinkas-jo-reggelt/hMailServer_Offsite_Backup"
	If ($UseHTML) {
		Email "[INFO] Upgrade available at <a href=`"https://github.com/palinkas-jo-reggelt/hMailServer_Offsite_Backup`">GitHub</a>"
	} Else {
		Email "[INFO] Upgrade available at https://github.com/palinkas-jo-reggelt/hMailServer_Offsite_Backup"
	}
}

<#  Finish up and send email  #>
Debug "hMailServer Backup & Upload routine completed in $(ElapsedTime $StartScript)"
Email " "
Email "hMailServer Backup & Upload routine completed in $(ElapsedTime $StartScript)"
If ($UseHTML) {Write-Output "</table></body></html>" | Out-File $EmailBody -Encoding ASCII -Append}
EmailResults