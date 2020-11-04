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
	Write-Output "$(Get-Date) -f G) : ERROR : Unable to load supporting PowerShell Scripts" | out-file "$PSScriptRoot\PSError.log" -append
	Write-Output "$(Get-Date) -f G) : ERROR : $Error" | out-file "$PSScriptRoot\PSError.log" -append
	Exit
}

<###   BEGIN SCRIPT   ###>
$StartScript = Get-Date
$DateString = (Get-Date).ToString("yyyy-MM-dd")
$BackupName = "$DateString-hMailServer"

<#  Clear out error variable  #>
$Error.Clear()

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

<#  Validate folders  #>
$hMSDir = $hMSDir -Replace('\\$','')
ValidateFolders $hMSDir
$MailDataDir = $MailDataDir -Replace('\\$','')
ValidateFolders $MailDataDir
$BackupTempDir = $BackupTempDir -Replace('\\$','')
ValidateFolders $BackupTempDir -Replace('\\$','')
$BackupLocation = $BackupLocation
$SADir = $SADir -Replace('\\$','')
If ($UseSA) {ValidateFolders $SADir}
$SAConfDir = $SAConfDir -Replace('\\$','')
If ($UseSA) {ValidateFolders $SAConfDir}

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
		Email "HMS Daily Spam Reject count: 0"
	}
	If ($hMSVirusCount -gt 0) {
		Email "HMS Daily Viruses Removed count: <span style=`"background-color:red;color:white;font-weight:bold;font-family:Courier New;`">$hMSVirusCount</span>"
	} Else {
		Email "HMS Daily Viruses Removed count: 0"
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
	$SACF = "$SADir\UpdateChannels.txt"
	Try {
		$SAUpdate = & $SAUD -v --nogpg --channelfile $SACF | Out-String
		Debug $SAUpdate
		Debug "Finished updating SpamAssassin in $(ElapsedTime $BeginSAUpdate)"
		Email "[OK] SpamAssassin updated"
		If ($SAUpdate -match "Update finished, no fresh updates were available"){
			Email "[INFO] No fresh SpamAssassin updates available"
		}
	}
	Catch {
		Debug "SpamAssassin update ERROR : $Error"
		Email "[ERROR] SpamAssassin update : Check Debug Log"
	}
}

<#  Cycle Logs  #>
If ($CycleLogs) {
	Debug "----------------------------"
	Debug "Cycling Logs"
	If (Test-Path "$hMSDir\Logs\hmailserver_events.log") {
		Try {Rename-Item "$hMSDir\Logs\hmailserver_events.log" "hmailserver_events_$((Get-Date).ToString('yyyy-MM-dd')).log"} 
		Catch {Debug "SpamAssassin update ERROR : $Error"}
		Debug "Cylcled hmailserver_events_$((Get-Date).ToString('yyyy-MM-dd')).log"
	} Else {
		Debug "hmailserver_events.log not found"
	}
	If ($UseSA) {
		If (Test-Path "$hMSDir\Logs\spamd.log") {
			Try {Rename-Item "$hMSDir\Logs\spamd.log" "spamd_$((Get-Date).ToString('yyyy-MM-dd')).log"} Catch {Debug "SpamAssassin update ERROR : $Error"}
			Debug "Cylcled spamd_$((Get-Date).ToString('yyyy-MM-dd')).log"
		} Else {
			Debug "spamd.log not found"
		}
	}
}

<#  Prune hMailServer logs  #>
If ($PruneLogs) {
	$FilesToDel = Get-ChildItem -Path "$hMSDir\Logs"  | Where-Object {$_.LastWriteTime -lt ((Get-Date).AddDays(-$DaysToKeepLogs))}
	$CountDelLogs = (Test-Path $FilesToDel -PathType Leaf).Count
	If ($CountDelLogs -gt 0) {
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
				Debug "Prune hMailServer logs ERROR : Filecount does not match delete count"
				Email "[ERROR] Prune hMailServer logs : Check Debug Log"
			}
		}
		Catch {
			Debug "Prune hMailServer logs ERROR : $Error"
			Email "[ERROR] Prune hMailServer logs : Check Debug Log"
		}
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
	Debug "Robocopy backup success: $Copied new, $Extras deleted, $Mismatch mismatched, $Failed failed"
	Email "[OK] hMailServer DataDir backed up:"
	Email "[OK] $Copied new, $Extras deleted, $Mismatch mismatched, $Failed failed"
}
Catch {
	Debug "RoboCopy ERROR : $Error"
	Email "[ERROR] RoboCopy : Check Debug Log"
}

<#  Backup database files  #>
$BeginDBBackup = Get-Date
If ($UseMySQL) {
	Debug "----------------------------"
	Debug "Begin backing up MySQL"
	Debug "Deleting old MySQL database dump"
	Try {
		Remove-Item "$BackupTempDir\hMailData\*.mysql"
		Debug "Old MySQL database successfully deleted"
	}
	Catch {
		Debug "Old MySQL database delete ERROR : $Error"
		Email "[ERROR] Old MySQL database delete : Check Debug Log"
	}
	Debug "Backing up MySQL"
	$MySQLDump = "$MySQLBINdir\mysqldump.exe"
	$MySQLDumpPass = "-p$MySQLPass"
	Try {
		& $MySQLDump -u $MySQLUser $MySQLDumpPass hMailServer > "$BackupTempDir\hMailData\MYSQLDump_$((Get-Date).ToString('yyyy-MM-dd')).mysql" | Out-String
		$BackupSuccess++
		Debug "MySQL successfully dumped in $(ElapsedTime $BeginDBBackup)"
	}
	Catch {
		Debug "MySQL Dump ERROR : $Error"
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
		Debug "RoboCopy Internal DB ERROR : $Error"
		Email "[ERROR] RoboCopy Internal DB : Check Debug Log"
	}
}

<#  Backup hMailServer.ini  #>
$BeginINIBackup = Get-Date
Debug "----------------------------"
Debug "Begin backing up hMailServer.ini"
Debug "Deleting old backup copy of hMailServer.ini"
If (Test-Path "$BackupTempDir\hMailData\hMailServer.ini") {
	Remove-Item -Force -Path "$BackupTempDir\hMailData\hMailServer.ini"
	Debug "Old hMailServer.ini successfully deleted"
} Else {
	Debug "No old backup copy of hMailServer.ini exists"
}
Debug "Backing up server copy of hMailServer.ini"
If (Test-Path "$hMSDir\Bin\hMailServer.INI") {
	Try {
		Copy-Item -Path "$hMSDir\Bin\hMailServer.INI" -Destination "$BackupTempDir\hMailData"
		$BackupSuccess++
		Debug "hMailServer.ini successfully backed up"
	}
	Catch {
		Debug "hMailServer.ini Backup ERROR : $Error"
		Email "[ERROR] Backup hMailServer.ini : Check Debug Log"
	}
} Else {
	Debug "hMailServer.ini copy ERROR : File path not validated"
}

<#  Report backup success  #>
If ($BackupSuccess -eq 2) {
	Email "[OK] Backed up hMailServer DB and INI"
}

<#  Restart SpamAssassin and hMailServer  #>
If ($UseSA) {ServiceStart $SAServiceName}
ServiceStart $hMSServiceName

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
				Debug "Prune backups ERROR : Filecount does not match delete count"
				Email "[ERROR] Prune backups : Check Debug Log"
			}
		}
		Catch {
			Debug "Prune backups ERROR : $Error"
			Email "[ERROR] Prune backups : Check Debug Log"
		}
	}
}

<#  Delete messages/empty folders older than N number of days  #>
PruneMessages

<#  Compress backup into 7z archives  #>
MakeArchive

<#  Upload archive to LetsUpload.io  #>
OffsiteUpload

<#  Finish up and send email  #>
Debug "hMailServer Backup & Upload routine completed in $(ElapsedTime $StartScript)"
Email " "
Email "hMailServer Backup & Upload routine completed in $(ElapsedTime $StartScript)"
If ($UseHTML) {Write-Output "</table></body></html>" | Out-File $EmailBody -Encoding ASCII -Append}
EmailResults