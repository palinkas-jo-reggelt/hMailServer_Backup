<#

.SYNOPSIS
	hMailServer Backup

.DESCRIPTION
	hMailServer Backup

.FUNCTIONALITY
	Backs up hMailServer, compresses backup and uploads to LetsUpload.io

.PARAMETER 

	
.NOTES
	Run at 11:58PM from task scheduler in order to properly cycle log files.
	
.EXAMPLE


#>

<###   LOAD SUPPORTING FILES   ###>
Try {
	.("$PSScriptRoot\hMailServerBackupConfig.ps1")
	.("$PSScriptRoot\hMailServerBackupFunctions.ps1")
}
Catch {
	Write-Output "$(Get-Date -f G) : ERROR : Unable to load supporting PowerShell Scripts" | Out-File "$PSScriptRoot\PSError.log" -Append
	Write-Output "$(Get-Date -f G) : ERROR : $($Error[0])" | Out-File "$PSScriptRoot\PSError.log" -Append
	Exit
}

<###   BEGIN SCRIPT   ###>
$StartScript = Get-Date
$DateString = (Get-Date).ToString("yyyy-MM-dd")
$BackupName = "$DateString-hMailServer"

<#  Clear out error variable  #>
$Error.Clear()

<#  Set counting variables that pass through functions  #>
Set-Variable -Name BackupSuccess -Value 0 -Option AllScope
Set-Variable -Name DoBackupDataDir -Value 0 -Option AllScope
Set-Variable -Name DoBackupDB -Value 0 -Option AllScope
Set-Variable -Name MiscBackupSuccess -Value 0 -Option AllScope
Set-Variable -Name TotalDeletedMessages -Value 0 -Option AllScope
Set-Variable -Name TotalDeletedFolders -Value 0 -Option AllScope
Set-Variable -Name DeleteMessageErrors -Value 0 -Option AllScope
Set-Variable -Name DeleteFolderErrors -Value 0 -Option AllScope
Set-Variable -Name TotalHamFedMessages -Value 0 -Option AllScope
Set-Variable -Name TotalSpamFedMessages -Value 0 -Option AllScope
Set-Variable -Name HamFedMessageErrors -Value 0 -Option AllScope
Set-Variable -Name SpamFedMessageErrors -Value 0 -Option AllScope
Set-Variable -Name LearnedHamMessages -Value 0 -Option AllScope
Set-Variable -Name LearnedSpamMessages -Value 0 -Option AllScope

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

<# Create hMailData folder if it doesn't exist #>
If (-not(Test-Path "$BackupTempDir\hMailData" -PathType Container)) {md "$BackupTempDir\hMailData"}

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
Debug "Server Uptime             : $(ElapsedTime (($BootTime).ToString('yyyy-MM-dd HH:mm:ss')))"
Debug "HMS Start Time            : $hMSStartTime"
Debug "HMS Uptime                : $(ElapsedTime $hMSStartTime)"
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
$BeginShutdownPeriod = Get-Date
ServiceStop $hMSServiceName
If ($UseSA) {ServiceStop $SAServiceName}

<#  Cycle Logs  #>
If ($CycleLogs) {CycleLogs}

<#  Update SpamAssassin  #>
If ($UseSA) {UpdateSpamassassin}

<#  Update Custom Rulesets  #>
If ($UseSA) {If ($UseCustomRuleSets) {UpdateCustomRulesets}}

<#  Backup files using RoboCopy  #>
If ($BackupDataDir) {BackuphMailDataDir}

<#  Backup database files  #>
If ($BackupDB) {BackupDatabases}

<#  Backup Miscellaneous Files  #>
$MiscBackupSuccess = 0
If ($BackupMisc) {BackupMiscellaneousFiles}

<#  Report backup success  #>
Debug "----------------------------"
If ($BackupSuccess -eq ($DoBackupDataDir + $DoBackupDB + $MiscBackupFiles.Count)) {
	Debug "All files backed up successfully"
	Email "[OK] Backed up data successfully"
} Else {
	Debug "[ERROR] Backup count mismatch."
	Email "[ERROR] Backup count mismatch : Check Debug Log"
}

<#  Restart SpamAssassin and hMailServer  #>
If ($UseSA) {ServiceStart $SAServiceName}
ServiceStart $hMSServiceName
Debug "----------------------------"
Debug "hMailServer was out of service for $(ElapsedTime $BeginShutdownPeriod)"

<#  Prune hMailServer logs  #>
If ($PruneLogs) {PruneLogs}

<#  Prune backups  #>
If ($PruneBackups) {PruneBackups}

<#  Prune messages/empty folders older than N number of days  #>
If ($PruneMessages) {PruneMessages}

<#  Feed Beyesian database  #>
If ($UseSA) {If ($FeedBayes) {FeedBayes}}

<#  Compress backup into 7z archives  #>
If ($UseSevenZip) {MakeArchive}

<#  Upload archive to LetsUpload.io  #>
If ($UseLetsUpload) {OffsiteUpload}

<#  Check for updates  #>
CheckForUpdates

<#  Finish up and send email  #>
Debug "----------------------------"
If (((Get-Item $DebugLog).length/1MB) -ge $MaxAttachmentSize) {
	Debug "Debug log larger than specified max attachment size. Do not attach to email message."
	Email "[INFO] Debug log larger than specified max attachment size. Log file stored in backup folder on server file system."
}
Debug "hMailServer Backup & Upload routine completed in $(ElapsedTime $StartScript)"
Email " "
Email "hMailServer Backup & Upload routine completed in $(ElapsedTime $StartScript)"
If ($UseHTML) {Write-Output "</table></body></html>" | Out-File $EmailBody -Encoding ASCII -Append}
EmailResults