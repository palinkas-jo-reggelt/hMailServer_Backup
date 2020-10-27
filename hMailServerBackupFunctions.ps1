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


Function Debug ($DebugOutput) {
	If ($VerboseFile) {Write-Output "$(Get-Date -f G) : $DebugOutput" | Out-File $DebugLog -Encoding ASCII -Append}
	If ($VerboseConsole) {Write-Host "$(Get-Date -f G) : $DebugOutput"}
}

Function Email ($EmailOutput) {
	Write-Output $EmailOutput | Out-File $EmailBody -Encoding ASCII -Append
}

Function EmailResults {
	Try {
		$Body = (Get-Content -Path $EmailBody | Out-String )
		If (($AttachDebugLog) -and (Test-Path $DebugLog) -and (((Get-Item $DebugLog).length/1MB) -lt $MaxAttachmentSize)){$Attachment = New-Object System.Net.Mail.Attachment $DebugLog}
		$Message = New-Object System.Net.Mail.Mailmessage $EmailFrom, $EmailTo, $Subject, $Body
		$Message.IsBodyHTML = $HTML
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
		Email "Error : Folder location $Folder does not exist : Quitting script"
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

Function ServiceStart ($ServiceName) {
	<#  Check to see if already running  #>
	If ($ServiceName -eq $hMSServiceName) {$ServiceDescription = "hMailServer"}
	If ($ServiceName -eq $SAServiceName) {$ServiceDescription = "SpamAssassin"}
	Debug "----------------------------"
	Debug "Start $ServiceDescription"
	$ServiceStopped = $False
	$hMSServiceStart = $False
	$SAServiceStart = $False
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
			Start-Sleep -Seconds 60
			(Get-Service $ServiceName).Refresh()
			$ServiceStatus = (Get-Service $ServiceName).Status
		} Until (((New-Timespan -Start $BeginStartup -End (Get-Date)).TotalMinutes -gt $ServiceTimeout) -or ($ServiceStatus -eq "Running"))

		If ($ServiceStatus -ne "Running"){
			Debug "$ServiceDescription failed to start"
			GmailResults "$ServiceDescription failed to start during backup routine! Check status NOW!"
			Break
		} Else {
			Debug "$ServiceDescription successfully started"
			Email "* $ServiceDescription successfully started"
			If ($ServiceDescription -eq $hMSServiceName) {$hMSServiceStart = $True}
			If ($ServiceDescription -eq $SAServiceName) {$SAServiceStart = $True}
		}
	}
}

Function ServiceStop ($ServiceName) {
	<#  Check to see if already stopped  #>
	If ($ServiceName -eq $hMSServiceName) {$ServiceDescription = "hMailServer"}
	If ($ServiceName -eq $SAServiceName) {$ServiceDescription = "SpamAssassin"}
	Debug "----------------------------"
	Debug "Stop $ServiceDescription"
	$ServiceRunning = $False
	$hMSServiceStop = $False
	$SAServiceStop = $False
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
			Start-Sleep -Seconds 60
			(Get-Service $ServiceName).Refresh()
			$ServiceStatus = (Get-Service $ServiceName).Status
		} Until (((New-Timespan $BeginShutdown).TotalMinutes -gt $ServiceTimeout) -or ($ServiceStatus -eq "Stopped"))

		If ($ServiceStatus -ne "Stopped"){
			Debug "$ServiceDescription failed to stop."
			GmailResults "$ServiceDescription failed to stop during backup routine! Check status NOW!"
			Break
		} Else {
			Debug "$ServiceDescription successfully stopped"
			Email "* $ServiceDescription successfully stopped"
			If ($ServiceDescription -eq $hMSServiceName) {$hMSServiceStop = $True}
			If ($ServiceDescription -eq $SAServiceName) {$SAServiceStop = $True}
		}
	}
}

Function MakeArchive {
	$StartArchive = Get-Date
	$MakeArchiveSuccess = $False
	Debug "----------------------------"
	Debug "Create archive : $BackupName"
	Debug "Archive folder : $BackupTempDir"
	$VolumeSwitch = "-v$VolumeSize"
	$PWSwitch = "-p$ArchivePassword"
	Try {
		& cmd /c 7z a $VolumeSwitch -t7z -m0=lzma2 -mx=9 -mfb=64 -md=32m -ms=on -mhe=on $PWSwitch "$BackupLocation\$BackupName\$BackupName.7z" "$BackupTempDir\*"
		Debug "Archive creation finished in $(ElapsedTime $StartArchive)"
		Debug "Wait a few seconds to make sure archive is finished"
		Email "* 7-Zip archive of backup files creation successful"
		Start-Sleep -Seconds 3
		$MakeArchiveSuccess = $True
	}
	Catch {
		Debug "Archive Creation ERROR : $Error"
		Email "[ERROR] Archive Creation : Check Debug Log"
		Email "[ERROR] Archive Creation : $Error"
		EmailResults
		Exit
	}
}

