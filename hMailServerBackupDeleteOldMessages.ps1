<#

.SYNOPSIS
	hMailServer Backup

.DESCRIPTION
	Delete messages in specified folders older than N days

.FUNCTIONALITY
	Backs up hMailServer, compresses backup and uploads to LetsUpload

.PARAMETER 

	
.NOTES
	7-Zip required - install and place in system path
	Run at 12:58PM from task scheduler
	
	
.EXAMPLE


#>

Function DeleteOldMessages {
	
	$BeginDeletingOldMessages = Get-Date
	Debug "----------------------------"
	Debug "Begin deleting old messages"

	<#  Authenticate hMailServer COM  #>
	$hMS = New-Object -COMObject hMailServer.Application
	$hMS.Authenticate("Administrator", $hMSAdminPass) | Out-Null
	
	$ArrayTotalCount = @()

	$EnumDomain = 0
	Do {
		$hMSDomain = $hMS.Domains.Item($EnumDomain)

		If ($hMSDomain.Active) {
			$EnumAccount = 0

			Do {
				$hMSAccount = $hMSDomain.Accounts.Item($EnumAccount)

				If ($hMSAccount.Active) {
					$AccountAddress = $hMSAccount.Address
					$EnumFolder = 0

					Do {
						$hMSIMAPFolder = $hMSAccount.IMAPFolders.Item($EnumFolder)
						$ImapFolderName = $hMSIMAPFolder.Name

						If ($ImapFolderName -match [regex]$CleanupFolders) {

							If ($hMSIMAPFolder.SubFolders.Count -gt 0) {
								$EnumSubFolder = 0

								Do {
									$SubFolders = $hMSIMAPFolder.SubFolders.Item($EnumSubFolder)
									$SFName = $SubFolders.Name

									If ($SubFolders.Subfolders.Count -gt 0) {
										$EnumSubSubFolder = 0

										Do {
											$SubSubFolders = $SubFolders.SubFolders.Item($EnumSubSubFolder)
											$SsFName = $SubSubFolders.Name

											If ($SubSubFolders.Subfolders.Count -gt 0) {
												$EnumSubSubSubFolder = 0

												Do {
													$SubSubSubFolders = $SubSubFolders.SubFolders.Item($EnumSubSubSubFolder)
													$SssFName = $SubSubSubFolders.Name

													If ($SubSubSubFolders.Messages.Count -gt 0) {
														$EnumSubSubSubFolderMsg = 0
														$SssFDeleteCount = 0
														$ArraySssFMsgID = @()

														Do {
															$SssFMsg = $SubSubSubFolders.Messages.Item($EnumSubSubSubFolderMsg)

															If (($SssFMsg.InternalDate) -lt ((Get-Date).AddDays(-$DaysBeforeDelete))){
																$ArraySssFMsgID += $SssFMsg.ID
																$ArrayTotalCount += $SssFMsg.ID
																$SssFDeleteCount++
															}

															$EnumSubSubSubFolderMsg++

														} Until ($EnumsubSubSubFolderMsg -eq $SubSubSubFolders.Messages.Count)

														$ArraySssFMsgID | ForEach {
															$SubSubSubFolders.Messages.DeleteByDBID($_)
														}

														If ($SssFDeleteCount -gt 0) {
															Debug "Deleted $SssFDeleteCount messages older than $DaysBeforeDelete days in $AccountAddress > $ImapFolderName > $SFName > $SsFName > $SssFName"
														}

													} # IF SUBFOLDER LEVEL 3 MESSAGES > 0

													$EnumSubSubSubFolder++

												} Until ($EnumSubSubSubFolder -eq $SubSubFolders.Subfolders.Count)

											} # IF SUBFOLDER LEVEL 3 COUNT > 0

											If ($SubSubFolders.Messages.Count -gt 0) {
												$EnumSubSubFolderMsg = 0
												$SsFDeleteCount = 0
												$ArraySsFMsgID = @()

												Do {
													$SsFMsg = $SubSubFolders.Messages.Item($EnumSubSubFolderMsg)

													If (($SsFMsg.InternalDate) -lt ((Get-Date).AddDays(-$DaysBeforeDelete))){
														$ArraySsFMsgID += $SsFMsg.ID
														$ArrayTotalCount += $SsFMsg.ID
														$SsFDeleteCount++
													}

													$EnumSubSubFolderMsg++

												} Until ($EnumSubSubFolderMsg -eq $SubSubFolders.Messages.Count)

												$ArraySsFMsgID | ForEach {
													$SubSubFolders.Messages.DeleteByDBID($_)
												}

												If ($SsFDeleteCount -gt 0) {
													Debug "Deleted $SsFDeleteCount messages older than $DaysBeforeDelete days in $AccountAddress > $ImapFolderName > $SFName > $SsFName"
												}

											} # IF SUBFOLDER MESSAGES > 0

											$EnumSubSubFolder++

										} Until ($EnumSubSubFolder -eq $SubFolders.Subfolders.Count)

									} #IF SUBFOLDER LEVEL 2 COUNT > 0

									If ($SubFolders.Messages.Count -gt 0) {
										$EnumSubFolderMsg = 0
										$SFDeleteCount = 0
										$ArraySFMsgID = @()

										Do {
											$SFMsg = $SubFolders.Messages.Item($EnumSubFolderMsg)
											$SFMsgDate = $SFMsg.InternalDate

											If ($SFMsgDate -lt ((Get-Date).AddDays(-$DaysBeforeDelete))){
												$ArraySFMsgID += $SFMsg.ID
												$ArrayTotalCount += $SFMsg.ID
												$SFDeleteCount++
											}

											$EnumSubFolderMsg++

										} Until ($EnumSubFolderMsg -eq $SubFolders.Messages.Count)

										$ArraySFMsgID | ForEach {
											$SubFolders.Messages.DeleteByDBID($_)
										}

										If ($SFDeleteCount -gt 0) {
											Debug "Deleted $SFDeleteCount messages older than $DaysBeforeDelete days in $AccountAddress > $ImapFolderName > $SFName"
										}

									} # IF SUBFOLDER MESSAGES > 0

									$EnumSubFolder++

								} Until ($EnumSubFolder -eq $hMSIMAPFolder.SubFolders.Count)

							} # IF SUBFOLDER COUNT > 0

							$hMSMessages = $hMSIMAPFolder.Messages
							$MsgCount = $hMSMessages.Count

							If ($MsgCount -gt 0) {
								$EnumMessage = 0
								$DeleteCount = 0
								$ArrayMsgID = @()

								Do {
									$ItemMsg = $hMSMessages.Item($EnumMessage)
									$MsgDate = $ItemMsg.InternalDate

									If ($MsgDate -lt ((Get-Date).AddDays(-$DaysBeforeDelete))){
										$ArrayMsgID += $ItemMsg.ID
										$ArrayTotalCount += $ItemMsg.ID
										$DeleteCount++
									}

									$EnumMessage++

								} Until ($EnumMessage -eq $MsgCount)

								$ArrayMsgID | ForEach {
									$hMSMessages.DeleteByDBID($_)
								}

								If ($DeleteCount -gt 0) {
									Debug "Deleted $DeleteCount messages older than $DaysBeforeDelete days in $AccountAddress > $ImapFolderName"
								}

							} # IF FOLDER MESSAGES > 0

						} # IF FOLDERNAME MATCH REGEX

						$EnumFolder++

					} Until ($EnumFolder -eq $hMSAccount.IMAPFolders.Count)

				} #IF ACCOUNT ACTIVE

				$EnumAccount++

			} Until ($EnumAccount -eq $hMSDomain.Accounts.Count)

		} # IF DOMAIN ACTIVE

		$EnumDomain++

	} Until ($EnumDomain -eq $hMS.Domains.Count)

	$CountArrayTotalCount = $ArrayTotalCount.Count
	Debug "Finished deleting $CountArrayTotalCount messages in $(ElapsedTime $BeginDeletingOldMessages)"
	If ($CountArrayTotalCount -gt 0) {
		Email "* Deleted $CountArrayTotalCount messages in specified folders older than $DaysBeforeDelete days successfully"
	} Else {
		Email "* No messages in specified folders older than $DaysBeforeDelete days to delete"
	}

} # END FUNCTION