# hMailServer Offsite Backup
 hMailServer backup routine that also uploads to LetsUpload.io
 
 Discussion thread: https://hmailserver.com/forum/viewtopic.php?f=9&t=35447
 
# What does it do?
 1) Stops hMailServer and SpamAssassin Services
 2) Updates SpamAssassin
 3) Cycles hMailServer and SpamAssassin logs
 4) Backs up hMailServer data directory using RoboCopy
 5) Dumps MySQL database -or- internal database and adds to backup directory
 6) Backs up miscellaneous files
 7) Restarts SpamAssassin and hMailServer
 8) Prunes messages older than specified number of days in specified folders and subfolders (eg Trash, Spam, etc)
 9) Feeds messages newer than specified number of days to Bayes database through spamc.exe
 10) Prunes hMailServer logs older than specified number of days
 11) Prunes local backup copies older than specified number of days
 12) Compresses the backup directory into a multi-volume 7z archive with AES-256 encryption (including header encyrption)
 13) Creates a folder at LetsUpload.io and uploads the archive files
 14) Sends email with debug log attached

# Requirements
 Working hMailServer using either internal database or MySQL
 7-zip with path in the system path
 If using SpamAssassin, must configure service for --allow-tell in order to feed spamc

# Instructions
 Create account at LetsUpload.io and create API keys
 Fill in variables in hMailServerBackupConfig.ps1
 Run hMailServerBackup.ps1 from task scheduler at 11:58 PM (time allows for properly cycling logs)
 
# Notes
 Config switch $DeleteEmptySubFolders will delete empty subfolders found within matching message pruning folders. Run hMailServerBackupPruneMessagesTEST.ps1 with $DoDelete = FALSE to see how your system will react.
 
 Config has option to call 7-zip from system path or full path to executable. I don't know why, but system path is 3-4 faster than calling the full path. 