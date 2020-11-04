# hMailServer Offsite Backup
 hMailServer backup routine that also uploads to LetsUpload.io
 
# What does it do?
 1) Stops hMailServer and (if in use) SpamAssassin Services
 2) Updates SpamAssassin (if in use)
 3) Cycles hMailServer and (if in use) SpamAssassin logs
 4) Backs up hMailServer data directory using RoboCopy
 5) Dumps MySQL database -or- internal database and adds to backup directory
 6) Backs up hMailServer.INI
 7) Restarts SpamAssassin (if in use) and hMailServer
 8) Prunes messages older than specified number of days in specified folders and subfolders (eg Trash, Spam, etc)
 9) Prunes hMailServer logs older than specified number of days
 10) Compresses the backup directory into a multi-volume 7z archive with AES-256 encryption (including header encyrption)
 11) Creates a folder at LetsUpload.io and uploads the archive files
 12) Sends email with debug log attached

# Requirements
 Working hMailServer using either internal database or MySQL
 7-zip with path in the system path

# Instructions
 Create account at LetsUpload.io and create API keys
 Fill in variables in hMailServerBackupConfig.ps1
 Run hMailServerBackup.ps1 from task scheduler at 11:58 PM (time allows for properly cycling logs)
 
# Notes
 Config switch $DeleteEmptySubFolders will delete empty subfolders found within matching message pruning folders. Use with caution. Run hMailServerBackupPruneMessagesTEST.ps1 with $DoDelete = FALSE to see how your system will react.