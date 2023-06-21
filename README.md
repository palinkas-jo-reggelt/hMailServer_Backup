# hMailServer Offsite Backup
 hMailServer backup routine that also uploads to LetsUpload.io
 
 Discussion thread: https://hmailserver.com/forum/viewtopic.php?f=9&t=35447

# NEW
 OpenPhish database files updater. New variables added to config.  
 See topic for further info: https://hmailserver.com/forum/viewtopic.php?t=40295  
 
# What does it do?
 1) Stops hMailServer and SpamAssassin Services
 2) Updates SpamAssassin
 3) Cycles hMailServer and SpamAssassin logs
 4) Backs up hMailServer data directory using RoboCopy
 5) Dumps MySQL database -or- internal database and adds to backup directory
 6) Backs up miscellaneous files
 7) Updates OpenPhish database files
 8) Restarts SpamAssassin and hMailServer
 9) Prunes messages older than specified number of days in specified folders and subfolders (eg Trash, Spam, etc)
 10) Feeds messages newer than specified number of days to Bayes database through spamc.exe
 11) Prunes hMailServer logs older than specified number of days
 12) Prunes local backup copies older than specified number of days
 13) Compresses the backup directory into a multi-volume 7z archive with AES-256 encryption (including header encyrption)
 14) Creates a folder at LetsUpload.io and uploads the archive files
 15) Sends email with debug log attached

# Requirements
 Working hMailServer using either internal database or MySQL  
 If using SpamAssassin, must configure service for --allow-tell in order to feed spamc  
 OpenPhish update requires WGET in the system path  

# Instructions
 Create account at LetsUpload.io and create API keys  
 Fill in variables in hMailServerBackupConfig.ps1.dist and rename to hMailServerBackupConfig.ps1  
 Run hMailServerBackup.ps1 from task scheduler at 11:58 PM (time allows for properly cycling logs)  
 
# Notes
