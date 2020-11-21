<#

.SYNOPSIS
	hMailServer Backup

.DESCRIPTION
	Configuration for hMailServer Backup

.FUNCTIONALITY
	Backs up hMailServer, compresses backup and uploads to LetsUpload

.PARAMETER 

	
.NOTES
	7-Zip required
	Run at 12:58PM from task scheduler
	
	
.EXAMPLE


#>

<###   USER VARIABLES   ###>
$VerboseConsole        = $True                  # If true, will output debug to console
$VerboseFile           = $True                  # If true, will output debug to file
$UseSA                 = $True                  # Specifies whether SpamAssassin is in use
$PruneBackups          = $True                  # If true, will delete local backups older than N days
$DaysToKeepBackups     = 5                      # Number of days to keep backups - older backups will be deleted at end of script

<###   MISCELLANEOUS BACKUP FILES   ###>        # Array of additional miscellaneous files to backup
$MiscBackupFiles   = @(
	"C:\Program Files (x86)\hMailServer\Bin\hMailServer.INI"
	"C:\Program Files (x86)\hMailServer\Events\EventHandlers.vbs"
)

<###   FOLDER LOCATIONS   ###>
$hMSDir                = "C:\Program Files (x86)\hMailServer"  # hMailServer Install Directory
$SADir                 = "C:\Program Files\JAM Software\SpamAssassin for Windows"  # SpamAssassin Install Directory
$SAConfDir             = "C:\Program Files\JAM Software\SpamAssassin for Windows\etc\spamassassin"  # SpamAssassin Conf Directory
$SevenZipDir           = "C:\Program Files\7-Zip"  # Path to 7-Zip Install Directory
$MailDataDir           = "C:\HMS-DATA"          # hMailServer Data Dir
$BackupTempDir         = "C:\HMS-BACKUP-TEMP"   # Temporary backup folder for RoboCopy to compare
$BackupLocation        = "C:\HMS-BACKUP"        # Location archive files will be stored
$BayesBackupLocation   = "C:\Bayes\bayes_backup"  # Bayes backup folder

<###   HMAILSERVER COM VARIABLES   ###>
$hMSAdminPass          = "supersecretpassword"  # hMailServer Admin password

<###   WINDOWS SERVICE VARIABLES   ###>
$hMSServiceName        = "hMailServer"          # Name of hMailServer Service (check windows services to verify exact spelling)
$SAServiceName         = "spamassassin"         # Name of SpamAssassin Service (check windows services to verify exact spelling)
$ServiceTimeout        = 5                      # number of minutes to continue trying if service start or stop commands become unresponsive

<###   PRUNE MESSAGES VARIABLES   ###>
$DoDelete              = $True                  # FOR TESTING - set to false to run and report results without deleting messages and folders
$PruneMessages         = $True                  # True will run message pruning routine
$PruneSubFolders       = $True                  # True will prune all folders in levels below name matching folders
$PruneEmptySubFolders  = $True                  # True will delete empty subfolders below the matching level unless a subfolder within contains messages
$DaysBeforeDelete      = 30                     # Number of days to keep messages in pruned folders
$PruneFolders          = "Trash|Deleted|Junk|Spam|2020-[01][0-9]-[0-3][0-9]$|Unsubscribes"  # Names of IMAP folders you want to cleanup - uses regex

<###   FEED BAYES VARIABLES   ###>
$DoSpamC               = $True                  # FOR TESTING - set to false to run and report results without feeding SpamC with spam/ham
$FeedBayes             = $True                  # True will run Bayes feeding routine
$HamFolders            = "INBOX|Ham"            # Ham folders to feed messages to spamC for bayes database - uses regex
$SpamFolders           = "Spam|Junk"            # Spam folders to feed messages to spamC for bayes database - uses regex
$BayesDays             = 7                      # Number of days worth of spam/ham to feed to bayes

<###   MySQL VARIABLES   ###>
$UseMySQL              = $True                  # Specifies whether database used is MySQL
$MySQLBINdir           = "C:\xampp\mysql\bin"   # MySQL BIN folder location
$MySQLUser             = "hmailserver"          # hMailServer database user
$MySQLPass             = "supersecretpassword"  # hMailServer database password
$MySQLPort             = 3306                   # MySQL port

<###   7-ZIP VARIABLES   ###>
$SevenZipInSystemPath  = $True                  # If false, will call full path to executable - seems to run much faster when called from the system path
$VolumeSize            = "100m"                 # Size of archive volume parts - maximum 200m recommended - valid suffixes for size units are (b|k|m|g)
$ArchivePassword       = "supersecretpassword"  # Password to 7z archive

<###   LETSUPLOAD API VARIABLES   ###>
$APIKey1               = "1QFMyGCDgCH7BKG6ZKhxmUvAl98abP4bYiJ16iJTtLYZopqycRZJpndpca6ZgByT"
$APIKey2               = "Fky8b24HpzuYhPeXmZO8m1pe6vqcxluodasRtF1C6dnShutYkpguAlJYAWd7JgiB"
$IsPublic              = 0                      # 0 = Private, 1 = Unlisted, 2 = Public in LetsUpload.io site search

<###   HMAILSERVER LOG VARIABLES   ###>
$CycleLogs             = $True                  # Cycle SpamAssassin and Event logs nightly
$PruneLogs             = $True                  # If true, will delete logs in hMailServer \Logs folder older than N days
$DaysToKeepLogs        = 10                     # Number of days to keep old hMailServer Logs

<###   EMAIL VARIABLES   ###>
$EmailFrom             = "notify@mydomain.tld"
$EmailTo               = "admin@mydomain.tld"
$Subject               = "hMailServer Nightly Backup"
$SMTPServer            = "mail.mydomain.tld"
$SMTPAuthUser          = "notify@mydomain.tld"
$SMTPAuthPass          = "supersecretpassword"
$SMTPPort              =  587
$SSL                   = $True                  # If true, will use tls connection to send email
$UseHTML               = $True                  # If true, will format and send email body as html (with color!)
$AttachDebugLog        = $True                  # If true, will attach debug log to email report - must also select $VerboseFile
$MaxAttachmentSize     = 1                      # Size in MB

<###   GMAIL VARIABLES   ###>
<#  Alternate messaging in case of hMailServer service failure  #>
<#  "Less Secure Apps" must be enabled in gmail account settings  #>
$GmailUser             = "notifier@gmail.com"
$GmailPass             = "supersecretpassword"
$GmailTo               = "1234567890@tmomail.net"