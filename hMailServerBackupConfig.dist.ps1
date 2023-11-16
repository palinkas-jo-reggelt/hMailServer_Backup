<#

.SYNOPSIS
	hMailServer Backup

.DESCRIPTION
	Configuration for hMailServer Backup

.FUNCTIONALITY
	Backs up hMailServer, compresses backup and uploads to LetsUpload.io

.PARAMETER 

	
.NOTES
	Run at 11:58PM from task scheduler in order to properly cycle log files.
	
.EXAMPLE


#>

<###   USER VARIABLES   ###>
$VerboseConsole        = $True                  # If true, will output debug to console
$VerboseFile           = $True                  # If true, will output debug to file

<###   DATA DIR BACKUP   ###>
$BackupDataDir         = $True                  # If true, will backup data dir via robocopy

<###   MISCELLANEOUS BACKUP FILES   ###>
$BackupMisc            = $True                  # True will backup misc files listed below
$MiscBackupFiles       = @(                     # Array of additional miscellaneous files to backup (use full path)
	"C:\hMailServer\Bin\hMailServer.INI"
	"C:\hMailServer\Events\EventHandlers.vbs"
	"C:\Program Files\JAM Software\SpamAssassin for Windows\etc\spamassassin\local.cf"
)

<###   FOLDER LOCATIONS   ###>
$hMSDir                = "C:\hMailServer"       # hMailServer Install Directory
$SADir                 = "C:\Program Files\JAM Software\SpamAssassin for Windows"  # SpamAssassin Install Directory
$SAConfDir             = "C:\Program Files\JAM Software\SpamAssassin for Windows\etc\spamassassin"  # SpamAssassin Conf Directory
$MailDataDir           = "C:\HMS-DATA"          # hMailServer Data Dir
$BackupTempDir         = "C:\HMS-BACKUP-TEMP"   # Temporary backup folder for RoboCopy to compare
$BackupLocation        = "C:\HMS-BACKUP"        # Location archive files will be stored
$MySQLBINdir           = "C:\xampp\mysql\bin"   # MySQL BIN folder location

<###   HMAILSERVER COM VARIABLES   ###>
$hMSAdminPass          = "supersecretpassword"  # hMailServer Admin password

<###   SPAMASSASSIN VARIABLES   ###>
$UseSA                 = $True                  # Specifies whether SpamAssassin is in use
$UseCustomRuleSets     = $True                  # Specifies whether to download and update KAM.cf
$SACustomRules         = @(                     # URLs of custom rulesets
	"https://www.pccc.com/downloads/SpamAssassin/contrib/KAM.cf"
	"https://www.pccc.com/downloads/SpamAssassin/contrib/nonKAMrules.cf"
)

<###   OPENPHISH VARIABLES   ###>               # https://hmailserver.com/forum/viewtopic.php?t=40295
$UseOpenPhish          = $True                  # Specifies whether to update OpenPhish databases - for use with Phishing plugin for SA - requires wget in the system path
$PhishFiles            = @{
	"https://data.phishtank.com/data/online-valid.csv" = "$SAConfDir\phishtank-feed.csv"
	"https://openphish.com/feed.txt" = "$SAConfDir\openphish-feed.txt"
}
	# "https://phishstats.info/phish_score.csv" = "$SAConfDir\phishstats-feed.csv" #OpenPhish is dead.

<###   WINDOWS SERVICE VARIABLES   ###>
$hMSServiceName        = "hMailServer"          # Name of hMailServer Service (check windows services to verify exact spelling)
$SAServiceName         = "SpamAssassin"         # Name of SpamAssassin Service (check windows services to verify exact spelling)
$ServiceTimeout        = 5                      # number of minutes to continue trying if service start or stop commands become unresponsive

<###   PRUNE BACKUPS VARIABLES   ###>
$PruneBackups          = $True                  # If true, will delete local backups older than N days
$DaysToKeepBackups     = 5                      # Number of days to keep backups - older backups will be deleted

<###   PRUNE MESSAGES VARIABLES   ###>
$DoDelete              = $True                  # FOR TESTING - set to FALSE to run and report results without deleting messages and folders
$PruneMessages         = $True                  # True will run message pruning routine
$PruneSubFolders       = $True                  # True will prune messages in folders levels below name matching folders
$PruneEmptySubFolders  = $True                  # True will delete empty subfolders below the matching level unless a subfolder within contains messages
$DaysBeforeDelete      = 30                     # Number of days to keep messages in pruned folders
$SkipAccountPruning    = "user@dom.com|a@b.com" # User accounts to skip - uses regex (disable with "" or $NULL)
$SkipDomainPruning     = "domain.tld|dom2.com"  # Domains to skip - uses regex (disable with "" or $NULL)
$PruneFolders          = "Trash|Deleted|Junk|Spam|Folder-[0-9]{6}|Unsubscribes"  # Names of IMAP folders you want to cleanup - uses regex

<###   FEED BAYES VARIABLES   ###>
$FeedBayes             = $True                  # True will run Bayes feeding routine
$DoSpamC               = $True                  # FOR TESTING - set to FALSE to run and report results without feeding SpamC with spam/ham
$BayesSubFolders       = $True                  # True will feed messages from subfolders within regex name matching folders
$BayesDays             = 7                      # Number of days worth of spam/ham to feed to bayes
$HamFolders            = "INBOX|Ham"            # Ham folders to feed messages to spamC for bayes database - uses regex
$SpamFolders           = "Spam|Junk"            # Spam folders to feed messages to spamC for bayes database - uses regex
$SkipAccountBayes      = "user@dom.com|a@b.com" # User accounts to skip - uses regex (disable with "" or $NULL)
$SkipDomainBayes       = "domain.tld|dom2.com"  # Domains to skip - uses regex (disable with "" or $NULL)
$SyncBayesJournal      = $True                  # True will sync bayes_journal after feeding messages to SpamC
$BackupBayesDatabase   = $True                  # True will backup the bayes database to bayes_backup - NOT insert the file in the backup/upload routine
$BayesBackupLocation   = "C:\bayes_backup"      # Bayes backup FILE

<###   MySQL VARIABLES   ###>
$BackupDB              = $True                  # Specifies whether to run BackupDatabases function (options below)(FALSE will skip)
$UseMySQL              = $True                  # Specifies whether database used is MySQL
$BackupAllMySQLDatbase = $True                  # True will backup all databases, not just hmailserver - must use ROOT user for this
$MySQLDatabase         = "hmailserver"          # MySQL database name
$MySQLUser             = "root"                 # hMailServer database user
$MySQLPass             = "supersecretpassword"  # hMailServer database password
$MySQLPort             = 3306                   # MySQL port

<###   7-ZIP VARIABLES   ###>
$UseSevenZip           = $True                  # True will compress backup files into archive
$PWProtectedArchive    = $True                  # False = no-password zip archive, True = AES-256 encrypted multi-volume 7z archive
$VolumeSize            = "100m"                 # Size of archive volume parts - maximum 200m recommended - valid suffixes for size units are (b|k|m|g)
$ArchivePassword       = "supersecretpassword"  # Password to 7z archive

<###   HMAILSERVER LOG VARIABLES   ###>
$PruneLogs             = $True                  # If true, will delete logs in hMailServer \Logs folder older than N days
$DaysToKeepLogs        = 10                     # Number of days to keep old hMailServer Logs

<###   CYCLE LOGS VARIABLES   ###>              # Array of logs to cycle - Full file path required - not limited to hmailserver log dir
$CycleLogs             = $True                  # True will cycle logs (rename with today's date)
$LogsToCycle           = @(
	"C:\hMailServer\Logs\hmailserver_events.log"
	"C:\hMailServer\Logs\spamd.log"
)

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
<#  Alternate messaging in case of hMailServer failure  #>
<#  "Less Secure Apps" must be enabled in gmail account settings  #>
$GmailUser             = "notifier@gmail.com"
$GmailPass             = "supersecretpassword"
$GmailTo               = "1234567890@tmomail.net"