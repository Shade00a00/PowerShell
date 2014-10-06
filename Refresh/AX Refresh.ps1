Param(
[String]$instanceName="",
[String]$refreshDateParam=""
)
#TDB Version 1 30/09/2014
Set-StrictMode -Version latest

if ($instanceName -eq "" -or $instanceName -eq $null)
{
    Write-host "Instance name not set, exiting"
    Exit
}

#############################
# Set your parameters below #
#############################

set-variable -name scriptLogFile -value ((get-date -Format "yyyy-MM-dd") + "$instancename refresh log.log")

if ($refreshDateParam -ne "")
    {Set-Variable -Name refreshDate -Value ([system.datetime]::Parse($refreshDateParam)) -Option Constant}
if ($refreshDateParam -eq "")
    {Set-Variable -Name refreshDate -Value (get-date).adddays(- ((get-date).DayOfWeek.value__ + 2)) -Option Constant}
#Set Refresh Date to last friday.

# Alternatively you can set your date yourself

Set-Variable -Name AXCPath -Value "\\ts8srvbell\m$\AX_Server_Configs\*.axc" -Option Constant

#Lookup tables for all AOTs for various instances. Update this in case of new instances.

Set-Variable -Name SourceAOTHash -Option Constant -Value @{
	'DEV'='\\axfileserver\AXApplications\Application\Appl\AX2009_PROD';
	'DEV2'='\\DEVERP1SRVBELL\Application\Appl\AX2009_DEV';
	'TEST'='\\axfileserver\AXApplications\Application\Appl\AX2009_PROD';
	'TRN'='\\axfileserver\AXApplications\Application\Appl\AX2009_PROD'}

Set-Variable -Name DestinationAOTHash -Option Constant -Value @{
	'DEV'="\\deverp1srvbell\d$\Program Files\Microsoft Dynamics AX\50\Application\Appl\AX2009_DEV";
	'DEV2'="\\RDBMS0SRVBELL\u$\axfileserver\Application\Appl\AX2009_DEV2";
	'TEST'="\\RDBMS0SRVBELL\u$\axfileserver\Application\Appl\AX2009_TEST";
	'TRN'="\\erp1srvmtl\AXApplication\Appl\AX2009_TRN"}
	
Set-Variable -Name AOSServerHash -Value @{
	'DEV2'='AOS0SRVBELL';
	'TEST'='AOS0SRVBELL';
	'TRN'='ERP1SRVMTL';
	'DEV'='DEVERP1SRVBELL'}
	
Set-Variable -Name SQLServerHash -Value @{
	'DEV2'='RDBMS0SRVBELL';
	'TEST'='RDBMS0SRVBELL';
	'TRN'='ERP1SRVMTL';
	'DEV'='DEVERP1SRVBELL'}

Set-Variable -Name SQLBackupPathHash -Value @{
	'DEV2'='\\DEVERP1SRVBELL\e$\MSSQL\BACKUP\DAILYBACKUPS\AX2009_DEV';
	'TEST'='\\rdbmssqlclu\dbbackups\AX2009_PROD';
	'TRN'='\\rdbmssqlclu\dbbackups\AX2009_PROD';
	'DEV'='\\rdbmssqlclu\dbbackups\AX2009_PROD'}

Set-Variable -Name AOSServiceHash -Value @{
	'DEV2'='AOS50$05';
	'TEST'='AOS50$06';
	'TRN'='AOS50$01';
	'DEV'='AOS50$01'}

Set-Variable -Name SQLDataHash -Value @{
	'DEV2'='E:\MSSQL\DATA\';
	'TEST'='E:\MSSQL\DATA\';
	'TRN'='E:\MSSQL\DATA\';
	'DEV'='E:\MSSQL\MSSQL.1\MSSQL\Data\'}

Set-Variable -Name SQLLogHash -Value @{
	'DEV2'='F:\MSSQL\LOG\';
	'TEST'='F:\MSSQL\LOG\';
	'TRN'='F:\MSSQL\Log\';
	'DEV'='E:\MSSQL\MSSQL.1\MSSQL\Data\'}

	
####################
# Functions reused #
####################

function logThis
{
    param([string] $toLog)
    $curDateTime = get-date -Format u
    
    Write-host ("log: "+$curDateTime + "|" + $toLog)
    add-content  .\logs\$scriptLogFile ($curDateTime + "|" + $toLog)
}
function sendMail
{
    Param ([string] $pBody="Generic error")
    $curDateTime = get-date -Format "dd-MMMM-yyyy HH:mm"
    $curDate = get-date -Format "dd MMMM"
	
	write-host $pBody
    
    #Creating a Mail object
    $msg = new-object Net.Mail.MailMessage
    
    #Creating SMTP server object
    $smtp = new-object Net.Mail.SmtpClient("relay.sanimax.int")
    
    #Email structure
    $msg.From = "s3tech@sanimax.com"
    $msg.To.Add("s3tech@sanimax.com")
    $msg.subject = $env:COMPUTERNAME + ": "+ $curDate + " Script email for script in: "+ (pwd).toString()
    $msg.body = $curDateTime + "`n`n" + $pBody
    
    #Sending email
    logThis "**SENDING EMAIL**"
	$msg.body = $msg.body + "`n`n" + (get-content .\logs\$scriptLogFile | out-string)
    logThis $pBody
    $smtp.Send($msg)
    
}

logThis ("Script running as user " + [Environment]::UserName + "(" +[Environment]::UserDomainName + ") on " + [Environment]::MachineName)

try {
if ($DestinationAOTHash.Keys -notcontains $instanceName) {sendMail("Instance name not defined in hash tables, exiting");Exit}

	Set-Variable -Name SourceAOT -Value ($SourceAOTHash).$instanceName -Option Constant
	Set-Variable -Name backupPath -Value ($SQLBackupPathHash).$instanceName -Option Constant
    Set-Variable -Name AOSServer -Value ($AOSServerHash).$instanceName -Option Constant
    Set-Variable -Name AOSService -Value ($AOSServiceHash).$instanceName -Option Constant
    Set-Variable -Name DestinationAOT -Value ($DestinationAOTHash).$instanceName -Option Constant
    Set-Variable -Name SQLServer -Value ($SQLServerHash).$instanceName -Option Constant
    Set-Variable -Name SQLDataPath -Value ($SQLDataHash).$instanceName -Option Constant
    Set-Variable -Name SQLLogPath -Value ($SQLLogHash).$instanceName -Option Constant
    Set-Variable -Name DBName -Value "AX2009_$instancename" -Option Constant
    Set-Variable -Name SQLScriptPath -Value ".\$instancename\SQL\*.sql" -Option Constant
    }

catch [system.exception]
{
	sendMail ("Unable to set required variables for execution.`n`n"+$_.exception.toString())
}

Set-Variable -Name backupFile -Value (gci $backupPath -Filter "*$($refreshDate.toString("yyyyMMdd"))*").FullName -Option Constant

if ($backupFile -eq "" -or $backupFile -eq $null) { sendmail("Cannot find backup") }

#################################
# FUNCTIONAL part of the script #
#################################
# Adjust AXCs #
###############

$refreshDateString=$refreshDate.tostring("dd MMMM yyyy")

try
{
	$axcFiles = Get-ChildItem $AXCPath -Force -Recurse | where {$_.Name -match "$instanceName[^s2]"}
	#clever regex is clever : prodstaging and dev2 don't get caught when asking for prod or dev.
}
catch [system.exception]
{
	sendMail ("Access denied or path incorrect on AXC file directory`n`n"+$_.exception.toString())
}

foreach ($File in $axcFiles) {
	try
	{
		$AXCContent = Get-Content -Path $File
		logThis "Read $($file.fullname)"
        #Read in each file
	}
	catch [system.exception]
	{
		sendMail ("Could not read AXC file`n`n"+$_.exception.toString())
	}
	$newAXC = @()
	#Initialize a new array

	foreach ($line in $AXCContent){
		if ($line -like "*startupmsg,Text,*"){$newAXC += "    startupmsg,Text,Welcome to AX $instancename, refreshed $refreshDateString"};
		if ($line -notlike "*startupmsg*") {$newAXC += $line}}
		# Iterate through file and replace just the startup message
	Out-File -InputObject $newAXC -FilePath $File.fullname -Force
}

################
# STOP SERVICE # 
################

try
{
	#stop AOS services
	$svc= (get-service -name $AOSService -computername $AOSServer | where {$_.status -ne "stopped"})
    if ($svc -ne $null){$svc.Stop()}
	Start-sleep -s 30
	logThis "Service stopped"
}
catch [system.exception]
{
	sendMail ("Could not stop service`n`n"+$_.exception.toString())
}

Start-sleep -s 10
####################
# Load SMO objects #
####################

try
{
    #load assemblies
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.SMO") | Out-Null
    #Need SmoExtended for backup
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.SmoExtended") | Out-Null
    [Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.ConnectionInfo") | Out-Null
    [Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.SmoEnum") | Out-Null
}
catch [system.exception]
{
    sendMail("Error loading assemblies, please make sure script is running on an SQL server or install SMO libraries")
}
logThis "Loaded assemblies"

########################
# Create backup object #
########################

if ($backupFile -eq $null)
{
	sendMail("No backup for $refreshDate in location $backupPath, check backup process or script source")
}

#we will query the database name from the backup header later
$server = New-Object ("Microsoft.SqlServer.Management.Smo.Server") $SqlServer
$server.ConnectionContext.StatementTimeout=0

try
{
	$backupDevice = New-Object("Microsoft.SqlServer.Management.Smo.BackupDeviceItem") ($backupFile, "File")
}
catch [system.exception]
{
	sendMail("Invalid backup format, check that backup is working and completed")
}

$smoRestore = new-object("Microsoft.SqlServer.Management.Smo.Restore")

logThis ("Backup object for "+ $backupFile.toupper()+" created")

#####################
# DROP OLD DATABASE #
#####################


if ($server.Databases[$dbname] -ne $null)
{
	try
    {
        $db = $server.Databases[$dbname]
		$server.ConnectionContext.ExecuteNonQuery("ALTER DATABASE $dbname SET OFFLINE WITH ROLLBACK IMMEDIATE")
		$server.ConnectionContext.ExecuteNonQuery("ALTER DATABASE $dbname SET ONLINE")
        $db.drop()
        logThis "dropped $dbname"
    }
    
    catch [system.exception]
    {
        sendMail($_.exception.tostring() + "`n`n Error dropping $dbname to single-user mode, check that database is not currently in use.")
    }
}
Else
{
	sendMail("$dbname does not exist for refresh step DROP $dbname, continuing (this is not necessarily fatal)")
}


#################
# COPY AOT OVER #
#################

try
{
	#Copy AOD from PROD1 to PROD2
	if (!(test-path ($DestinationAOT) -PathType container))
	{
		#this only happens if the folder does not exist
		mkdir ($DestinationAOT)
		sendMail("$DestinationAOT did not exist, created folder")
	}
	del ($DestinationAOT+"\*") -force -recurse
	copy ($SourceAOT +"\*") ($DestinationAOT+ "\") -Force -Recurse -Exclude Backup,Old
	
	logThis "AOD copied from $SourceAOT to $DestinationAOT"
}
catch [system.exception]
{
	sendMail ("Could copy AOD to $DestinationAOT-- Check if $SourceAOT exists `n`n"+$_.exception.toString())
}

####################
# START DB RESTORE #
####################

$smoRestore.NoRecovery = $false;
$smoRestore.ReplaceDatabase = $false;
$smoRestore.Action = "Database"
$smoRestorePercentCompleteNotification = 10;
$smoRestore.Devices.Add($backupDevice)
$smoRestore.Database=$dbname

logThis "Set database restore options"

#get database logicalname in file list from backup
$smoRestoreDetails=$smorestore.ReadFileList($server)

logThis "Read logical name from backup source $backupfile"

#specify new data and log files (mdf and ldf)
$smoRestoreFile = New-Object("Microsoft.SqlServer.Management.Smo.RelocateFile")
$smoRestoreLog = New-Object("Microsoft.SqlServer.Management.Smo.RelocateFile")

$smoRestoreFile.LogicalFileName = ($smoRestoreDetails | where {$_.physicalname -like "*mdf*"}).logicalname
$smoRestoreFile.PhysicalFileName = $SQLDataPath + $dbname + ".mdf"
$smoRestoreLog.LogicalFileName = ($smoRestoreDetails | where {$_.physicalname -like "*ldf*"}).logicalname
$smoRestoreLog.PhysicalFileName = $SQLLogPath + $dbname + ".ldf"
$smoRestore.RelocateFiles.Add($smoRestoreFile)
$smoRestore.RelocateFiles.Add($smoRestoreLog)

logThis "Added new location for $dbname to restore job"

logThis "Starting restore..."

try
{
    $smoRestore.sqlrestore($server)
}
catch [system.exception]
{
    sendMail($_.exception.toString() + "`n`nRestore failed")
}
logThis "Restore step complete"

start-sleep -s 10
$db = $server.Databases[$dbname]

start-sleep -s 10

try
{
$ScriptList = Get-ChildItem -Path $SQLScriptPath | Sort-Object -Property Name

logThis "Generated script list"

ForEach ($Script in $ScriptList) {
    $ScriptContents = Get-Content $Script
    $db.executeNonQuery($ScriptContents)
    logThis "Ran script $($script.FullName)"
	start-sleep -s 60
    }
}

catch [system.exception]
{
    sendMail("SQL Script execution failed`n`n"+$_.exception.toString())
	exit
	#really important not to restart instance if SQL scripts didn't run.
}

#################
# START SERVICE # 
#################

try
{
	#start AOS services
	(get-service -name $AOSService -computername $AOSServer).start()
	Start-sleep -s 30
	logThis "Service restarted"
}
catch [system.exception]
{
	sendMail ("Could not stop service`n`n"+$_.exception.toString())
}