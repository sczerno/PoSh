<# Citrix XenApp / XenDesktop Heath and Reporting 

I found a great script written by David Ott at http://www.citrixirc.com/?p=617. So I decided to add to 
it and make it something that can be used throughout the day.

Other parts of the script also originated from a script written by Jason Poyner, XenApp Farm Health Report, 
http://deptive.co.nz/xenapp-farm-health-report.

This is meant to augment Citrix Director. Director is a great tool by itself, and has great information 
in the Monitoring Database, similar to the old Resource Manager. I created this as an easy way for Citrix 
Admins and others to get a quick look of current XenDesktop and XenApp usage as well as the basic health 
of the XenApp Servers, Sessions, Latency, etc. 

This is in no way meant to be used for a deployment of more than 10 XenApp Servers or 50 XenDesktop Computers.

Some of the reports may not be useful for Dynamic Deployments that use Provisioning Services or Machine 
Creation Services. One of the environments that I manage is static hardware for the XenApp Servers and 
static VMs for XenDesktop. I am not able not to use Provisioning Services or Machine Creation Services 
because of bureaucartic reasons but the real reason is development configurations, which is basically 
because users just don't get the dynamic creation (and destruction) of VMs, and these users are developers 
which makes it even worse. Anyway, I digress...

The solution features the following:

    - Shows a table of the each XenApp Servers stats, such as average CPU and Memnory usage, sessions, 
      Server Uptime, Registration State, Maintenance State, Agent Version, and Pending Reboot. 

          - Shows if the Application or System Event Logs on the XenApp Servers has any Critical or Errors 
            Events during a configurable time span. This can also be filtered based on key phrases in the 
            messages portion of the Event. There is a text file named EventstoIgnore.txt used for excluding
            Event Descriptions. Partial entries can be used, examples provided.

          - Shows how many Services are not running, per server, which can also be configured to monitor 
            which services are important to you. There is a text file named ServicestoMonitor.txt used for
            adding names of the service(s) to monitor. Partial names are accepted, examples are provided.

    - Shows Total XenDesktop and XenApp Sessions and License Usage.

    - Shows how many session failures have occurred in the last 15 minutes as well as how many unregistered
      computers that are not in maintenance mode. Alerts will be emailed if there are more than 5 session failures 
      or 5 unregistered computers.

    - Shows a Graph for current XenDesktop Sessions, Total, Active and Disconnected.

    - Shows a Graph for current XenApp Sessions, Total, Active and Disconnected.

    - Shows a graph of the Top 10 Latency in any of the Sessions over 10ms. There is a chart for XenDesktop Sessions 
      as well as XenApp.

      I find this useful when a remote user, or sometimes a local user, calls and says everything is responding slowly 
      in their Session. The first thing I look at is the latency chart since I support remote users that connect via 
      the Internet. 

      In addition to the charts, the script will alert you if there are 5 or more sessions that are over 300ms
      latency. These number of sessions and the latency value can be changed via a variable in the User Variables section.
      
    - Includes basic reports such as XenDesktop and XenApp usage, Citrix Receiver Versions, XenDesktops 
      not used in last 60 days, all XenDesktop and XenApp Delivery Groups, and all XenApp Published 
      Applications, etc.

#>

<# Setup and Usage

The User Variables Section has values that need to be set before running the script.

Valid Arguments for the script are: FirstRun, RunReports, XDLatency, XALatency, XAInfo or Current. 

A description of each is listed below:

    FirstRun: 

        This will run all routines of the script. This will run every part of the script in order to 
        populate all required files for the HTML page to display correctly. This only needs to be ran once.

        Usage: XDandXAHealth.ps1 FirstRun

    RunReports: 

        This will run Daily Reports such as XenDesktop and XenApp Usage for Yesterday, Last 7 days, and Last 
        30 Days as well as Client Version Reports and XenDesktop Delivery Groups idle for x days. 
        This should be ran at least once per day, preferably at 12AM.

        Usage: XDandXAHealth.ps1 RunReports

    XDLatency: 

        This will collect XenDesktop Session Latency and create a graph to display in the final HTML output page. 

        REQUIREMENT: Remote Registry must be enabled on the XenApp or XenDesktop Computer in order to read the Performance 
        Counters.

        This can be ran at least once per minute. You also have the choice of capturing the data and storing it 
        in a SQL DB. This can create a very large Data Set, but can be very useful if you support a lot of Remote 
        Access users. There is a CSV file created if you do not want to use the SQL option. 

        If you do not want to ue the SQL option, simply set the "$UseCustomDB" variable to "$false".

        There is a Text File included named "LatencyTable.txt" that can be used to create a Table for storing the latency 
        data. Just create a Database named XDXALatency and run the query to create the table, Fill in the SQL Server name
        in the variable $SQLDBName

        A Basic Report is also included, Average Latency Report.rdl

        If you do not want to collect latency counters at all, simply set the "$CollectLatency" variable to 
        "$false".

        Usage: XDandXAHealth.ps1 XDLatency

    XALatency: 

        This will collect XenApp Session Latency and create a graph to display in the final HTML output page. 

        REQUIREMENT: Remote Registry must be enabled on the XenApp or XenDesktop Computer in order to read the Performance 
        Counters.

        This can be ran at least once per minute. You also have the choice of capturing the data and storing it 
        in a SQL DB. This can create a very large Data Set, but can be very useful if you support a lot of Remote 
        Access users. There is a CSV file created if you do not want to use the SQL option. 

        If you do not want to ue the SQL option, simply set the "$UseCustomDB" variable to "$false".

        If you do not want to collect latency counters at all, simply set the "$CollectLatency" variable to 
        "$false".

        Usage: XDandXAHealth.ps1 XALatency

    Current: 

        This will gather all current XenApp and XenDesktop Sessions and creates the HTML output page.

        This can be ran has often as you like as it just polls the DDC and updates two graphs, four text fields
        and creates the HTML. These routines will always run, regardless of which argument you choose. 
        
        It is here if you just want to update the current sessions graphs more fequently.

        Usage: XDandXAHealth.ps1 Current

    XAInfo:

        This will gather all of the XenApp Servers metadata, such as CPU Usage, Memory Usage, Active and 
        Disconnected Sessions, Critical or Error Events in the Event Logs, etc.

        This could be ran at least once every 10 minutes, for small deployments.

        If you have a large deployment of XenApp, your may need more time for this portion of the script to 
        complete. You probably should be using a commercial product for larger deployments.

        If more time is needed, update the User Variable named, $EventLogCheck, to match your Scheduled Task 
        Interval. This is how far back the script looks for Critical or Errors in the Event Logs before it
        displays an error in the HTML.

        You also have the ability to filter out Event Logs to not show certain errors that are annoying, occur 
        often, and are not really anything to be concerned with. For example, my security team scans all 
        computers every two hours for weak ciphers and there is an Error Event created when the cipher is 
        unavailable. You need to create a Text File named "EventstoIgnore.txt" in the same folder as the script. 
        You can enter partial or the complete message of the Event on each line of the text
        file. A sample file has been provided.

        You also have the option of monitoring certain Services. The Text File Named "ServicestoMonitor.txt,
        located in the same folder as the script is used to list the "Display Name" of the services you want 
        to monitor. Partial or Complete names can be added to the text file, one per line. A sample file has 
        been provided.

        Usage: XDandXAHealth.ps1 XAInfo


#>

param ( [string]$paramarg )
cls

#region User Variables - Change these accordingly

$DDCName = "DDC.domain.local" ### Change this to one of your Desktop Delivery Controllers

$SQLServer = "SQLServer.domain.local" #### Change this to your sql server where the Citrix Monitoring Database is hosted - if it is an instance use: servername\instance 

$SQLDBName = "CitrixXA7Monitoring" #### Change this to your Citrix Monitor Database name containing the Citrix monitor data tables, this is installed by Director, your name will most likely be different.

$UseCustomDB = $true ### $true Indicates you want to use the Custom DB option, if you do not want to save the latency data to a Database, set this to $false and only the CSV file is created

$CollectLatency = $true ### $true Indicates you want to collect Latency counters and diplay a graph of the Top 10 Latency over 10 ms

$DatabaseName = "XDXALatency" #### Your Custom Database name containing the latency table, you will need to create this manually, and only if you choose to save the Latency data to a SQL DB

$HTMLServer = "C:\inetpub\wwwroot\Monitoring\" ### This is where the script will copy the HTML and other files to, DON'T FORGET THE TRAILING BACKSLASH "\"

#$HTMLServer = "\\WebServer\c$\inetpub\wwwroot\Monitoring\"  ### You can also use a UNC path for the $HTMLServer variable if you want, DON'T FORGET THE TRAILING BACKSLASH "\"

$HTMLFilename = "default.htm" ### File name for the file that gets created and copied to web server folder

$LogoImage = "CitrixLogoSmall.png" ### logo file used for the HTML Page replace with your own if you want

$favicon="health.ico" ### Browser page icon, replace with your own if you want

$MonitorName = "Your Citrix Deployment Name" ### The name of the Monitor Page, used in emails and the web page

$DirectorURL = "https://Director.domain.local/Director/" ## URL of your Citrix Director

$SSRSURL = "https://SSRS.domain.local/reports/report/XenDesktop%20Users%20Average%20Latency" ### SQL Server Reporting Server Latency Report URL

[string]$smtpServer = "SMTP.domain.local" ### Your SMTP Server

[string]$emailFrom = "CitrixMonitor@domain.local" ### The From Email Address for errors

[string]$emailTo = "youremailaddress@domain.local" ### The To Email Address for errors

[int]$FailedConnMinutestoCheck = "15" ## Time span to show failed connections, in minutes

[int]$MStoBeConsideredLatent = "300"  ## Latency in MS to be considered an issue.

[int]$NumOfLatencytoCauseAlerts = "5"  ## Number of Sessions to be considered a problem if they are equal to or over the $MStoBeConsideredLatent value

[int]$LatencyMinutestoCheck = "15" ## Time span to show failed connections, in minutes

$EventLogCheck = "15" ### How far back to look for Critical or Error events in the Event Logs, in minutes, match your scheduled task for the Current argument at a minimum

$UpdateInterval = "15" ### If the HTML page has not been updated in x minutes, there will be an alert on the HTML page to let you know the script is not running correctly, match your scheduled task for the Current argument at a minimum

$DesktopOSTypes = "Windows 7","Windows 8","Windows 10","Red Hat Enterprise Linux" ### OS Names of the Workstations used in your XenDesktop deployment, see notes below to determine OSTypes.

$ServerOSTypes = "Windows 2012 R2","Windows 2016","Windows 2019" ### OS Names of the Windows Servers used in your XenApp deployment, see notes below to determine OSTypes.

<# 
To get a list of OSTypes, run the following:

$OS = get-brokerdesktop -AdminAddress <Delivery Controller> | select OSType 
$OS.OSType | select -Unique

The script uses 'match' to compare the names so it only has to be a partial name of the OS, for example, it may return "Windows 7 Service Pack 1" but "Windows 7" is enough information.

#>

#endregion User Variables

################## You should not have to change anything below this line, but help yourself ##################

#region Script Variables

$FailedConnectionTimeSpanMinutes = [DateTime]::Now - [TimeSpan]::FromMinutes($FailedConnMinutestoCheck) 
$currentDir = Split-Path $MyInvocation.MyCommand.Path
$ScriptFilePath = $currentDir+"\"
$LogFilesPath = $ScriptFilePath+ "Logs\"
If(!(test-path $LogFilesPath)) { New-Item -ItemType Directory -Force -Path $LogFilesPath }
$LogFileName = "XDandXALog.log"
$LogFile=$LogFilesPath+"$LogFileName"
$LogFileEmailName = "EmailLastSent.log"
$LogEmailFile = $LogFilesPath+"$LogFileEmailName"
$LogFileLatencyXDEmailName = "EmailLastSentforXDLatency.log"
$LogXDLatencyEmailFile = $LogFilesPath+"$LogFileLatencyXDEmailName"
$LogFileLatencyXAEmailName = "EmailLastSentforXALatency.log"
$LogXALatencyEmailFile = $LogFilesPath+"$LogFileLatencyXAEmailName"
$LogErrorFileName = "XDandXAErrorLog.log"
$LogErrorFile=$LogFilesPath+"$LogErrorFileName"
$LatencyLogFileName = "XDandXALatencyLog.log"
$LatencyErrorLogFileName = "XDandXALatencyErrorLog.log"
$LatencyLogFile=$LogFilesPath+"$LatencyLogFileName"
$LatencyErrorLogFile=$LogFilesPath+"$LatencyErrorLogFileName"
$PreviousLatencyLogFile = $LogFilesPath+"XDandXALatencyLog_PreviuosRun.log"
$AlertsEmailed = $LogFilesPath+"AlertsEmailed.log"
$CurrentAlerts = $LogFilesPath+"AlertsCurrent.log"
$AlertEmail = $LogFilesPath+"AlertsEmailTimeStamp.log"
$ErrorStyle = "style=""background-color: #000000; color: #FF3300;"""
$LatencyLogFile=$LogFilesPath+"$LatencyLogFileName"
$fav="favicon.ico"
$Latencyfile = $ScriptFilePath+"LatencyHistory.csv"
$HTMLFilePath = $ScriptFilePath+ "HTML\"
If(!(test-path $HTMLFilePath)) { New-Item -ItemType Directory -Force -Path $HTMLFilePath }
$ScriptGraphs = $ScriptFilePath+ "Graphs\"
If(!(test-path $ScriptGraphs)) { New-Item -ItemType Directory -Force -Path $ScriptGraphs }
$ScriptCurrent = $ScriptFilePath+ "Current\"
If(!(test-path $ScriptCurrent)) { New-Item -ItemType Directory -Force -Path $ScriptCurrent }
$ScriptReports = $ScriptFilePath+ "Reports\"
If(!(test-path $ScriptReports)) { New-Item -ItemType Directory -Force -Path $ScriptReports }
$ScriptEvents = $ScriptFilePath+ "Events\"
If(!(test-path $ScriptEvents)) { New-Item -ItemType Directory -Force -Path $ScriptEvents }
$HTMLImages = $HTMLServer+ "Images\"
If(!(test-path $HTMLImages)) { New-Item -ItemType Directory -Force -Path $HTMLImages }
$HTMLGraphs = $HTMLServer+ "Graphs\"
If(!(test-path $HTMLGraphs)) { New-Item -ItemType Directory -Force -Path $HTMLGraphs }
$HTMLCurrent = $HTMLServer+ "Current\"
If(!(test-path $HTMLCurrent)) { New-Item -ItemType Directory -Force -Path $HTMLCurrent }
$HTMLReports = $HTMLServer+ "Reports\"
If(!(test-path $HTMLReports)) { New-Item -ItemType Directory -Force -Path $HTMLReports }
$HTMLEvents = $HTMLServer+ "Events\"
If(!(test-path $HTMLEvents)) { New-Item -ItemType Directory -Force -Path $HTMLEvents }
$ScriptServices = $ScriptFilePath+ "Services\"
If(!(test-path $ScriptServices)) { New-Item -ItemType Directory -Force -Path $ScriptServices }
$HTMLServices = $HTMLServer+ "Services\"
If(!(test-path $HTMLServices)) { New-Item -ItemType Directory -Force -Path $HTMLServices }
$DesktopGraphFilename = "XDLatencyGraph.png"
$DesktopGraphFile = $ScriptGraphs+"$DesktopGraphFilename"
$ServerGraphFilename = "XALatencyGraph.png"
$ServerGraphFile = $ScriptGraphs+"$ServerGraphFilename"
$XDUsageGraphFilename = "XDUsageGraph.png"
$XDUsageGraphFile = $ScriptGraphs+$XDUsageGraphFilename
$XAUsageGraphFilename = "XAUsageGraph.png"
$XAUsageGraphFile = $ScriptGraphs+$XAUsageGraphFilename
$yesterdayfilename = $ScriptReports+"yesterday.html"
$last7daysfilename = $ScriptReports+"last7days.html"
$last30daysfilename = $ScriptReports+"last30days.html"
$XDyesterdayfilename = $ScriptReports+"XDyesterday.html"
$XDlast7daysfilename = $ScriptReports+"XDlast7days.html"
$XDlast30daysfilename = $ScriptReports+"XDlast30days.html"
$XAyesterdayfilename = $ScriptReports+"XAyesterday.html"
$XAlast7daysfilename = $ScriptReports+"XAlast7days.html"
$XAlast30daysfilename = $ScriptReports+"XAlast30days.html"
$XAPAyesterdayfilename = $ScriptReports+"XAPAyesterday.html"
$XAPAlast7daysfilename = $ScriptReports+"XAPAlast7days.html"
$XAPAlast30daysfilename = $ScriptReports+"XAPAlast30days.html"
$XDCurrentFileName = $ScriptCurrent+"XDCurrent.html"
$XACurrentFileName = $ScriptCurrent+"XACurrent.html"
$XAInfoFileName = $ScriptReports+"XAInfo.html"
$clientversiondetailslast30filename = $ScriptReports+"clientverdetails30days.html"
$clientversiondetailslast7filename = $ScriptReports+"clientverdetails7days.html"
$clientversions30daysfilename = $ScriptReports+"clientver30days.html"
$clientversions7daysfilename = $ScriptReports+"clientver7days.html"
$Script:date = (Get-Date -format g)
[datetime]$24hrFormat = ([datetime]$Script:date)
$date24 = $24hrFormat.ToString("MM/dd/yyyy HH:mm")
$ecg = "ecg_wide.png"
$EventstoIgnoreFile = $ScriptFilePath+"EventstoIgnore.txt"
$EventstoIgnoreGC = gc $EventstoIgnoreFile | ? {$_.trim() -ne "" }
$EventstoIgnore = ($EventstoIgnoreGC –join “|”)
$ServicestoMonitorGC= $ScriptFilePath+"ServicestoMonitor.txt"
$ServicestoMonitorFile = gc $ServicestoMonitorGC | ? {$_.trim() -ne "" }
$ServicestoMonitor = ($ServicestoMonitorFile -join "|")
$XALatencyTimeStampFile = $ScriptFilePath+"XALatencyTimeStamp.txt"
$XDLatencyTimeStampFile = $ScriptFilePath+"XDLatencyTimeStamp.txt"
$XAInfoTimeStampFile = $ScriptFilePath+"XAInfoTimeStamp.txt"
$StdHeader = @" 
<head>
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<style>
body {background-color: #a5a5a5;}
TABLE {margin-left: auto; margin-right: auto; border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;width: 95%} 
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #636363;color: #ffffff;} 
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;text-align: center;} 
H3 {text-align: center;}
.odd { background-color:#ffffff; } 
.even { background-color:#dddddd; } 
</style>
</head>
"@
$CountDownHeader = @" 
<head>
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
<meta http-equiv="refresh" content="60">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<style>
body {background-color: #a5a5a5;}
TABLE {margin-left: auto; margin-right: auto; border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;width: 95%} 
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #636363;color: #ffffff;} 
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;text-align: center;} 
H3 {text-align: center;}
.odd { background-color:#ffffff; } 
.even { background-color:#dddddd; } 
</style>
</head>
"@
$EVHeader = @" 
<head>
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<style>
body {background-color: #a5a5a5;}
TABLE {margin-left: auto; margin-right: auto; border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;width: 95%} 
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #8e0000; color: #ffffff;} 
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;} 
H3 {text-align: center;}
.odd { background-color:#ffffff; } 
.even { background-color:#dddddd; } 
</style>
</head>
"@

#endregion Script Variables

#region Script Functions

Function LoadCitrixSnapin {
Add-PSSnapin Citrix* -ErrorAction Continue
$PSS = Get-PSSnapin | ? {$_.Name -match "Citrix"}

if ($PSS -eq $null) {
    "Error loading Citrix Powershell snapin" | LogMe -error 
    exit
    }
}

Function ArgHelp {
cls
$ArgHelp = @"
You must provide a valid argument to the script.

Valid Arguments are: FirstRun, RunReports, Latency, XAInfo or Current. 

A description of each is listed below.

FirstRun: 
    This will run all routines of the script. This will run every part of the script in order to 
    populate all required files for the HTML page to display correctly. This only needs to be ran once.

    Usage: XDandXAHealth.ps1 FirstRun

RunReports: 
    This will run Daily Reports such as XenDesktop and XenApp Usage for Yesterday, Last 7 days, and Last 
    30 Days as well as Client Version Reports and XenDesktop Delivery Groups idle for x days. 
    This should be ran at least once per day, preferably at 12AM.

    Usage: XDandXAHealth.ps1 RunReports

XDLatency: 
    This will collect XenDesktop Sessions Latency and create a graph to display in the final HTML output 
    page. This can be ran at least once per minute. You also have the choice of capturing the data and 
    storing it in a SQL DB. This can create a very large Data Set, but can be very useful if you support 
    a lot of Remote Access users. 

    Usage: XDandXAHealth.ps1 XDLatency

XALatency: 
    This will collect XenApp Sessions Latency and create a graph to display in the final HTML output page. 
    This can be ran at least once per minute. You also have the choice of capturing the data and storing it 
    in a SQL DB. This can create a very large Data Set, but can be very useful if you support a lot of Remote 
    Access users. 

    Usage: XDandXAHealth.ps1 XALatency


Current: 
    This will gather all current XenApp and XenDesktop Sessions and creates the HTML output page.
    This should be ran at least once every 5 minutes. 

    Usage: XDandXAHealth.ps1 Current

XAInfo:
    This will gather all of the XenApp Servers metadata, such as CPU Usage, Memory Usage, Active and 
    Disconnected Sessions, Critical or Error Events in the Event Logs, etc.
    This should be ran at least once every 10 minutes.
    If you have a large deployment of either XenApp, your deploy may need more time for the script to complete.
    If more time is needed, update the User Variable named, `$EventLogCheck, to match your Scheduled Task 
    Interval.

    Usage: XDandXAHealth.ps1 XAInfo


"@

Write-Warning $ArgHelp
}

Function CheckCpuUsage($server){
    Try { $CpuUsage=(get-counter -ComputerName $server -Counter "\Processor(_Total)\% Processor Time" -SampleInterval 1 -MaxSamples 5 -ErrorAction Stop | select -ExpandProperty countersamples | select -ExpandProperty cookedvalue | Measure-Object -Average).average
        $CpuUsage = "{0:N1}" -f $CpuUsage; return $CpuUsage
        } 
    Catch { $CpuUsage = "Warn N/A"; return $CpuUsage } 
}

Function CheckMemoryUsage($server){
    Try 
	{   $SystemInfo = (Get-WmiObject -computername $Server -Class Win32_OperatingSystem -ErrorAction Stop | Select-Object TotalVisibleMemorySize, FreePhysicalMemory)
    	$TotalRAM = $SystemInfo.TotalVisibleMemorySize/1MB 
    	$FreeRAM = $SystemInfo.FreePhysicalMemory/1MB 
    	$UsedRAM = $TotalRAM - $FreeRAM 
    	$RAMPercentUsed = ($UsedRAM / $TotalRAM) * 100 
    	$RAMPercentUsed = "{0:N2}" -f $RAMPercentUsed
    	return $RAMPercentUsed
    } Catch { $RAMPercentUsed = "Warn N/A"; return $RAMPercentUsed } 
}

Function get-freespace($server){    
    try {
        $Disks = gwmi -computername $Server win32_logicaldisk -filter "drivetype=3" -ErrorAction SilentlyContinue
        foreach ($Disk in $Disks){
            $freeGB=[Math]::round((($disk.freespace/$disk.size) * 100))
            $Drive = $Disk.deviceid 
            if(![bool]$freeGB){$freeGB=0}
            if ($freeGB -le 10){$freespace = "Critical $Drive $freeGB%";return $freespace}
            elseif ($freeGB -le 15){$freespace = "Warning $Drive $freeGB%";return $freespace}
            }
        }
    catch {$freespace = "Warn N/A";return $freespace}
   $freespace = "OK";return $freespace    
}

Function get-pendingreboot($server){
    $HKLM = 2147483650
    $pending="No"
    $reg = gwmi -List -Namespace root\default -ComputerName $server | Where-Object {$_.Name -eq "StdRegProv"}
	    if($reg.Enumkey($HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update").snames -contains "RebootRequired"){$pending="Yes"}
	    elseif($reg.Enumkey($HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing").snames -contains "RebootPending"){$pending="Yes"}
	    elseif($reg.GetStringValue($HKLM,"SYSTEM\CurrentControlSet\Control\Session Manager","PendingFileRenameOperations").sValue){$pending="Yes"}
	    elseif($reg.GetStringValue($HKLM,"SOFTWARE\Wow6432Node\Sophos\AutoUpdate\UpdateStatus\VolatileFlags","RebootRequired").sValue){$pending="Yes"}
    return $pending
}

Function Get-EventLogCount($server){
    $EventLogs = Get-WinEvent -FilterHashtable @{Logname='System', 'Application','application';Level=1,2;StartTime=[datetime]::Now.AddMinutes(-$EventLogCheck)} -ErrorAction SilentlyContinue -ComputerName $Server | where {$_.Message -notmatch $EventstoIgnore}
    $Script:NumofEventsCount = ($EventLogs|measure-object).count
    If ($Script:NumofEventsCount -gt 0) {$NumofEvents="Red<a title=`"Click here for the Event Log entries for $Server`" href=`"./events/EventLog$Server.html`" onclick=`"window.open('./events/EventLog$Server.html', 'newwindow', 'width=1000,height=600,top=100,left=100'); return false;`"><div style=`"height:100%;width:100%`">Red"+$Script:NumofEventsCount+"</div></a>"}
    else {$NumofEvents=0}
    return $NumofEvents
}

Function Get-EventLogs($server){
    $EventLogs = Get-WinEvent -FilterHashtable @{Logname='System', 'Application','application';Level=1,2;StartTime=[datetime]::Now.AddMinutes(-$EventLogCheck)} -ErrorAction SilentlyContinue -ComputerName $Server | where {$_.Message -notmatch $EventstoIgnore}
    $EventLogMessage = $EventLogs
    return $EventLogMessage
}

Function Get-Uptime ($Server){
    $os = Get-WmiObject win32_operatingsystem -ComputerName $Server -ErrorAction SilentlyContinue
    $uptime = (Get-Date) - $os.ConvertToDateTime($os.LastBootUpTime)
    $getuptime = $uptime.Days
    return $getuptime
}

Function CheckService($Server){
    $Script:ServiceResult = @()
    $Script:ServiceResult += Get-Service -computername $Server | ?{($_.DisplayName -match $ServicestoMonitor) -and ($_.starttype -eq "Automatic") -and ($_.Status -ne "Running")} | select name,starttype,status,displayname
    $Script:ServiceCount = $($Script:ServiceResult | measure).Count
    if ($Script:ServiceCount -gt 0) {$Script:ServiceStatus="Red<a title=`"Click here for the Event Log entries for $Server`" href=`"./Services/Services$Server.html`" onclick=`"window.open('./Services/Services$Server.html', 'newwindow', 'width=1000,height=600,top=100,left=100'); return false;`"><div style=`"height:100%;width:100%`">Red"+$Script:ServiceCount+"</div></a>"}
    else {$Script:ServiceStatus=0}
    $svccurrfile = $null  
    $svccurrfile = "<title>Services not running on $Server</title>"
    $svccurrfile = $svccurrfile + $EVHeader
    $svccurrfile = $svccurrfile + "<body><html><h3>Services not running on $Server as of $(get-date)</h3>"
    $svccurrfile = $svccurrfile + ($Script:ServiceResult | ConvertTo-Html -Fragment name,starttype,status,displayname | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd)
    $svccurrfile = $svccurrfile + "</html></body>"
    $ServiceFile = "Services$Server.html"
    $svccurrfile | out-file -FilePath "$ScriptServices$ServiceFile"
    copy-item $ScriptServices$ServiceFile $HTMLServices$ServiceFile -force
}

Function get-licenses{
    $Script:citrixlicenses=""
    $Script:citrixlicenses=get-brokersite -AdminAddress $DDCName | select -expand licensedsessionsactive 
    }

Function LogMeLatency() {
    #===================================================================== #
    # Sends the results into a logfile as well as in the powershell window #
    #===================================================================== #
    Param( [parameter(Mandatory = $true, ValueFromPipeline = $true)] $logEntry,
	   [switch]$displaygreen,
	   [switch]$error,
	   [switch]$warning,
	   [switch]$displaynormal,
       [switch]$displayscriptstartend
	   )
    
    $Status = $logEntry
    if($error) { 
        $script:ErrorMessage += "<p $ErrorStyle>$logEntry<p>"
        Write-Host "$logEntry" -Foregroundcolor Red; $logEntry = [DateTime]::Now.ToString("[MM/dd/yyy HH:mm:ss.fff]: ") + "[ERROR] $logEntry"
        $logEntry | Out-File -FilePath $LatencyErrorLogFile -Append
        }
	elseif($warning) { Write-Host "$logEntry" -Foregroundcolor Yellow; $logEntry = [DateTime]::Now.ToString("[MM/dd/yyy HH:mm:ss.fff]: ") + "[WARNING] $logEntry"}
	elseif ($displaynormal) { Write-Host "$logEntry" -Foregroundcolor White; $logEntry = [DateTime]::Now.ToString("[MM/dd/yyy HH:mm:ss.fff]: ") + "[INFO] $logEntry" }
	elseif($displaygreen) { Write-Host "$logEntry" -Foregroundcolor Green; $logEntry = [DateTime]::Now.ToString("[MM/dd/yyy HH:mm:ss.fff]: ") + "[SUCCESS] $logEntry" }
    elseif($displayscriptstartend) { Write-Host "$logEntry" -Foregroundcolor Magenta; $logEntry = "[SCRIPT_STARTEND] $logEntry" }
    else { Write-Host "$logEntry"; $logEntry = "$logEntry" }
    if ($logEntry -ne $null) {$logEntry | Out-File -FilePath $LatencyLogFile -Append}

}

Function LogMe() {
    #===================================================================== #
    # Sends the results into a logfile as well as in the powershell window #
    #===================================================================== #
    Param( [parameter(Mandatory = $true, ValueFromPipeline = $true)] $logEntry,
	   [switch]$displaygreen,
	   [switch]$error,
	   [switch]$warning,
	   [switch]$displaynormal,
       [switch]$displayscriptstartend
	   )
    
    $Status = $logEntry
    if($error) { 
        $script:ErrorMessage += "<p $ErrorStyle>$logEntry<p>"
        Write-Host "$logEntry" -Foregroundcolor Red; $logEntry = [DateTime]::Now.ToString("[MM/dd/yyy HH:mm:ss.fff]: ") + "[ERROR] $logEntry"
        $logEntry | Out-File -FilePath $LogErrorFile -Append
        }
	elseif($warning) { Write-Host "$logEntry" -Foregroundcolor Yellow; $logEntry = [DateTime]::Now.ToString("[MM/dd/yyy HH:mm:ss.fff]: ") + "[WARNING] $logEntry"}
	elseif ($displaynormal) { Write-Host "$logEntry" -Foregroundcolor White; $logEntry = [DateTime]::Now.ToString("[MM/dd/yyy HH:mm:ss.fff]: ") + "[INFO] $logEntry" }
	elseif($displaygreen) { Write-Host "$logEntry" -Foregroundcolor Green; $logEntry = [DateTime]::Now.ToString("[MM/dd/yyy HH:mm:ss.fff]: ") + "[SUCCESS] $logEntry" }
    elseif($displayscriptstartend) { Write-Host "$logEntry" -Foregroundcolor Magenta; $logEntry = "[SCRIPT_STARTEND] $logEntry" }
    else { Write-Host "$logEntry"; $logEntry = "$logEntry" }
    if ($logEntry -ne $null) {$logEntry | Out-File -FilePath $LogFile -Append}

}

Function CreatePreviousLatencyRunLog {
    # Creating Log file for previuos ran jobs
    # File is cleared when it reaches 1MB
    if (test-path $PreviousLatencyLogFile) { 
        $size = Get-ChildItem $PreviousLatencyLogFile
        if ($size.length -ge 1048576) {clear-content $PreviousLatencyLogFile}
        }
    if (test-path $LatencyErrorLogFile) {
        $size = Get-ChildItem $LatencyErrorLogFile
        if ($size.length -ge 1048576) {clear-content $LatencyErrorLogFile}
        }
    if (test-path $LogFile) {        
        $size = Get-ChildItem $LogFile
        if ($size.length -ge 1048576) {clear-content $LogFile}
        }
    if (test-path $LogErrorFile) {
        $size = Get-ChildItem $LogErrorFile
        if ($size.length -ge 1048576) {clear-content $LogErrorFile}
        }
    if (Test-Path $LatencyLogFile) {
        $g = Get-Content $LatencyLogFile 
        $g | out-file $PreviousLatencyLogFile -append
        }
    if (test-path $LatencyLogFile) { rm -path $LatencyLogFile -force }
}

Function Error-Message {
    “Caught an exception:” | LogMeLatency -error
    “Exception Type: $($_.Exception.GetType().FullName)” | LogMeLatency -error
    “Exception Message: $($_.Exception.Message)” | LogMeLatency -error
    $Script:CurrentErrors += @($($_.Exception.Message))          
}

Function RunQuery {
    param ($CommandText)
        try
        {
        $Write2DB = $SQLConnection.CreateCommand()
        $Write2DB.CommandText = $CommandText
        $Write2DB.ExecuteNonQuery()>$null
        }
        catch [System.Data.SqlClient.SqlException] 
            {
            $error[0]|format-list -force | out-string | LogMeLatency -error
            $SQLConnection.Close()
                if ($SQLConnection.State -eq "Closed") { "Succesfully Closed connection to SQL Server..."; EXIT } 
                else  { "Failed to close Connection to SQL Server..." | LogMeLatency -error; EXIT }
            }
        catch 
            { "An error occurred while attempting to open the database connection and execute a command." | LogMeLatency -error
            $SQLConnection.Close()
                if ($SQLConnection.State -eq "Closed") { "Succesfully Closed connection to SQL Server..."; EXIT } 
                else { "Failed to close Connection to SQL Server..." | LogMeLatency -error; EXIT }
            } 
    $CommandText = $null
}

function sqlquery ($q) {
    $SqlQuery = $q
    $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
    $SqlCmd.CommandText = $SqlQuery
    $SqlCmd.Connection = $SqlConnection
    $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $SqlAdapter.SelectCommand = $SqlCmd
    $DataSet = New-Object System.Data.DataSet
    $SqlAdapter.Fill($DataSet)
    return $DataSet.tables[0]
}

Function Set-AlternatingRows {
    [CmdletBinding()]
   	Param(
       	[Parameter(Mandatory,ValueFromPipeline)]
        [string]$Line,
   	    [Parameter(Mandatory)]
       	[string]$CSSEvenClass,
        [Parameter(Mandatory)]
   	    [string]$CSSOddClass
   	)
	Begin {$ClassName = $CSSEvenClass}
	Process {
		If ($Line.Contains("<tr><td>")) {
            $Line = $Line.Replace("<tr>","<tr class=""$ClassName"">")
			If ($ClassName -eq $CSSEvenClass){$ClassName = $CSSOddClass}
			Else {$ClassName = $CSSEvenClass}
		}
		Return $Line
	}
}

Function GetIdleDesktops{
    $logonDate = (get-date).AddDays(-60)
    [System.Collections.ArrayList]$Script:xdidle = @()
    $Desktops = Get-BrokerDesktop -maxrecordcount 5000 -adminaddress $DDCName | where {$_.OSType -match ( $DesktopOSTypes -join '|')} | Select-Object DesktopGroupName, MachineName, LastConnectionTime | where-object {$_.LastConnectionTime -le $logonDate} | Sort-Object LastConnectionTime
    foreach ($u in $Desktops){
        $IncUsers = Get-BrokerAccessPolicyRule -DesktopGroupName $u.DesktopGroupName | Select-Object -ExpandProperty IncludedUsers 
        $Users=$IncUsers.Name  | select -Unique
        $desktopgroupname = $u.DesktopGroupName
        $machinename = $u.MachineName
        $lastconnectime = $u.LastConnectionTime
        $xdidleinforow = @("1")
        $Script:xdidle.add(($xdidleinforow | select @{n='Desktop Group Name';e={$desktopgroupname}},@{n='Machine Name';e={"$machinename"}},@{n='Last Connection Time';e={"$lastconnectime"}})) | Out-Null
        }
    $xdidlecurrfile = $null  
    $xdidlecurrfile = $xdidlecurrfile + "<title>XenDesktops that have not been used in 60 Days or more as of $Script:date</title>"
    $xdidlecurrfile = $xdidlecurrfile + $StdHeader
    $xdidlecurrfile = $xdidlecurrfile + "<body><html>"
    $xdidlecurrfile = $xdidlecurrfile + "<h3>XenDesktops that have not been used in 60 Days as of $Script:date</h3>"
    $xdidlecurrfile = $xdidlecurrfile + ($Script:xdidle | ConvertTo-Html -Fragment | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd)
    $xdidlecurrfile = $xdidlecurrfile + "</html></body>"
    $xdidlefile = "XDIdle.html"
    $xdidlecurrfile | out-file -FilePath $ScriptReports$xdidlefile
    try {copy-item $ScriptReports$xdidlefile $HTMLReports -force}
    catch { "Error Copying $ScriptReports$xdidlefile to $HTMLReports" | LogMe -error
        $_.Exception.Message | LogMe -error}
}

Function GetIdleServers{
    $logonDate = (get-date).AddDays(-60)
    [System.Collections.ArrayList]$Script:xdidle = @()
    $Desktops = Get-BrokerDesktop -maxrecordcount 5000 -adminaddress $DDCName | where {$_.OSType -match ( $ServerOSTypes -join '|')} | Select-Object DesktopGroupName, MachineName, LastConnectionTime | where-object {$_.LastConnectionTime -le $logonDate} | Sort-Object LastConnectionTime
    foreach ($u in $Desktops){
        $IncUsers = Get-BrokerAccessPolicyRule -DesktopGroupName $u.DesktopGroupName | Select-Object -ExpandProperty IncludedUsers 
        $Users=$IncUsers.Name  | select -Unique
        $desktopgroupname = $u.DesktopGroupName
        $machinename = $u.MachineName
        $lastconnectime = $u.LastConnectionTime
        $xdidleinforow = @("1")
        $Script:xdidle.add(($xdidleinforow | select @{n='Desktop Group Name';e={$desktopgroupname}},@{n='Machine Name';e={"$machinename"}},@{n='Last Connection Time';e={"$lastconnectime"}})) | Out-Null
        }
    $xaidlecurrfile = $null  
    $xaidlecurrfile = $xaidlecurrfile + "<title>XenApp Servers that have not been used in 60 Days or more as of $Script:date</title>"
    $xaidlecurrfile = $xaidlecurrfile + $StdHeader
    $xaidlecurrfile = $xaidlecurrfile + "<body><html>"
    $xaidlecurrfile = $xaidlecurrfile + "<h3>XenApp Servers that have not been used in 60 Days as of $Script:date</h3>"
    $xaidlecurrfile = $xaidlecurrfile + ($Script:xdidle | ConvertTo-Html -Fragment | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd)
    $xaidlecurrfile = $xaidlecurrfile + "</html></body>"
    $xaidlefile = "XAIdle.html"
    $xaidlecurrfile | out-file -FilePath $ScriptReports$xaidlefile
    try {copy-item $ScriptReports$xaidlefile $HTMLReports -force}
    catch { "Error Copying $ScriptReports$xaidlefile to $HTMLReports" | LogMe -error
        $_.Exception.Message | LogMe -error}
}

Function GetXenDesktopDGs{
    [System.Collections.ArrayList]$Script:xdddgs = @()
    $Desktops = Get-BrokerDesktop -maxrecordcount 5000 -adminaddress $DDCName | where {$_.OSType -match ( $DesktopOSTypes -join '|')} | Select-Object DesktopGroupName, MachineName, OSType, LastConnectionTime, LastConnectionUser | Sort-Object DesktopGroupName
    foreach ($u in $Desktops){
        $IncUsers = Get-BrokerAccessPolicyRule -DesktopGroupName $u.DesktopGroupName | Select-Object -ExpandProperty IncludedUsers 
        $Users=$IncUsers.Name  | select -Unique
        $desktopgroupname = $u.DesktopGroupName
        $machinename = $u.MachineName
        $OSType = $u.OSType
        $lastconnectime = $u.LastConnectionTime
        $LastConnectionUser = $u.LastConnectionUser
        $xdddgsinforow = @("1")
        $Script:xdddgs.add(($xdddgsinforow | select @{n='Desktop Group Name';e={$desktopgroupname}},@{n='Machine Name';e={"$machinename"}},@{n='Operating System';e={"$OSType"}},@{n='Last Connection Time';e={"$lastconnectime"}},@{n='Last Connection User';e={"$LastConnectionUser"}})) | Out-Null
        }
    $xdddgscurrfile = $null
    $xdddgscurrfile = $xdddgscurrfile + "<title>XenDesktop Delivery Groups</title>"
    $xdddgscurrfile = $xdddgscurrfile + "<body><html>"
    $xdddgscurrfile = $xdddgscurrfile + $StdHeader
    $xdddgscurrfile = $xdddgscurrfile + "<h3>XenDesktop Delivery Groups as of $Script:date</h3>"
    $xdddgscurrfile = $xdddgscurrfile + ($Script:xdddgs | ConvertTo-Html -Fragment | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd)
    $xdddgscurrfile = $xdddgscurrfile + "</html></body>"
    $xddgfile = "XDDGs.html"
    $xdddgscurrfile | out-file -FilePath $ScriptReports$xddgfile
    try {copy-item $ScriptReports$xddgfile $HTMLReports -force}
    catch { "Error Copying $ScriptReports$xddgfile to $HTMLReports" | LogMe -error
        $_.Exception.Message | LogMe -error}
}

Function GetXenAppDGs{
    [System.Collections.ArrayList]$Script:xaddgs = @()
    $Desktops = Get-BrokerDesktop -maxrecordcount 5000 -adminaddress $DDCName | where {$_.OSType -match ( $ServerOSTypes -join '|')} | Select-Object DesktopGroupName, MachineName, OSType, LastConnectionTime, LastConnectionUser | Sort-Object DesktopGroupName
    foreach ($u in $Desktops){
        #$IncUsers = Get-BrokerAccessPolicyRule -DesktopGroupName $u.DesktopGroupName | Select-Object -ExpandProperty IncludedUsers 
        $Users=$IncUsers.Name  | select -Unique
        $desktopgroupname = $u.DesktopGroupName
        $machinename = $u.MachineName
        $OSType = $u.OSType
        #$lastconnectime = $u.LastConnectionTime
        #$LastConnectionUser = $u.LastConnectionUser
        $xaddgsinforow = @("1")
        $Script:xaddgs.add(($xaddgsinforow | select @{n='Desktop Group Name';e={$desktopgroupname}},@{n='Machine Name';e={"$machinename"}},@{n='Operating System';e={"$OSType"}})) | Out-Null
        }
    $xaddgscurrfile = $null  
    $xaddgscurrfile = $xaddgscurrfile + "<title>XenApp Delivery Groups</title>"
    $xaddgscurrfile = $xaddgscurrfile + "<body><html><h3>XenApp Delivery Groups as of $Script:date</h3>"
    $xaddgscurrfile = $xaddgscurrfile + $StdHeader
    $xaddgscurrfile = $xaddgscurrfile + ($Script:xaddgs | ConvertTo-Html -Fragment | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd)
    $xaddgscurrfile = $xaddgscurrfile + "</html>"
    $xaddgscurrfile = $xaddgscurrfile + "</body>"
    $xaddgfile = "XADDGs.html"
    $xaddgscurrfile | out-file -FilePath $ScriptReports$xaddgfile
    try {copy-item $ScriptReports$xaddgfile $HTMLReports -force}
    catch { "Error Copying $ScriptReports$xaddgfile to $HTMLReports" | LogMe -error
        $_.Exception.Message | LogMe -error}
}

Function GetXenAppPubApps{
    [System.Collections.ArrayList]$Script:xaapps = @()
    $Apps = Get-BrokerApplication * -adminaddress $DDCName | where {($_.OSType -match $ServerOSTypes -join '|')} | select AllAssociatedDesktopGroupUids,AdminFolderName, ApplicationType, ApplicationName, Description, Enabled, CommandLineExecutable, CommandLineArguments, WorkingDirectory, IconFromClient, PublishedName, WaitForPrinterCreation, ShortcutAddedToDesktop, ShortcutAddedToStartMenu  | Sort-Object ApplicationName
    foreach ($a in $Apps){
        $Appname = $a.ApplicationName
        $AdminFolderName = $a.AdminFolderName
        $ApplicationType = $a.ApplicationType
        $Description = $a.Description
        $Enabled = $a.Enabled
        $CommandLineExecutable = $a.CommandLineExecutable
        $CommandLineArguments = $a.CommandLineArguments
        $WorkingDirectory = $a.WorkingDirectory
        $IconFromClient = $a.IconFromClient
        $PublishedName = $a.PublishedName
        $WaitForPrinterCreation = $a.WaitForPrinterCreation
        $ShortcutAddedToDesktop = $a.ShortcutAddedToDesktop
        $ShortcutAddedToStartMenu = $a.ShortcutAddedToStartMenu
        [string]$AllAssociatedDesktopGroupUids = $a.AllAssociatedDesktopGroupUids
        $uid = Get-BrokerDesktopGroup -AdminAddress $DDCName -Uid $AllAssociatedDesktopGroupUids 
        $DeliveryGroupName = $uid.Name
        $xaappsinforow = @("1")
        $Script:xaapps.add(($xaappsinforow | select @{n='Application Name';e={$Appname}},@{n='Published Name';e={"$PublishedName"}},@{n='Delivery Group';e={"$DeliveryGroupName"}},@{n='Admin Folder Name';e={"$AdminFolderName"}},@{n='Application Type';e={"$ApplicationType"}},@{n='Description';e={"$Description"}},@{n='Is Enabled';e={"$Enabled"}},@{n='Command Line Executable';e={"$CommandLineExecutable"}},@{n='Command Line Arguments';e={"$CommandLineArguments"}},@{n='Working Directory';e={"$WorkingDirectory"}})) | Out-Null
        }
    $xaappsHeader = @" 
<title>XenApp Published Applications</title>
<head>
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
<meta http-equiv="refresh" content="60">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<style>
body {background-color: #a5a5a5;}
TABLE {margin-left: auto; margin-right: auto; border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;width: 95%} 
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #636363; font-size: 16px; color: #ffffff;} 
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;font-size: 14px;} 
H3 {text-align: center;}
.odd { background-color:#ffffff; } 
.even { background-color:#dddddd; } 
</style>
</head>
"@
    $xaappscurrfile = $null  
    $xaappscurrfile = $xaappscurrfile + $xaappsHeader
    $xaappscurrfile = $xaappscurrfile + "<body><h3>XenApp Published Applications as of $Script:date</h3><html>"
    $xaappscurrfile = $xaappscurrfile + ($Script:xaapps | ConvertTo-Html -Fragment | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd)
    $xaappscurrfile = $xaappscurrfile + "</html>"
    $xaappscurrfile = $xaappscurrfile + "</body>"
    $xaappsfile = "XAApps.html"
    $xaappscurrfile | out-file -FilePath $ScriptReports$xaappsfile
    try {copy-item $ScriptReports$xaappsfile $HTMLReports -force}
    catch { "Error Copying $ScriptReports$xaappsfile to $HTMLReports" | LogMe -error
        $_.Exception.Message | LogMe -error}
}

Function GetXAUsageReports($dur){
    if ($dur -eq "w") {
        $sdate = [datetime]::today 
        $ldate = ([datetime]::today).AddDays(-7) #### 1 week ago
        $rptrange = "Last 7 Days"
        } 
    elseif ($dur -eq "m") {
        $sdate = [datetime]::today 
        $ldate = ([datetime]::today).AddDays(-30) #### 30 days ago
        $rptrange = "Last 30 Days"
        }
    else {
        $sdate = [datetime]::today 
        $ldate = ([datetime]::today).AddDays(-1) #### yesterday
        $rptrange = "Yesterday"
    }
    $filter = "and logonenddate >= convert(datetime,'"+(get-date ($ldate) -Format "MM/dd/yyyy HH:mm:ss")+"') and logonenddate <= convert(datetime,'"+(get-date ($sdate) -Format "MM/dd/yyyy HH:mm:ss")+"')"
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security = True; MultipleActiveResultSets = True"
    $SQLQuery = @"
select monitordata.session.SessionKey
,monitordata.session.StartDate
,LogOnDuration
,monitordata.session.EndDate
,ConnectionState
,UserName
,FullName
,monitordata.application.Name
,PublishedName
,monitordata.machine.HostedMachineName
,monitordata.DesktopGroup.Name
,IsRemotePC
,DesktopKind
,SessionSupport
,DeliveryType
,ClientName
,ClientAddress
,ClientVersion
,ConnectedViaHostName
,ConnectedViaIPAddress
,LaunchedViaHostName
,LaunchedViaIPAddress
,IsReconnect
,Protocol
,LogOnStartDate
,LogOnEndDate
,BrokeringDuration
,BrokeringDate
,DisconnectCode
,DisconnectDate
,VMStartStartDate
,VMStartEndDate
,ClientSessionValidateDate
,ServerSessionValidateDate
,EstablishmentDate
,HdxStartDate
,AuthenticationDuration
,GpoStartDate
,GpoEndDate
,LogOnScriptsStartDate
,LogOnScriptsEndDate
,ProfileLoadStartDate
,ProfileLoadEndDate
,InteractiveStartDate
,InteractiveEndDate
,Datediff(minute,logonenddate,DisconnectDate) as 'SessionLength'
from monitordata.Session
join monitordata.[user] on monitordata.session.UserId = monitordata.[user].Id
join monitordata.Machine on monitordata.session.MachineId = monitordata.machine.Id
join monitordata.DesktopGroup on monitordata.machine.DesktopGroupId = monitordata.desktopgroup.Id
join monitordata.connection on monitordata.session.SessionKey = monitordata.connection.SessionKey
join monitordata.applicationinstance on monitordata.ApplicationInstance.SessionKey = monitordata.session.SessionKey
join monitordata.application on monitordata.application.id = monitordata.ApplicationInstance.ApplicationId
where UserName <> '' and sessiontype = '1'  and Protocol = 'HDX'
$filter
order by logonenddate,SessionKey
"@
    $SQLResult = sqlquery -q $SQLQuery | ?{$_ -notlike "*[0-9]*"}
if ($SQLResult -eq $Null){ "SQL Result is 0" }
elseif ($($SQLResult | measure).count -eq 1){
[System.Collections.ArrayList]$appsessions = @()
[System.Collections.ArrayList]$appsessions = @($SQLResult) # | ?{$_ -notlike "*[0-9]*"}
}
else {
[System.Collections.ArrayList]$appsessions = @()
[System.Collections.ArrayList]$appsessions = $SQLResult # | ?{$_ -notlike "*[0-9]*"}
}
    if ($appsessions -ne $null) {
        $appsessions | ?{$_.sessionlength.gettype().name -eq "dbnull"} | %{
        if ($_.connectionstate -eq "5") {
            $_.sessionlength = [math]::Round(($sdate - (get-date $_.logonenddate).ToLocalTime()).totalminutes,0)
            } 
        elseif ($_.connectionstate -eq "3") {
            $_.sessionlength = [math]::Round(($_.enddate - $_.logonenddate).totalminutes,0)
            }
        }
    }
    $allappsessions = $appsessions | sort username
    $allapps = ($appsessions).publishedname | sort -unique | sort
    $appusernames = $allappsessions.username | sort -Unique
    [System.Collections.ArrayList]$xainfo = @()
    $z = 0
    $ucount = ($appusernames | measure).count
    foreach ($auser in $appusernames) { 
        $z ++
        $t = @("1")
        $un1 = ($allappsessions | ?{$_.username -eq $auser}).fullname | sort -Unique | select -First 1
        $apps = ($allappsessions | ?{$_.username -eq $auser}).publishedname | sort -unique
        foreach ($app in $apps) {
            $xasescount = (($allappsessions | ?{$_.username -eq $auser -and $_.publishedname -eq $app}).sessionkey.guid | sort -Unique | measure).count
            $totalxasesscount += $xasescount
            $xaactivetime = (($allappsessions | ?{$_.username -eq $auser -and $_.publishedname -eq $app}).sessionlength | measure -Sum).sum
            $avghrs1 = [math]::round(($xaactivetime/$xasescount/60),2)
            $totalhrs1 = [math]::round(($xaactivetime/60),2)
            $xainfo.add(($t | select @{n='Published Application';e={$app}},@{n='User';e={$auser}},@{n='Name';e={$un1}},@{n='Session Count';e={$xasescount}},@{n='Total Hours';e={$totalhrs1}},@{n='Average Hours';e={$avghrs1}})) | Out-Null
            }
        }
    if ($totalxasesscount -eq $null){
        $totalxasesscount = 0
        "Totlal Sessions: " +$totalxasesscount
        }
    else {"Totlal Sessions: " +$totalxasesscount}
    $e = $sdate.tostring("MM/dd/yyyy")
    $s = $ldate.tostring("MM/dd/yyyy")
    if (($dur -eq "w") -or ($dur -eq "m")){
        $xaheading = "<h3>XenApp Sessions for the $rptrange ($totalxasesscount Total) $s - $e</h3>"
        }
    else {
        $xaheading = "<h3>XenApp Sessions $rptrange ($totalxasesscount Total) $s </h3>"
        }
    $file = $null
    $file = "<title>XenApp Usage $rptrange</title>"
    $file = $file + $StdHeader
    $file = $file + $xaheading
    $file = $file + ($xainfo | ConvertTo-Html -Fragment | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd)
    $file = $file + "</html>"
    $file = $file + "</body>"
    if ($file -ne $null -and $dur -eq "w") {
        $file | out-file -FilePath $XAlast7daysfilename
        try {copy-item $XAlast7daysfilename $HTMLReports -force}
        catch { "Error Copying $XAlast7daysfilename to $HTMLReports" | LogMe -error
        $_.Exception.Message | LogMe -error}
        }
    elseif ($file -ne $null -and $dur -eq "m") {
        $file | out-file -FilePath $XAlast30daysfilename
        try {copy-item $XAlast30daysfilename $HTMLReports -force}
        catch { "Error Copying $XAlast30daysfilename to $HTMLReports" | LogMe -error
        $_.Exception.Message | LogMe -error}
        }
    else {
        $file | out-file -FilePath $XAyesterdayfilename
        try {copy-item $XAyesterdayfilename $HTMLReports -force}
        catch { "Error Copying $XAyesterdayfilename to$HTMLReports" | LogMe -error
        $_.Exception.Message | LogMe -error}
        }
}

Function GetXDUsageReports($dur){
    if ($dur -eq "w") {
        $sdate = [datetime]::today 
        $ldate = ([datetime]::today).AddDays(-7) #### 1 week ago
        $rptrange = "Last 7 Days"
        } 
    elseif ($dur -eq "m") {
        $sdate = [datetime]::today 
        $ldate = ([datetime]::today).AddDays(-30) #### 30 days ago
        $rptrange = "Last 30 Days"
        } 
    else {
        $sdate = [datetime]::today 
        $ldate = ([datetime]::today).AddDays(-1) #### yesterday
        $rptrange = "Yesterday"
        }
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security = True; MultipleActiveResultSets = True"
$filter = "and logonenddate >= convert(datetime,'"+(get-date ($ldate) -Format "MM/dd/yyyy HH:mm:ss")+"') and logonenddate <= convert(datetime,'"+(get-date ($sdate) -Format "MM/dd/yyyy HH:mm:ss")+"')"
$SQLQuery = @"
select monitordata.session.SessionKey
,startdate
,logonduration
,enddate
,connectionstate
,username
,fullname
,monitordata.machine.HostedMachineName
,monitordata.desktopgroup.Name
,IsRemotePC
,DesktopKind
,SessionSupport
,SessionType
,DeliveryType
,ClientName
,ClientAddress
,ClientVersion
,ConnectedViaHostName
,ConnectedViaIPAddress
,LaunchedViaHostName
,LaunchedViaIPAddress
,IsReconnect
,Protocol
,LogOnStartDate
,LogOnEndDate
,BrokeringDuration
,BrokeringDate
,DisconnectCode
,DisconnectDate
,VMStartStartDate
,VMStartEndDate
,ClientSessionValidateDate
,ServerSessionValidateDate
,EstablishmentDate
,HdxStartDate
,HdxEndDate
,AuthenticationDuration
,GpoStartDate
,GpoEndDate
,LogOnScriptsStartDate
,LogOnScriptsEndDate
,ProfileLoadStartDate
,ProfileLoadEndDate
,InteractiveStartDate
,InteractiveEndDate
,Datediff(minute,logonenddate,DisconnectDate) as 'SessionLength'
 from monitordata.session
join monitordata.[user] on monitordata.session.UserId = monitordata.[user].Id
join monitordata.Machine on monitordata.session.MachineId = monitordata.machine.Id
join monitordata.DesktopGroup on monitordata.machine.DesktopGroupId = monitordata.desktopgroup.Id
join monitordata.connection on monitordata.session.SessionKey = monitordata.connection.SessionKey
where UserName <> '' and SessionType = '0' and Protocol = 'HDX'
$filter
order by logonenddate,SessionKey
"@
$SQLResult = sqlquery -q $SQLQuery | ?{$_ -notlike "*[0-9]*"}
if ($SQLResult -eq $Null){ "SQL Result is 0" }
elseif ($($SQLResult | measure).count -eq 1){
[System.Collections.ArrayList]$sessions = @()
[System.Collections.ArrayList]$sessions = @($SQLResult) # | ?{$_ -notlike "*[0-9]*"}
}

else {
[System.Collections.ArrayList]$sessions = @()
[System.Collections.ArrayList]$sessions = $SQLResult # | ?{$_ -notlike "*[0-9]*"}
}

if ($sessions -ne $null) {
    $sessions | ?{$_.sessionlength.gettype().name -eq "dbnull"} | %{
        if ($_.connectionstate -eq "5") {
            $_.sessionlength = [math]::Round(($sdate - (get-date $_.logonenddate).ToLocalTime()).totalminutes,0)
            } 
        elseif ($_.connectionstate -eq "3") {
            $_.sessionlength = [math]::Round(($_.enddate - $_.logonenddate).totalminutes,0)
        }
        }
    }
$allsessions = $sessions | sort username
$usernames = $allsessions.username | sort -unique
[System.Collections.ArrayList]$xdinfo = @()
$z = 0  
$ucount = ($usernames | measure).count 
foreach ($user in $usernames) { 
    $z ++
    $t = @("1")
    $un = ($allsessions | ?{$_.username -eq $user}).fullname | sort -Unique | select -First 1
    $xdsescount = (($allsessions | ?{$_.username -eq $user}).sessionkey.guid | sort -Unique | measure).count
    $totalxdsesscount += $xdsescount
    $activetime = (($allsessions | ?{$_.username -eq $user}).sessionlength | measure -Sum).sum
    $avghrs = [math]::round(($activetime/$xdsescount/60),2)
    $totalhrs = [math]::round(($activetime/60),2)
    $xdinfo.add(($t | select @{n='User';e={$user}},@{n='Name';e={$un}},@{n='Session Count';e={$xdsescount}},@{n='Total Hours';e={$totalhrs}},@{n='Average Hours';e={$avghrs}})) | Out-Null
    }
if ($totalxdsesscount -eq $null){
        "Totlal Sessions: " +$totalxdsesscount
        $totalxdsesscount = 0
        "Totlal Sessions: " +$totalxdsesscount
        }
else {"Totlal Sessions: " +$totalxdsesscount}
$e = $sdate.tostring("MM/dd/yyyy")
$s = $ldate.tostring("MM/dd/yyyy")
if (($dur -eq "w") -or ($dur -eq "m")){ $xdheading = "<h3>XenDesktop Sessions $rptrange ($totalxdsesscount Total) $s - $e</h3>" } 
else { $xdheading = "<h3>XenDesktop Sessions $rptrange ($totalxdsesscount Total) $s </h3>" }
$file = $null
$file = "<title>XenDesktop Usage $rptrange</title>"
$file = $file + $StdHeader
$file = $file + "<body><html>$xdheading"
$file = $file + ($xdinfo | ConvertTo-Html -Fragment | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd)
$file = $file + "</html>"
$file = $file + "</body>"
if ($file -ne $null -and $dur -eq "w") {
    $file | out-file -FilePath $XDlast7daysfilename
    try {copy-item $XDlast7daysfilename $HTMLReports -force}
    catch { "Error Copying $XDlast7daysfilename to $HTMLReports" | LogMe -error
    $_.Exception.Message | LogMe -error}
    } 
elseif ($file -ne $null -and $dur -eq "m") {
    $file | out-file -FilePath $XDlast30daysfilename
    try {copy-item $XDlast30daysfilename $HTMLReports -force}
    catch { "Error Copying $XDlast30daysfilename to $HTMLReports" | LogMe -error
    $_.Exception.Message | LogMe -error}
    }
else {
    $file | out-file -FilePath $XDyesterdayfilename
    try {copy-item $XDyesterdayfilename $HTMLReports -force}
    catch { "Error Copying $XDyesterdayfilename to $HTMLReports" | LogMe -error
    $_.Exception.Message | LogMe -error}
    }
}

Function GetPubAppUsageReport($dur){
    if ($dur -eq "w") {
        $sdate = [datetime]::today 
        $ldate = ([datetime]::today).AddDays(-7) #### 1 week ago
        $rptrange = "Last 7 Days"
        } 
    elseif ($dur -eq "m") {
        $sdate = [datetime]::today 
        $ldate = ([datetime]::today).AddDays(-30) #### 30 days ago
        $rptrange = "Last 30 Days"
        } 
    else {
        $sdate = [datetime]::today 
        $ldate = ([datetime]::today).AddDays(-1) #### yesterday
        $rptrange = "Yesterday"
        }
    $filter = "and logonenddate >= convert(datetime,'"+(get-date ($ldate) -Format "MM/dd/yyyy HH:mm:ss")+"') and logonenddate <= convert(datetime,'"+(get-date ($sdate) -Format "MM/dd/yyyy HH:mm:ss")+"')"
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security = True; MultipleActiveResultSets = True"

    $SQLQuery = @"
        select monitordata.application.Name
        ,PublishedName
        ,Protocol,
        CONVERT(VARCHAR(10), monitordata.session.StartDate, 101) as LaunchDate
        from monitordata.Session
        join monitordata.[user] on monitordata.session.UserId = monitordata.[user].Id
        join monitordata.Machine on monitordata.session.MachineId = monitordata.machine.Id
        join monitordata.DesktopGroup on monitordata.machine.DesktopGroupId = monitordata.desktopgroup.Id
        join monitordata.connection on monitordata.session.SessionKey = monitordata.connection.SessionKey
        join monitordata.applicationinstance on monitordata.ApplicationInstance.SessionKey = monitordata.session.SessionKey
        join monitordata.application on monitordata.application.id = monitordata.ApplicationInstance.ApplicationId
        where UserName <> '' and sessiontype = '1' and Protocol = 'HDX'
        $filter
        order by LaunchDate,PublishedName
"@
    $SQLResult = sqlquery -q $SQLQuery
    if ($SQLResult -eq $Null){ "SQL Result is 0" }
    elseif ($($SQLResult | measure).count -eq 1){
        [System.Collections.ArrayList]$pubappsrpt = @()
        [System.Collections.ArrayList]$pubappsrpt = @($SQLResult)
        }
    else {
        [System.Collections.ArrayList]$pubappsrpt = @()
        [System.Collections.ArrayList]$pubappsrpt = $SQLResult
        }
    $allpappsq = $pubappsrpt | sort PublishedName
    $allpaapps = $allpappsq.publishedname | sort -Unique
    $pubcnt = @()
    foreach ($pa in $allpaapps){
        $xapacount = (($allpappsq | ?{$_.publishedname -eq $pa}) | measure).count
        $item = New-Object PSObject
        $item | Add-Member -MemberType NoteProperty -Name "PubApp" -Value $pa -Force
        $item | Add-Member -MemberType NoteProperty -Name "Count" -Value $xapacount -Force
        $pubcnt += $item
        }
    $patotals = $pubcnt | sort -Descending count
    $patotalscnt = $patotals.count
    "Total Applications: " + $patotalscnt
    $e = $sdate.tostring("MM/dd/yyyy")
    $s = $ldate.tostring("MM/dd/yyyy")
    if (($dur -eq "w") -or ($dur -eq "m")){
        $xapaheading = "<h3>Published Applications Launched $rptrange ($patotalscnt Total) $s - $e</h3>"
        } 
    else {
        $xapaheading = "<h3>Published Applications Launched $rptrange ($patotalscnt Total) $s </h3>"
        }
    $file = $null
    $file = ""
    $file = "<title>Published Applications Launched $rptrange</title>"
    $file = $file + $StdHeader
    $file = $file + "<body><html>$xapaheading"
    $file = $file + ($patotals | ConvertTo-Html -Fragment | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd)
    $file = $file + "</html>"
    $file = $file + "</body>"
    if ($file -ne $null -and $dur -eq "w") {
        $file | out-file -FilePath $XAPAlast7daysfilename
        try {copy-item $XAPAlast7daysfilename $HTMLReports -force}
        catch { "Error Copying $XAPAlast7daysfilename to $HTMLReports" | LogMe -error
        $_.Exception.Message | LogMe -error}
        } 
    elseif ($file -ne $null -and $dur -eq "m") {
        $file | out-file -FilePath $XAPAlast30daysfilename
        try {copy-item $XAPAlast30daysfilename $HTMLReports -force}
        catch { "Error Copying $XAPAlast30daysfilename to $HTMLReports" | LogMe -error
        $_.Exception.Message | LogMe -error}
        } 
    else {
        $file | out-file -FilePath $XAPAyesterdayfilename
        try {copy-item $XAPAyesterdayfilename $HTMLReports -force}
        catch { "Error Copying $XAPAyesterdayfilename to $HTMLReports" | LogMe -error
        $_.Exception.Message | LogMe -error}
        }
}

Function GetClientVersionsReport($dur){
    if ($dur -eq "w") {
        $sdate = [datetime]::today 
        $ldate = ([datetime]::today).AddDays(-7) #### 1 week ago
        $rptrange = "Last 7 Days"
        } 
    else {
        $sdate = [datetime]::today 
        $ldate = ([datetime]::today).AddDays(-30) #### Last 30 days
        $rptrange = "Last 30 days"
        }
    $filter = "and logonenddate >= convert(datetime,'"+(get-date ($ldate).ToUniversalTime() -Format "MM/dd/yyyy HH:mm:ss")+"') and logonenddate <= convert(datetime,'"+(get-date ($sdate).ToUniversalTime() -Format "MM/dd/yyyy HH:mm:ss")+"')"
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security = True; MultipleActiveResultSets = True"
    [System.Collections.ArrayList]$sessions = @()
    [System.Collections.ArrayList]$sessions = sqlquery -q `
"select
monitordata.session.SessionKey
,startdate
,logonduration
,enddate
,connectionstate
,username
,fullname
,monitordata.machine.HostedMachineName
,monitordata.desktopgroup.Name
,IsRemotePC
,DesktopKind
,SessionSupport
,SessionType
,DeliveryType
,ClientName
,ClientAddress
,ClientVersion
,ConnectedViaHostName
,ConnectedViaIPAddress
,LaunchedViaHostName
,LaunchedViaIPAddress
,IsReconnect
,Protocol
,LogOnStartDate
,LogOnEndDate
,BrokeringDuration
,BrokeringDate
,DisconnectCode
,DisconnectDate
,VMStartStartDate
,VMStartEndDate
,ClientSessionValidateDate
,ServerSessionValidateDate
,EstablishmentDate
,HdxStartDate
,HdxEndDate
,AuthenticationDuration
,GpoStartDate
,GpoEndDate
,LogOnScriptsStartDate
,LogOnScriptsEndDate
,ProfileLoadStartDate
,ProfileLoadEndDate
,InteractiveStartDate
,InteractiveEndDate
,Datediff(minute,logonenddate,DisconnectDate) as 'SessionLength'
 from monitordata.session
join monitordata.[user] on monitordata.session.UserId = monitordata.[user].Id
join monitordata.Machine on monitordata.session.MachineId = monitordata.machine.Id
join monitordata.DesktopGroup on monitordata.machine.DesktopGroupId = monitordata.desktopgroup.Id
join monitordata.connection on monitordata.session.SessionKey = monitordata.connection.SessionKey
where UserName <> '' and SessionType = '0' and Protocol = 'HDX' and ClientVersion <> ''
$filter
order by logonenddate,SessionKey" | ?{$_ -notlike "*[0-9]*"}
    if ($sessions -ne $null) {
        $sessions | ?{$_.sessionlength.gettype().name -eq "dbnull"} | %{
            if ($_.connectionstate -eq "5") {
                $_.sessionlength = [math]::Round(($sdate - (get-date $_.logonenddate).ToLocalTime()).totalminutes,0)
                } 
            elseif ($_.connectionstate -eq "3") {
                $_.sessionlength = [math]::Round(($_.enddate - $_.logonenddate).totalminutes,0)
                }
            }
        }
    $allsessions = $sessions | sort ClientVersion
    $clientversions = $allsessions.ClientVersion | sort -unique
    [System.Collections.ArrayList]$cvinfo = @()
    $z = 0  
    $ucount = ($clientversions | measure).count 
    foreach ($version in $clientversions) { 
        $z ++
        $t = @("1")
        $cv = (($allsessions | ?{$_.ClientVersion -eq $version}).ClientVersion | measure).count
        $cvinfo.add(($t | select @{n='Client Version';e={$version}},@{n='Session Count';e={$cv}})) | Out-Null
        }
    $e = $sdate.tostring("MM/dd/yyyy")
    $s = $ldate.tostring("MM/dd/yyyy")
    $cvfile = $null
    $cvfile = "<title>Client Versions $rptrange</title>"
    $cvfile = $cvfile + $StdHeader
    $cvfile = $cvfile + "<body><html><h3>Client Versions $s - $e </h3>"
    $cvfile = $cvfile + ($cvinfo | ConvertTo-Html -Fragment | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd)
    if ($cvfile -ne $null -and $dur -eq "w") {
        $cvfile = $cvfile + "<p style=`"text-align: center;`"><a href=`"clientverdetails7days.html`" onclick=`"window.open('clientverdetails7days.html', 'verlast7det', 'width=1000,height=800'); return false;`">Citrix Receiver Versions Last 7 Days - Details</a></p>"
        $cvfile = $cvfile + "</html></body>" 
        $cvfile | out-file -FilePath $clientversions7daysfilename
        try {copy-item $clientversions7daysfilename $HTMLReports -force}
        catch { "Error Copying $clientversions7daysfilename to $HTMLReports" | LogMe -error
        $_.Exception.Message | LogMe -error}
        } 
    else {
        $cvfile = $cvfile + "<p style=`"text-align: center;`"><a href=`"clientverdetails30days.html`" onclick=`"window.open('clientverdetails30days.html', 'verlast30det', 'width=1000,height=800'); return false;`">Citrix Receiver Versions Last 30 Days - Details</a></p>"
        $cvfile = $cvfile + "</html></body>" 
        $cvfile | out-file -FilePath $clientversions30daysfilename
        try {copy-item $clientversions30daysfilename $HTMLReports -force}
        catch { "Error Copying $clientversions30daysfilename to $HTMLReports" | LogMe -error
        $_.Exception.Message | LogMe -error}
        }
    $ClientNames = $allsessions.ClientName | sort -unique
    [System.Collections.ArrayList]$cvdetinfo = @()
    $z = 0  
    $ucount = ($ClientNames | measure).count 
    foreach ($Client in $ClientNames) { 
        $z ++
        $t = @("1")
        $cv = ($allsessions | ?{$_.ClientName -eq $Client}).ClientVersion | sort -Unique | select -First 1
        $cvsesscount = (($allsessions | ?{$_.ClientName -eq $Client}).sessionkey.guid | sort -Unique | measure).count
        $cvdetinfo.add(($t | select @{n='User';e={$Client}},@{n='Name';e={$cv}},@{n='Session Count';e={$cvsesscount}})) | Out-Null
        }
    $e = $sdate.tostring("MM/dd/yyyy")
    $s = $ldate.tostring("MM/dd/yyyy")
    $cvdfile = $null
    $cvdfile = "<title>Client Versions $rptrange</title>" 
    $cvdfile = $cvdfile + $StdHeader
    $cvdfile = $cvdfile + "<body><html><h3>Client Versions $s - $e </h3>"
    $cvdfile = $cvdfile + ($cvdetinfo | ConvertTo-Html -Fragment | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd)
    $cvdfile = $cvdfile + "</html></body>"   
    if ($cvdfile -ne $null -and $dur -ne "w") {
        $cvdfile | out-file -FilePath $clientversiondetailslast30filename
        try {copy-item $clientversiondetailslast30filename $HTMLReports -force}
        catch { "Error Copying $clientversiondetailslast30filename to $HTMLReports" | LogMe -error
        $_.Exception.Message | LogMe -error}
        } 
    else {
        $cvdfile | out-file -FilePath $clientversiondetailslast7filename
        try {copy-item $clientversiondetailslast7filename $HTMLReports -force}
        catch { "Error Copying $clientversiondetailslast7filename to $HTMLReports" | LogMe -error
        $_.Exception.Message | LogMe -error}
    }
}

Function GetXDCurrentUsage{
    [System.Collections.ArrayList]$xdcurrinfo = @()
    $Script:XDActiveSess=0
    $Script:XDDiscSess=0
    $XDsessions=get-brokersession -AdminAddress $DDCName -protocol HDX|select MachineName,sessionstate,appstate,protocol,SessionType,applicationsinuse,BrokeringUserName,ClientVersion,ConnectionMode,Clientaddress,SessionSupport,UserName | where {$_.SessionSupport -eq "Single" }
    $XDComputers=get-brokerdesktop -AdminAddress $DDCName | where {$_.OSType -match ( $DesktopOSTypes -join '|') -and ($_.SummaryState -eq "InUse") -or ($_.SummaryState -eq "Disconnected") } | select CatalogName,DesktopGroupName,AssociatedUserUPNs,AgentVersion,InMaintenanceMode,LastConnectionFailure,LastDeregistrationTime,MachineInternalState,MachineName,Tags,DesktopKind,OSType,ApplicationsInUse,SessionUserName | sort SessionUserName
    #$XDComputers=get-brokerdesktop -AdminAddress $DDCName | where {$_.OSType -match ( $DesktopOSTypes -join '|') -and ($_.SummaryState -ne "Available") -and ($_.SummaryState -ne "off") -and ($_.SummaryState -ne "Unregistered")} | select CatalogName,DesktopGroupName,AssociatedUserUPNs,AgentVersion,InMaintenanceMode,LastConnectionFailure,LastDeregistrationTime,MachineInternalState,MachineName,Tags,DesktopKind,OSType,ApplicationsInUse,SessionUserName | sort SessionUserName
    $XDComputers|%{
        $objserver=""|select servername,load,missingwk,maintenance,activesessions,inactivesessions,pendingreboot,applicationsinuse,SessionUserName,DesktopGroupName,Other,PreparingSession,Connected,Reconnecting
        $machinename=$_.machinename
        $SessionUserName=$_.SessionUserName
        $objserver.servername=$machinename.split('\')[1]
        $objserver.activesessions=($XDsessions|?{$_.machinename -eq $machinename -and $_.sessionstate -eq "Active"}|measure-object).count
        $objserver.inactivesessions=($XDsessions|?{$_.machinename -eq $machinename -and $_.sessionstate -eq "Disconnected"}|measure-object).count
        $objserver.Other=($XDsessions|?{$_.machinename -eq $machinename -and $_.sessionstate -eq "Other"}|measure-object).count
        $objserver.PreparingSession=($XDsessions|?{$_.machinename -eq $machinename -and $_.sessionstate -eq "PreparingSession"}|measure-object).count
        $objserver.Connected=($XDsessions|?{$_.machinename -eq $machinename -and $_.sessionstate -eq "Connected"}|measure-object).count
        $objserver.Reconnecting=($XDsessions|?{$_.machinename -eq $machinename -and $_.sessionstate -eq "Reconnecting"}|measure-object).count
        $objserver.DesktopGroupName=$_.DesktopGroupName
        If ($($objserver.activesessions) -eq 1){
            $xdsesstat="Active"
            $Script:XDActiveSess++
            }
        elseIf ($($objserver.inactivesessions) -eq 1){
            $xdsesstat="Disconnected"
            $Script:XDDiscSess++
            }
        elseif ($($objserver.Other) -eq 1) {$xdsesstat="Other"}
        elseif ($($objserver.PreparingSession) -eq 1) {$xdsesstat="PreparingSession"}
        elseif ($($objserver.Connected) -eq 1) {$xdsesstat="Connected"}
        elseif ($($objserver.Reconnecting) -eq 1) {$xdsesstat="Reconnecting"}
        $xdcurrrow = @("1")
        $xdcurrinfo.add(($xdcurrrow | select @{n='User';e={$SessionUserName}},@{n='Computer Name';e={$($objserver.servername)}},@{n='Session State';e={$xdsesstat}},@{n='Desktop Group Name';e={$($objserver.DesktopGroupName)}})) | Out-Null
        }
    $Script:xdcurrcount = ($xdcurrinfo | measure).count
    "XD Current Count: " +$Script:xdcurrcount
    "XD Active Count: " +$Script:XDActiveSess
    "XD Disconnected Count: " +$Script:XDDiscSess
    $Script:XDCurrentGraph = @()
    $item = New-Object PSObject
    $item | Add-Member -MemberType NoteProperty -Name "XDTotalCount" -Value $Script:xdcurrcount -Force
    $item | Add-Member -MemberType NoteProperty -Name "XDActiveCount" -Value $Script:XDActiveSess -Force
    $item | Add-Member -MemberType NoteProperty -Name "XDDisconnectedCount" -Value $Script:XDDiscSess -Force
    $Script:XDCurrentGraph += $item
    CreateXenDesktopUsageChart
    $xdcurrfile = $null
    $xdcurrfile = "<title>$MonitorName XenDesktop Current Usage</title>"
    $xdcurrfile = $xdcurrfile + $CountDownHeader
    $xdcurrfile = $xdcurrfile + "<body><html><h3>$MonitorName Current XenDesktop Sessions ($Script:xdcurrcount Total) as of $Script:Date Auto-Refresh in <strong><font color='#003399'><span id=`"CDTimer`">180</span> secs. <script language=`"JavaScript`" type=`"text/javascript`"> /*<![CDATA[*/ var TimerVal = 60; var TimerSPan = document.getElementById(`"CDTimer`"); function CountDown(){ setTimeout( `"CountDown()`", 1000 ); TimerSPan.innerHTML=TimerVal; TimerVal=TimerVal-1;} CountDown() /*]]>*/ </script></font></strong></h3>"
    $xdcurrfile = $xdcurrfile + ($xdcurrinfo | ConvertTo-Html -Fragment | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd)
    $xdcurrfile = $xdcurrfile + "</html></body>"
    $xdcurrfile | out-file -FilePath $XDCurrentFileName
    try {copy-item $XDCurrentFileName $HTMLCurrent -force}
    catch { "Error Copying $XDCurrentFileName to $HTMLCurrent" | LogMe -error
        $_.Exception.Message | LogMe -error}
}

Function GetXACurrentUsage{
    [System.Collections.ArrayList]$xacurrinfo = @()
    $XAsessions=get-brokersession -AdminAddress $DDCName -protocol HDX | where {($_.SessionSupport -eq "MultiSession") -and ($_.UserName -ne $null) } | select MachineName,sessionstate,appstate,protocol,SessionType,applicationsinuse,BrokeringUserName,ClientVersion,ConnectionMode,Clientaddress,SessionSupport,UserName
    "Total Sessions: " + ($XAsessions | measure-object).count
    foreach ($User in $XAsessions){
        $Username = $User.UserName
        $MachineName = $User.MachineName
        $SessionState = $User.SessionState
        $AppState = $User.AppState
        $ApplicationsInUse = $User | select –ExpandProperty applicationsinuse
        $ClientAddress = $User.ClientAddress
        $AppCount = ($ApplicationsInUse|measure-object).Count
        if ($AppCount -gt 1) {
            $App = $Null
            $AppTitle = "Published Applications"
            foreach ($a in $ApplicationsInUse){
                $App += $a + "LINEBREAK"
                }
            }
        else {
            $App = $Null
            $App = $ApplicationsInUse
            $AppTitle = "Published Application"
            }
        $xacurrrow = @("1")
        $xacurrinfo.add(($xacurrrow | select @{n='User';e={$Username}},@{n='Server Name';e={$MachineName}},@{n='Session State';e={$SessionState}},@{n='App State';e={$AppState}},@{n='Apps in Use';e={$App}})) | Out-Null
        }
    $Script:xacurrcount = ($XAsessions | measure).count
    $Script:xacurractive = ($XAsessions |  ? {$_.sessionstate -ne "Disconnected"} | measure ).count
    $Script:xacurrdisc = ($XAsessions |  ? {$_.sessionstate -eq "Disconnected"} | measure).count
    if ($Script:xacurrcount -eq $Null) {$Script:xacurrcount = 0}
    if ($Script:xacurractive -eq $Null) {$Script:xacurractive = 0}
    if ($Script:xacurrdisc -eq $Null) {$Script:xacurrdisc = 0}
    "XA Current Count: " + $Script:xacurrcount
    "XA Active Count: " + $Script:xacurractive
    "XA Disconnected Count: " + $Script:xacurrdisc
    $Script:XACurrentGraph = @()
    $item = New-Object PSObject
    $item | Add-Member -MemberType NoteProperty -Name "XATotalCount" -Value $Script:xacurrcount -Force
    $item | Add-Member -MemberType NoteProperty -Name "XAActiveCount" -Value $Script:xacurractive -Force
    $item | Add-Member -MemberType NoteProperty -Name "XADisconnectedCount" -Value $Script:xacurrdisc -Force
    $Script:XACurrentGraph += $item
    CreateXenAppUsageChart
    $xacurrfile = $null
    $xacurrfile = "<title>$MonitorName XenApp Current Usage</title>"
    $xacurrfile = $xacurrfile + $CountDownHeader
    $xacurrfile = $xacurrfile + "<body><html><h3>$MonitorName Current XenApp Sessions ($Script:xacurrcount Total) as of $Script:Date  Auto-Refresh in <strong><font color='#003399'><span id=`"CDTimer`">180</span> secs. <script language=`"JavaScript`" type=`"text/javascript`"> /*<![CDATA[*/ var TimerVal = 60; var TimerSPan = document.getElementById(`"CDTimer`"); function CountDown(){ setTimeout( `"CountDown()`", 1000 ); TimerSPan.innerHTML=TimerVal; TimerVal=TimerVal-1;} CountDown() /*]]>*/ </script></font></strong></h3>"
    $xacurrfile = $xacurrfile + ($xacurrinfo | ConvertTo-Html -Fragment | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd)
    $xacurrfile = $xacurrfile + "</html></body>"
    $xacurrfile = $xacurrfile | foreach { ($_ -replace "LINEBREAK",$("<br/>"))}
    $xacurrfile | out-file -FilePath $XACurrentFileName
    try {copy-item $XACurrentFileName $HTMLCurrent -force}
    catch { "Error Copying $XACurrentFileName to $HTMLCurrent" | LogMe -error
        $_.Exception.Message | LogMe -error}
}

Function GetXAServerInfo{
    $XAsessions=get-brokersession -AdminAddress $DDCName -protocol HDX | where {$_.SessionSupport -eq "MultiSession" } | select MachineName,sessionstate,appstate,protocol,SessionType,applicationsinuse,BrokeringUserName,ClientVersion,ConnectionMode,Clientaddress,SessionSupport,UserName 
    $XAComputers=get-brokerdesktop -AdminAddress $DDCName | where {($_.OSType -match ( $ServerOSTypes -join '|'))} 
    #[System.Collections.ArrayList]$Script:xainfo = @()
    $XAComputers|%{
	    $objserver=""|select servername,load,maintenance,activesessions,inactivesessions,pendingreboot,applicationsinuse,UserName,DesktopGroupName,AppState
        $XAmachinename=$_.machinename
	    $Server=$XAmachinename.split('\')[1]
        $ServerCol = "ServerCol $Server"
        $AgVer = $_.AgentVersion
        $InMaint = $_.InMaintenanceMode
        $RegState = $_.RegistrationState
        if (test-connection -computername $Server -quiet){
            "Collecting info for $Server"
            $objserver.activesessions=($XAsessions|?{$_.machinename -eq $XAmachinename -and $_.sessionstate -eq "Active"}|measure-object).count
            $XAActiveSess = $objserver.activesessions
            $objserver.inactivesessions=($XAsessions|?{$_.machinename -eq $XAmachinename -and $_.sessionstate -eq "Disconnected"}|measure-object).count
            $XADiscSess = $objserver.inactivesessions
            "     Checking for pending reboot for $Server"
            $PendReboot=get-pendingreboot $Server
            if ($PendReboot -eq "True") {$PendReboot="Red $PendReboot"}
            "     Checking free disk space for $Server"
            $freeGB=get-freespace $Server
            if ($freeGB -like "Critical*") {$freeGB="Red $freeGB"}
            elseif ($freeGB -like "Warning*") {$freeGB="Warn $freeGB"}
            "     Checking memory usage for $Server"
            $MemUsg=CheckMemoryUsage $Server
            if ([int] $MemUsg -lt 80){$MemUsg = $MemUsg}
            elseif ([int] $MemUsg -lt 90){$MemUsg = "Warn $MemUsg"}
            elseif ([int] $MemUsg -lt 100){$MemUsg = "Red $MemUsg"}
            "     Checking CPU usage for $Server"
            $CPUUsg=CheckCpuUsage $Server
            if ([int] $CPUUsg -lt 80){$CPUUsg = $CPUUsg}
            elseif ([int] $CPUUsg -lt 90){$CPUUsg = "'Warn $CPUUsg"}
            elseif ([int] $CPUUsg -lt 100){$CPUUsg = "Red $CPUUsg"}
            "     Getting Event Log Count for $Server"
            $EventLogErrors=Get-EventLogCount $Server
            if ($Script:NumofEventsCount -gt 0) { 
                "     Getting Event Logs for $Server"
                $EventLog = Get-EventLogs $Server }
            else {$EventLog = ""}
            "     Getting Uptime for $Server"
            $ServerUpTime = Get-Uptime $Server
            "     Checking Services for $Server"
            CheckService $Server
            }
        else {
            "Cannot connect to $Server during XenApp Info collection" | LogMe -error
            $EventLog = "Warn N/A"
            $XAActiveSess = "Warn N/A"
            $XADiscSess = "Warn N/A"
            $PendReboot = "Warn N/A"
            $freeGB = "Warn N/A"
            $MemUsg = "Warn N/A"
            $CPUUsg = "Warn N/A"
            $EventLogErrrors = "Warn N/A"
            $ServerUpTime = "Warn N/A"
            $Script:ServiceStatus = "Warn N/A"
            }
        $evcurrfile = $null  
        $evcurrfile = "<title>Crtitical and Error events on $Server</title>"
        $evcurrfile = $evcurrfile + $EVHeader
        $evcurrfile = $evcurrfile + "<body><html><h3>Crtitical and Error events on $Server (last $EventLogCheck minutes) as of $(get-date)</h3>"
        $evcurrfile = $evcurrfile + ($EventLog | ConvertTo-Html -Fragment TimeCreated,ID,Message | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd)
        $evcurrfile = $evcurrfile + "</html></body>"
        $EventFile = "EventLog$Server.html"
        $evcurrfile | out-file -FilePath "$ScriptEvents$EventFile"
        try {copy-item $ScriptEvents$EventFile $HTMLEvents$EventFile -force}
        catch { "Error Copying $ScriptEvents$EventFile to $HTMLEvents$EventFile" | LogMe -error
            $_.Exception.Message | LogMe -error}
        #$xainforow = @("1")
        #$Script:xainfo.add(($xainforow | select @{n='XA Server';e={$ServerCol}},@{n='Free Disk Space';e={"$freeGB %"}},@{n='Memory Usage';e={"$MemUsg %"}},@{n='CPU Usage';e={"$CPUUsg %"}},@{n='Active Sessions';e={$XAActiveSess}},@{n='Disconnected Sessions';e={$XADiscSess}},@{n='Error Events';e={"$Script:NumofEventsCount"}},@{n='Uptime (days)';e={"$ServerUpTime"}},@{n='Registration State';e={$RegState}},@{n='In Maintenance';e={$InMaint}},@{n='Agent Version';e={$AgVer}},@{n='Pending Reboot';e={$($PendReboot)}},@{n='Services not running';e={$($Script:ServiceStatus)}})) | Out-Null
        $Script:XAInfoHTML = $Script:XAInfoHTML+"<tr class='trxainfo'><td class='tdxaservercol'>$ServerCol</td><td class='tdxainfo'>$freeGB</td><td class='tdxainfo'>$MemUsg %</td><td class='tdxainfo'>$CPUUsg %</td><td class='tdxainfo'>$XAActiveSess</td><td class='tdxainfo'>$XADiscSess</td><td class='tdxainfo'>$EventLogErrors</td><td class='tdxainfo'>$ServerUpTime</td><td class='tdxainfo'>$RegState</td><td class='tdxainfo'>$InMaint</td><td class='tdxainfo'>$AgVer</td><td class='tdxainfo'>$PendReboot</td><td class='tdxainfo'>$Script:ServiceStatus</td></tr>"
    }
    $xainfosection = "<table class='tblxainfo'><colgroup><col/><col/><col/><col/><col/><col/><col/><col/><col/><col/><col/><col/><col/></colgroup><tr class='trxainfo'><th class='thxainfo'>XenApp Server</th><th class='thxainfo'>Free Disk Space</th><th class='thxainfo'>Memory Usage</th><th class='thxainfo'>CPU Usage</th><th class='thxainfo'>Active Sessions</th><th class='thxainfo'>Disconnected Sessions</th><th class='thxainfo'>Error Events</th><th class='thxainfo'>Uptime (days)</th><th class='thxainfo'>Registration State</th><th class='thxainfo'>In Maintenance</th><th class='thxainfo'>Agent Version</th><th class='thxainfo'>Pending Reboot</th><th class='thxainfo'>Services not running</th></tr>"
    $xainfosection = $xainfosection + $Script:XAInfoHTML | 
    foreach { ($_ -replace "<td class='tdxaservercol'>ServerCol",$("<td class='tdxaservercol'>"))} |
    foreach { ($_ -replace "<td class='tdxainfo'>Stopped</td>",$("<td class='tdxainfoerror'>Stopped</td>"))} | 
    foreach { ($_ -replace "<td class='tdxainfo'>Unregistered</td>",$("<td class='tdxainfoerror'>Unregistered</td>"))} | 
    foreach { ($_ -replace "<td class='tdxainfo'>True</td>",$("<td class='tdxainfowarn'>True</td>"))} | 
    foreach { ($_ -replace "<td class='tdxainfo'>Yes</td>",$("<td class='tdxainfowarn'>Yes</td>"))} | 
    foreach { ($_ -replace "<td class='tdxainfo'>Warn",$("<td class='tdxainfowarn'>"))} | 
    foreach { ($_ -replace "<td class='tdxainfo'>Red",$("<td class='tdxainfoerror'>"))} |
    foreach { ($_ -replace "return false;`"><div style=`"height:100%;width:100%`">Red",$("return false;`"><div style=`"color:#FFFFFF;height:100%;width:100%`">"))}
    $xainfosection = $xainfosection + "</table>"
    $XAInfoFileName = "XAInfo.txt"
    $xainfosection | Out-File "XAInfo.txt"
    $date24 | out-file $XAInfoTimeStampFile
}

Function CollectXDPerf {            
    $Script:obj = @()
    $Script:XDobjover300 = @()
    $counters = "\ICA Session(*)\Latency - Last Recorded"
    foreach ($Server in $Servers) {
        $Computername = $Server.DNSName
        $TempError = $null
        if (test-connection -computername $Computername -quiet){
            if ($Server.OSType -like "*Windows*"){
                "Checking Computer: $Computername"
                try { $CheckCounter = Get-Counter -ListSet "ICA Session" -ComputerName $Computername -ErrorAction Stop }
                catch {
                Error-Message
                $TempError = $($_.Exception.Message)
                }
                if ( $TempError -like "*find any performance counter sets*") {
                "ICA Session Performance Counters are missing on Server: $Computername" | LogMeLatency -Error
                Continue}
                else {
                    try { $metrics = Get-Counter -ComputerName $Computername -Counter $counters -ErrorAction stop }
                    catch {$_.Exception.Message; Error-Message }
                    if ($? -eq $false) {"ICA Session Counter Exists, but there is nothing to report. Probably disconnected or no data exists." | LogMeLatency -DisplayNormal} 
                    else {
                        foreach($metric in $metrics.CounterSamples) {
                            if (($metric.InstanceName -like "console*") -or ($metric.InstanceName -like "*ica*")){
                                $newname = ""
                                $newname = $metric.InstanceName -replace '.*\(' -replace '\)'
                                $XDUser = $Server.BrokeringUserName
                                # add these columns as data
                                $item = New-Object PSObject
                                $item | Add-Member -MemberType NoteProperty -Name "LatencyinMS" -Value $metric.CookedValue -Force
                                $item | Add-Member -MemberType NoteProperty -Name "DateTime" -Value $metric.Timestamp -Force
                                $item | Add-Member -MemberType NoteProperty -Name "User" -Value $newname -Force
                                $item | Add-Member -MemberType NoteProperty -Name "Computer" -Value $computername   -Force
                                $xdsession = Get-BrokerSession -AdminAddress $DDCName | Where-Object DNSName -eq $computername | where-object {($_.SessionState -eq "Active") -or ($_.SessionState -eq "Application")} | Where-Object UserName -eq $XDUser
                                $item | Add-Member -MemberType NoteProperty -Name "ClientIP" -Value $xdsession.ClientAddress  -Force
                                "User: " +$XDUser
                                "Client IP Address: " + $xdsession.ClientAddress
                                "Current Latency: " + $metric.CookedValue
                                if ($metric.CookedValue -ge 10){
                                        "Current Latency is greater than 10, taking note of this session"
                                        $Script:obj += $item
                                        }
                                elseif ($metric.CookedValue -ge $MStoBeConsideredLatent){
                                        "Current Latency is greater than $MStoBeConsideredLatent, taking note of this session"
                                        $Script:XDobjover300 += $item
                                        }
                                else {"Current Latency is less than 10, not taking note of this session"}
                                }
                            }
                        }
                    }
                }
            } 
        else { "Connection to $Computername failed" | LogMeLatency -error}
        }
}

Function CollectXAPerf {            
    $Script:obj = @()
    $Script:XAobjover300 = @()
    $counters = "\ICA Session(*)\Latency - Last Recorded"
    if ($($Servers | measure).count -gt 0) {
        "Servers to query:"
        $Servers.DNSName
        foreach ($Server in $Servers) {
            $Computername = $Server.DNSName
            $TempError = $null
            $metrics = $Null
            if (test-connection -computername $Computername -quiet){
                if ($Server.OSType -like "*Windows*"){
                    "Checking Computer: $Computername"
                    try { $CheckCounter = Get-Counter -ListSet "ICA Session" -ComputerName $Computername -ErrorAction Stop }
                    catch {
                        Error-Message
                        $TempError = $($_.Exception.Message)
                        }
                    if ( $TempError -like "*find any performance counter sets*") {
                        "ICA Session Performance Counters are missing on Server: $Computername" 
                        Continue
                        }
                    else {
                        try { $metrics = Get-Counter -ComputerName $Computername -Counter $counters -ErrorAction stop }
                        catch {
                            $_.Exception.Message; #Error-Message 
                            }
                        if ($? -eq $false) {"ICA Session Counter Exists, but there is nothing to report. Maybe disconnected, lingering or no data exists."} 
                        else {
                            $XDUser = $Server.BrokeringUserName
                            $domain,$username = $XDUser.split('\')
                            foreach($metric in $metrics.CounterSamples) {
                                $metric.CounterSamples
                                if (($metric.InstanceName -like "*ica*") -and ($metric.InstanceName -like "*"+$username+"*")){
                                    $newname = ""
                                    $newname = $metric.InstanceName -replace '.*\(' -replace '\)'
                                    $XDUser = $Server.BrokeringUserName
                                    # add these columns as data
                                    $item = New-Object PSObject
                                    $item | Add-Member -MemberType NoteProperty -Name "LatencyinMS" -Value $metric.CookedValue -Force
                                    $item | Add-Member -MemberType NoteProperty -Name "DateTime" -Value $metric.Timestamp -Force
                                    $item | Add-Member -MemberType NoteProperty -Name "User" -Value $newname -Force
                                    $item | Add-Member -MemberType NoteProperty -Name "Computer" -Value $computername   -Force
                                    $xdsession = Get-BrokerSession -AdminAddress $DDCName | Where-Object DNSName -eq $computername | where-object {($_.SessionState -eq "Active") -or ($_.SessionState -eq "Application")} | Where-Object UserName -eq $XDUser
                                    $item | Add-Member -MemberType NoteProperty -Name "ClientIP" -Value $xdsession.ClientAddress  -Force
                                    "User: " + $username
                                    "Client IP Address: " + $xdsession.ClientAddress
                                    if ($metric.CookedValue -ge 10){
                                        $Script:obj += $item
                                        }
                                    if ($metric.CookedValue -ge $MStoBeConsideredLatent){
                                        "Current Latency is greater than $MStoBeConsideredLatent, taking note of this session"
                                        $Script:XAobjover300 += $item
                                        }
                                    }
                                }
                            }
                        }
                    } 
                }
            else { "Connection to $Computername failed" | LogMeLatency -error }
            }
        }
    else {"No Servers to check" | LogMeLatency -DisplayNormal }
}

Function CreateDesktopLatencyChart {
    "Creating XenDesktop Chart"
    $Graph = $Script:obj | Sort-Object LatencyinMS -descending | Select-Object -Property User,LatencyinMS,DateTime -first 10
    [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
    [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")
    $Chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart 
    $Chart.Width = 800 
    $Chart.Height = 300
    $Chart.Left = 1 
    $Chart.Top = 1
    $Chart.BackColor = [System.Drawing.Color]::FromArgb(236,236,236)
    $ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
    $ChartArea.AxisY.Title = "Milliseconds (ms)"
    $ChartArea.AxisX.LabelStyle.Font = "Arial,11pt"
    $ChartArea.AxisY.TitleFont = "Arial,10pt"
    $ChartArea.AxisX.TitleFont = "Arial,11pt"
    $ChartArea.BackColor = [System.Drawing.Color]::FromArgb(204,204,204)
    $Chart.ChartAreas.Add($ChartArea) 
    [void]$Chart.Series.Add("Data") 
    foreach ($server in $Graph) { 
        if ($server.LatencyinMS -eq 0) {$server.LatencyinMS = "0.1"}
        $dp1 = new-object System.Windows.Forms.DataVisualization.Charting.DataPoint(0, $server.LatencyinMS)
        if ($server.LatencyinMS -ge 200) {$dp1.Color = "Red"}
        elseif ($server.LatencyinMS -ge 150) {$dp1.Color = "Tomato"}
        elseif ($server.LatencyinMS -ge 100) {$dp1.Color = "Orange"}
        elseif ($server.LatencyinMS -ge 75) {$dp1.Color = "Gold"}
        elseif ($server.LatencyinMS -ge 25) {$dp1.Color = "LimeGreen"}
        else {$dp1.Color = "Green"}
        $dp1.AxisLabel = $server.User 
        $Chart.Series["Data"].Points.Add($dp1)
        $ChartArea.AxisX.IsLabelAutoFit = "true"
        #$ChartArea.AxisX.LabelStyle.Angle = "-90"
        $ChartArea.AxisX.LabelStyle.Interval = "1"
        }
    $CurrDatetime = Get-Date
    $title = new-object System.Windows.Forms.DataVisualization.Charting.Title 
    if ($Script:obj.Count -lt 10) {$c = $Chart.Titles.Add( "XenDesktop Top Latency over 10ms")}
    else {$c = $Chart.Titles.Add( "XenDesktop Top 10 Latency over 10ms")}
    $c = $Chart.Titles[0].Font = New-Object System.Drawing.Font("arial",12,[System.Drawing.FontStyle]::Bold)
    $c = $Chart.Titles.Add("Last polled: " +$CurrDatetime + ", XenDesktop Active Sessions over 10ms: " +$Script:obj.Count )
    $c = $Chart.Titles[1].Font = "Arial,10pt"
    if ($dp1 -eq $null) {
        $c = $dp1 = new-object System.Windows.Forms.DataVisualization.Charting.DataPoint(0, 0)
        $c = $dp1.Color = "Green"
        $c = $dp1.AxisLabel = "Users"
        $c = $Chart.Series["Data"].Points.Add()
        $annotation = New-Object System.Windows.Forms.DataVisualization.Charting.TextAnnotation
        $annotation.Text = "There are no Active Sessions over 10ms as of: $CurrDatetime"
        $annotation.AnchorX = 50
        $annotation.AnchorY = 60
        #$annotation.Font = New-Object System.Drawing.Font("Arial", 12,[System.Drawing.FontStyle]::Bold)
        $annotation.Font = "Arial,12pt"
        $annotation.ForeColor = "Blue"
        $chart.Annotations.Add($annotation)
        }
    "Saving Chart Image to $DesktopGraphFile"
    $Chart.SaveImage("$DesktopGraphFile","png")
    ("Copying $DesktopGraphFile to: $HTMLGraphs") | LogMeLatency -displaynormal
    try {copy-item $DesktopGraphFile $HTMLGraphs}
    catch { "Error Copying $DesktopGraphFile to $HTMLGraphs" | LogMeLatency -error
        $_.Exception.Message | LogMeLatency -error
        }
    $date24 | out-file $XDLatencyTimeStampFile
}

Function CreateServerLatencyChart {
    "Creating XenApp Chart"
    $Graph = $Script:obj | Sort-Object LatencyinMS -descending | Select-Object -Property User,LatencyinMS,DateTime -first 10
    [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
    [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")
    $Chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart 
    $Chart.Width = 800 
    $Chart.Height = 300
    $Chart.Left = 1 
    $Chart.Top = 1
    $Chart.BackColor = [System.Drawing.Color]::FromArgb(236,236,236)
    $ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
    $ChartArea.AxisY.Title = "Milliseconds (ms)"
    $ChartArea.AxisX.LabelStyle.Font = "Arial,11pt"
    $ChartArea.AxisY.TitleFont = "Arial,10pt"
    $ChartArea.AxisX.TitleFont = "Arial,11pt"
    $ChartArea.BackColor = [System.Drawing.Color]::FromArgb(204,204,204)
    $Chart.ChartAreas.Add($ChartArea) 
    [void]$Chart.Series.Add("Data") 
    foreach ($server in $Graph) { 
        if ($server.LatencyinMS -eq 0) {$server.LatencyinMS = "0.1"}
        $dp1 = new-object System.Windows.Forms.DataVisualization.Charting.DataPoint(0, $server.LatencyinMS)
        if ($server.LatencyinMS -ge 300) {$dp1.Color = "Red"}
        elseif ($server.LatencyinMS -ge 200) {$dp1.Color = "Firebrick"}
        elseif ($server.LatencyinMS -ge 100) {$dp1.Color = "Orange"}
        elseif ($server.LatencyinMS -ge 75) {$dp1.Color = "Gold"}
        elseif ($server.LatencyinMS -ge 25) {$dp1.Color = "LimeGreen"}
        else {$dp1.Color = "Green"}
        $dp1.AxisLabel = $server.User 
        $Chart.Series["Data"].Points.Add($dp1)
        $ChartArea.AxisX.IsLabelAutoFit = "true"
        #$ChartArea.AxisX.LabelStyle.Angle = "-90"
        $ChartArea.AxisX.LabelStyle.Interval = "1"
        }
    $CurrDatetime = Get-Date
    $title = new-object System.Windows.Forms.DataVisualization.Charting.Title 
    if ($Script:obj.Count -lt 10) {$c = $Chart.Titles.Add( "XenApp Top Latency over 10ms")}
    else {$c = $Chart.Titles.Add( "XenApp Top 10 Latency over 10ms")}
    $c = $Chart.Titles[0].Font = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Bold)
    $c = $Chart.Titles.Add("Last polled: " +$CurrDatetime + ", XenApp Active Sessions over 10ms: " +$Script:obj.Count )
    $c = $Chart.Titles[1].Font = "Arial,10pt"
    if ($dp1 -eq $null) {
        $c = $dp1 = new-object System.Windows.Forms.DataVisualization.Charting.DataPoint(0, 0)
        $c = $dp1.Color = "Green"
        $c = $dp1.AxisLabel = "Users"
        $c = $Chart.Series["Data"].Points.Add()
        $annotation = New-Object System.Windows.Forms.DataVisualization.Charting.TextAnnotation
        $annotation.Text = "There are no Active Sessions over 10ms as of: $CurrDatetime"
        $annotation.AnchorX = 50
        $annotation.AnchorY = 60
        #$annotation.Font = New-Object System.Drawing.Font("Arial", 12,[System.Drawing.FontStyle]::Bold)
        $annotation.Font = "Arial,12pt"
        $annotation.ForeColor = "Blue"
        $chart.Annotations.Add($annotation)
        }
    "Saving Chart Image to $ServerGraphFile"
    $Chart.SaveImage("$ServerGraphFile","png")
    ("Copying $ServerGraphFile to: $HTMLGraphs") | LogMeLatency -displaynormal
    try {copy-item $ServerGraphFile $HTMLGraphs}
    catch { "Error Copying $ServerGraphFile to $HTMLGraphs" | LogMeLatency -error
        $_.Exception.Message | LogMeLatency -error}
    $date24 | out-file $XALatencyTimeStampFile
}

Function CreateXenDesktopUsageChart {
    "Creating XenDesktop Usage Chart"
    $Graph = $Script:XDCurrentGraph 
    $CurrDatetime = Get-Date
    $XDtotal = $Graph.XDTotalCount
    $XDActive = $Graph.XDActiveCount
    $XDDisc = $Graph.XDDisconnectedCount
    [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
    [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")
    $Chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart 
    $Chart.Width = 800 
    $Chart.Height = 200
    $Chart.Left = 1 
    $Chart.Top = 1
    $Chart.BackColor = [System.Drawing.Color]::FromArgb(236,236,236)
    $ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
    $ChartArea.AxisY.Title = "Sessions"
    $ChartArea.AxisX.LabelStyle.Font = "Arial,10pt"
    $ChartArea.AxisY.TitleFont = "Arial,10pt"
    $ChartArea.AxisX.TitleFont = "Arial,11pt"
    $ChartArea.BackColor = [System.Drawing.Color]::FromArgb(204,204,204)
    $Chart.ChartAreas.Add($ChartArea) 
    [void]$Chart.Series.Add("Data") 
    if ($XDTotal -ne 0){
        $dp1 = new-object System.Windows.Forms.DataVisualization.Charting.DataPoint(0, $Graph.XDTotalCount)
        $dp1.Color = "Green"
        $dp1.AxisLabel = "Total ($XDTotal)" 
        $Chart.Series["Data"].Points.Add($dp1)
        $dp2 = new-object System.Windows.Forms.DataVisualization.Charting.DataPoint(0, $Graph.XDActiveCount)
        $dp2.Color = "Green"
        $dp2.AxisLabel = "Active ($XDActive)"
        $Chart.Series["Data"].Points.Add($dp2)
        $dp3 = new-object System.Windows.Forms.DataVisualization.Charting.DataPoint(0, $Graph.XDDisconnectedCount)
        $dp3.Color = "Red"
        $dp3.AxisLabel = "Disconnected ($XDDisc)"
        $Chart.Series["Data"].Points.Add($dp3)
        $difference = [int]$XDtotal - ([int]$XDActive+[int]$XDDisc)
        if (([int]$XDActive+[int]$XDDisc) -lt [int]$XDtotal) {
            $dp4 = new-object System.Windows.Forms.DataVisualization.Charting.DataPoint(0,$difference)
            $dp4.Color = "Blue"
            $dp4.AxisLabel = "Other ($difference)"
            $Chart.Series["Data"].Points.Add($dp4)
            }
        $ChartArea.AxisX.IsLabelAutoFit = "true"
        #$ChartArea.AxisX.LabelStyle.Angle = "-90"
        $ChartArea.AxisX.LabelStyle.Interval = "1"
        }
    else {
        $dp1 = new-object System.Windows.Forms.DataVisualization.Charting.DataPoint(0, 0)
        $dp1.Color = "Green"
        $dp1.AxisLabel = "Total ($XDTotal)"
        $Chart.Series["Data"].Points.Add($dp1)
        $dp2 = new-object System.Windows.Forms.DataVisualization.Charting.DataPoint(0, 0)
        $dp2.Color = "Green"
        $dp2.AxisLabel = "Active (0)"
        $Chart.Series["Data"].Points.Add($dp2)
        $dp3 = new-object System.Windows.Forms.DataVisualization.Charting.DataPoint(0, 0)
        $dp3.Color = "Red"
        $dp3.AxisLabel = "Disconnected (0)"
        $Chart.Series["Data"].Points.Add($dp3)
        $annotation = New-Object System.Windows.Forms.DataVisualization.Charting.TextAnnotation
        $annotation.Text = "There are no Sessions as of: $CurrDatetime"
        $annotation.AnchorX = 50
        $annotation.AnchorY = 65
        #$annotation.Font = [System.Drawing.Font]::new("arial", 12,[System.Drawing.FontStyle]::Bold)
        $annotation.Font = "Arial,12pt"
        $annotation.ForeColor = "Blue"
        $chart.Annotations.Add($annotation)
        }
    $title = new-object System.Windows.Forms.DataVisualization.Charting.Title 
    $c = $Chart.Titles.Add("Total XenDesktop Sessions: " +$XDtotal)
    $c = $Chart.Titles[0].Font = New-Object System.Drawing.Font("arial",12,[System.Drawing.FontStyle]::Bold)
    $c = $Chart.Titles.Add("Last polled: " +$CurrDatetime)
    $c = $Chart.Titles[1].Font = "Arial,10pt"
    "Saving Chart Image to $XDUsageGraphFile"
    $Chart.SaveImage("$XDUsageGraphFile","png")
    ("Copying $XDUsageGraphFile to: $HTMLGraphs")
    try {copy-item $XDUsageGraphFile $HTMLGraphs}
    catch { "Error Copying $XDUsageGraphFile to $HTMLGraphs" | LogMe -error
        $_.Exception.Message | LogMe -error}
}

Function CreateXenAppUsageChart {
    "Creating XenApp Usage Chart"
    $Graph = $Script:XACurrentGraph 
    $CurrDatetime = Get-Date
    $XATotal = $Graph.XATotalCount
    $XAActive = $Graph.XAActiveCount
    $XADisc = $Graph.XADisconnectedCount
    [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
    [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")
    $Chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart 
    $Chart.Width = 800 
    $Chart.Height = 200
    $Chart.Left = 1 
    $Chart.Top = 1
    $Chart.BackColor = [System.Drawing.Color]::FromArgb(236,236,236)
    $ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
    $ChartArea.AxisY.Title = "Sessions"
    $ChartArea.AxisX.LabelStyle.Font = "Arial,10pt"
    $ChartArea.AxisY.TitleFont = "Arial,10pt"
    $ChartArea.AxisX.TitleFont = "Arial,11pt"
    $ChartArea.BackColor = [System.Drawing.Color]::FromArgb(204,204,204)
    $Chart.ChartAreas.Add($ChartArea) 
    [void]$Chart.Series.Add("Data") 
    if ($XATotal -ne 0){
        $dp1 = new-object System.Windows.Forms.DataVisualization.Charting.DataPoint(0, $Graph.XATotalCount)
        $dp1.Color = "Green"
        $dp1.AxisLabel = "Total ($XATotal)"
        $Chart.Series["Data"].Points.Add($dp1)
        $dp2 = new-object System.Windows.Forms.DataVisualization.Charting.DataPoint(0, $Graph.XAActiveCount)
        $dp2.Color = "Green"
        $dp2.AxisLabel = "Active ($XAActive)"
        $Chart.Series["Data"].Points.Add($dp2)
        $dp3 = new-object System.Windows.Forms.DataVisualization.Charting.DataPoint(0, $Graph.XADisconnectedCount)
        $dp3.Color = "Red"
        $dp3.AxisLabel = "Disconnected ($XADisc)"
        $Chart.Series["Data"].Points.Add($dp3)
        $difference = [int]$XAtotal - ([int]$XAActive+[int]$XADisc)
        if (([int]$XAActive+[int]$XADisc) -lt [int]$XAtotal) {
            $dp4 = new-object System.Windows.Forms.DataVisualization.Charting.DataPoint(0,$difference)
            $dp4.Color = "Blue"
            $dp4.AxisLabel = "Other ($difference)"
            $Chart.Series["Data"].Points.Add($dp4)
            }
        }
    else {
        $dp1 = new-object System.Windows.Forms.DataVisualization.Charting.DataPoint(0, 0)
        $dp1.Color = "Green"
        $dp1.AxisLabel = "Total ($XATotal)"
        $Chart.Series["Data"].Points.Add($dp1)
        $dp2 = new-object System.Windows.Forms.DataVisualization.Charting.DataPoint(0, 0)
        $dp2.Color = "Green"
        $dp2.AxisLabel = "Active (0)"
        $Chart.Series["Data"].Points.Add($dp2)
        $dp3 = new-object System.Windows.Forms.DataVisualization.Charting.DataPoint(0, 0)
        $dp3.Color = "Red"
        $dp3.AxisLabel = "Disconnected (0)"
        $Chart.Series["Data"].Points.Add($dp3)
        $annotation = New-Object System.Windows.Forms.DataVisualization.Charting.TextAnnotation
        $annotation.Text = "There are no Sessions as of: $CurrDatetime"
        $annotation.AnchorX = 50
        $annotation.AnchorY = 65
        #$annotation.Font = New-Object System.Drawing.Font("Arial", 12,[System.Drawing.FontStyle]::Bold)
        $annotation.Font = "Arial,12pt"
        $annotation.ForeColor = "Blue"
        $chart.Annotations.Add($annotation)
        }
    $ChartArea.AxisX.IsLabelAutoFit = "true"
    #$ChartArea.AxisX.LabelStyle.Angle = "-90"
    $ChartArea.AxisX.LabelStyle.Interval = "1"
    $title = new-object System.Windows.Forms.DataVisualization.Charting.Title 
    $c = $Chart.Titles.Add( "Total XenApp Sessions: $XATotal")
    $c = $Chart.Titles[0].Font = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Bold)
    $c = $Chart.Titles.Add("Last polled: $CurrDatetime" )
    $c = $Chart.Titles[1].Font = "Arial,10pt"
    "Saving Chart Image to $XAUsageGraphFile"
    $Chart.SaveImage($XAUsageGraphFile,"png")
    ("Copying $XAUsageGraphFile to: $HTMLGraphs")
    try {copy-item $XAUsageGraphFile $HTMLGraphs}
    catch { "Error Copying $XAUsageGraphFile to $HTMLGraphs" | LogMe -error
        $_.Exception.Message | LogMe -error }
}

Function getServerlatency{
    $Servers = Get-BrokerSession -AdminAddress $DDCName | where {($_.Protocol -eq "HDX") -and ($_.SessionState -eq "Active") -and ($_.OSType -match ( $ServerOSTypes -join '|'))}
    CollectXAPerf 
    $XATotalover300 = $Script:XAobjover300.count
    if ($XATotalover300 -ge $NumOfLatencytoCauseAlerts) {
        #Send-Email "Latency Alert for $MonitorName" "There are $NumOfLatencytoCauseAlerts or more sessions with latency over 300ms, there may be a Network or Hypervisor issue."
        $XAEmailLogFileExists = Test-Path $LogXALatencyEmailFile
        if ($XAEmailLogFileExists -eq $true) { $XAEmailFileTime = gc $LogXALatencyEmailFile -TotalCount 1
            $XAAlertTimeSpan = New-TimeSpan $XAEmailFileTime $(Get-Date -format g)
            $XAAlertTimeDifference = $XAAlertTimeSpan.Minutes
            if ($XAAlertTimeDifference -ge $LatencyMinutestoCheck) { 
                Send-Email "XenApp Latency Alert for $MonitorName" "There are $NumOfLatencytoCauseAlerts or more XenApp sessions with latency over $MStoBeConsideredLatent milliseconds, there may be a Network or Hypervisor issue."
                $XDCurrTime = get-date
                $XDCurrTime.ToString("g")  | Out-File $LogXALatencyEmailFile
                }
            elseif (!$XAAlertTimeDifference) { 
                Send-Email "XenApp Latency Alert for $MonitorName" "There are $NumOfLatencytoCauseAlerts or more XenApp sessions with latency over $MStoBeConsideredLatent milliseconds, there may be a Network or Hypervisor issue."
                $XDCurrTime = get-date
                $XDCurrTime.ToString("g")  | Out-File $LogXALatencyEmailFile
                }
            }
        elseif ($XAEmailLogFileExists -eq $false) { 
            Send-Email "XenApp Latency Alert for $MonitorName" "There are $NumOfLatencytoCauseAlerts or more XenApp sessions with latency over $MStoBeConsideredLatent milliseconds, there may be a Network or Hypervisor issue."
            $XDCurrTime = get-date
            $XDCurrTime.ToString("g")  | Out-File $LogXALatencyEmailFile
            }
        }
    $measuremax = $Script:obj | Measure LatencyinMS -Maximum
    if ($measuremax-ne $null){"Max latency: " +$measuremax.Maximum}
    $Script:ChartInterval = $measuremax.Maximum
    CreateServerLatencyChart
    "Exporting to CSV"
    if ($Script:obj) {
        if (Test-Path $Latencyfile) {
            try {$Script:obj | Export-Csv -Path $Latencyfile -NoTypeInformation -Append}
            catch { Error-Message }
            }
        else {
            try { $Script:obj | Export-Csv -Path $Latencyfile -NoTypeInformation }
            Catch { Error-Message }
            }
        }
    if (!$Script:CurrentErrors) {"No errors encountered"}
    if ($Script:CurrentErrors) {write-host "Errors were encountered: `r`n" -ForegroundColor Red
        foreach ($err in $Script:CurrentErrors) {
            write-host "Error $($Script:CurrentErrors.IndexOf($err)):"  -ForegroundColor Red 
            write-host $err -ForegroundColor Red
            }
        }
    if ($UseCustomDB -eq $true) {
        "Connect to SQL Server Database" | LogMeLatency -displaynormal
        $SQLConnection = New-Object System.Data.SqlClient.SqlConnection("Data Source=$SQLServer; Initial Catalog=$DatabaseName; Integrated Security=SSPI")
        $SQLConnection.Open()
        if ($SQLConnection.State -eq "Open") { "Succesfully Connected to SQL Server..." | LogMeLatency -displaynormal } 
        else { "Connection to SQL Server Failed...." | LogMeLatency -error; EXIT }
        "Write Latency info to Latency Table" | LogMeLatency -displaynormal
        $Script:obj | foreach {
            $TempSession=$_.User
            $TempLatencyinMS=$_.LatencyinMS
            $TempDateTime=$_.DateTime
            $TempComputer=$_.Computer
            $TempClientIP=$_.ClientIP
            $CommandText = @"
            INSERT INTO dbo.Latency (Username,LatencyinMS,Timestamp,Computer,ClientIP) 
            VALUES ('$TempSession','$TempLatencyinMS','$TempDateTime','$TempComputer','$TempClientIP')
"@
            RunQuery -CommandText $CommandText
        }
    "Close SQL DB Connection" | LogMeLatency -displaynormal
    $SQLConnection.Close()
    if ($SQLConnection.State -eq "Closed") { "Succesfully Closed connection to SQL Server..."  | LogMeLatency -displaynormal } 
    else { "Failed to close Connection to SQL Server..." | LogMeLatency -displaynormal }
    }
}

Function getDesktoplatency{
    $Servers = Get-BrokerSession -AdminAddress $DDCName | where {($_.SessionState -eq "Active") -and ($_.OSType -match ( $DesktopOSTypes -join '|'))}
    CollectXDPerf 
    [int]$XDTotalover300 = $Script:XDobjover300.count
    if ($XDTotalover300 -ge $NumOfLatencytoCauseAlerts) {
        #Send-Email "Latency Alert for $MonitorName" "There are $NumOfLatencytoCauseAlerts or more sessions with latency over 300ms, there may be a Network or Hypervisor issue."
        $XDEmailLogFileExists = Test-Path $LogXDLatencyEmailFile
        if ($XDEmailLogFileExists -eq $true) { $XDEmailFileTime = gc $LogXDLatencyEmailFile -TotalCount 1
            $XDAlertTimeSpan = New-TimeSpan $XDEmailFileTime $(Get-Date -format g)
            $XDAlertTimeDifference = $XDAlertTimeSpan.Minutes
            if ($XDAlertTimeDifference -ge $LatencyMinutestoCheck) { 
                Send-Email "XenDesktop Latency Alert for $MonitorName" "There are $NumOfLatencytoCauseAlerts or more XenDesktop sessions with latency over $MStoBeConsideredLatent milliseconds, there may be a Network or Hypervisor issue."
                $XDCurrTime = get-date
                $XDCurrTime.ToString("g")  | Out-File $LogXDLatencyEmailFile
                }
            elseif (!$XDAlertTimeDifference) { 
                Send-Email "XenDesktop Latency Alert for $MonitorName" "There are $NumOfLatencytoCauseAlerts or more XenDesktop sessions with latency over $MStoBeConsideredLatent milliseconds, there may be a Network or Hypervisor issue."
                $XDCurrTime = get-date
                $XDCurrTime.ToString("g")  | Out-File $LogXDLatencyEmailFile
                }
            }
        elseif ($XDEmailLogFileExists -eq $false) { 
            Send-Email "XenDesktop Latency Alert for $MonitorName" "There are $NumOfLatencytoCauseAlerts or more XenDesktop sessions with latency over $MStoBeConsideredLatent milliseconds, there may be a Network or Hypervisor issue."
            $XDCurrTime = get-date
            $XDCurrTime.ToString("g")  | Out-File $LogXDLatencyEmailFile
            }
        }
    $measuremax = $Script:obj | Measure LatencyinMS -Maximum
    "Max latency: " +$measuremax.Maximum
    $Script:ChartInterval = $measuremax.Maximum
    CreateDesktopLatencyChart
    "Exporting to CSV"
    if ($Script:obj) {
        if (Test-Path $Latencyfile) {
            try {$Script:obj | Export-Csv -Path $Latencyfile -NoTypeInformation -Append}
            catch { Error-Message }
            } 
        else {
            try { $Script:obj | Export-Csv -Path $Latencyfile -NoTypeInformation }
            Catch {Error-Message }
            }
        }
    if (!$Script:CurrentErrors) {"No errors encountered"}
    if ($Script:CurrentErrors) {write-host "Errors were encountered: `r`n" -ForegroundColor Red
        foreach ($err in $Script:CurrentErrors) {
            write-host "Error $($Script:CurrentErrors.IndexOf($err)):"  -ForegroundColor Red 
            write-host $err -ForegroundColor Red
            }
        }
    if ($UseCustomDB -eq $true) {
        "Connect to SQL Server Database" | LogMeLatency -displaynormal
        $SQLConnection = New-Object System.Data.SqlClient.SqlConnection("Data Source=$SQLServer; Initial Catalog=$DatabaseName; Integrated Security=SSPI")
        $SQLConnection.Open()
        if ($SQLConnection.State -eq "Open") { "Succesfully Connected to SQL Server..." | LogMeLatency -displaynormal } 
        else { "Connection to SQL Server Failed...." | LogMeLatency -error; EXIT }
        "Write Latency info to Latency Table" | LogMeLatency -displaynormal
        $Script:obj | foreach {
            $TempSession=$_.User
            $TempLatencyinMS=$_.LatencyinMS
            $TempDateTime=$_.DateTime
            $TempComputer=$_.Computer
            $TempClientIP=$_.ClientIP
            $CommandText = @"
            INSERT INTO dbo.Latency (Username,LatencyinMS,Timestamp,Computer,ClientIP) 
            VALUES ('$TempSession','$TempLatencyinMS','$TempDateTime','$TempComputer','$TempClientIP')
"@
            RunQuery -CommandText $CommandText
        }
    "Close SQL DB Connection" | LogMeLatency -displaynormal
    $SQLConnection.Close()
    if ($SQLConnection.State -eq "Closed") { "Succesfully Closed connection to SQL Server..."  | LogMeLatency -displaynormal } 
    else { "Failed to close Connection to SQL Server..." | LogMeLatency -displaynormal }
    }
}

Function CreateHTML{
$head = @"
<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'/>
<meta http-equiv="refresh" content="60"/>
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate"/>
<meta http-equiv="Pragma" content="no-cache"/>
<meta http-equiv="Expires" content="0"/>
<link rel="shortcut icon" href="ecg.ico" type="image/x-icon"/>
<link rel="shortcut icon" href="favicon.ico" />
<title>$MonitorName</title>
<style type="text/css">
<!--
td { font-family: Lao UI;font-size: 11px;border-top: 1px solid #999999;border-right: 1px solid #999999;border-bottom: 1px solid #999999;border-left: 1px solid #999999;
    padding-top: 0px;padding-right: 0px;padding-bottom: 0px;padding-left: 0px;overflow: hidden;
    }
th { font-family: Tahoma;font-size: 11px;font-weight:bold;border-top: 1px solid #999999;border-right: 1px solid #999999;border-bottom: 1px solid #999999;border-left: 1px solid #999999;
    padding-top: 0px;padding-right: 0px;padding-bottom: 0px;padding-left: 0px;overflow: hidden;color:#8a0808;
    }
.outertable { padding-top: 0px;padding-right: 0px;padding-bottom: 0px;padding-left: 0px;overflow: hidden;color:#000000;width: 100%;border: 1px solid #000000;cellspacing: 0; }
.outertabletd { padding-top: 0px;padding-right: 0px;padding-bottom: 0px;padding-left: 0px;overflow: hidden;width: 100%; border: 0px solid #000000; cellspacing: 0; }
.header { width: 100%;border-width: 0; }
.headertr { width: 100%; background-image: url('./images/$ecg'); background-repeat: no-repeat; background-position: left center; background-size: cover; cellspacing: 0; }
.headertdr { width: 50%;font-family: Tahoma;font-size: 40px;font-weight:bold;text-align: center;padding-top: 0px;padding-right: 0px;padding-bottom: 0px;padding-left: 0px;border: 0;
    overflow: hidden;text-shadow:0px 0px 14px #FF0000;align:center;valign:middle;height: 100px;cellspacing: 0;
    }
.headertdl { width: 50%;font-family: Tahoma;font-size: 40px;font-weight:bold;text-align: center;padding-top: 0px;padding-right: 0px;padding-bottom: 0px;padding-left: 0px;border: 0;
    overflow: hidden;text-shadow:0px 0px 14px #FF0000;align:center;valign:middle;height: 100px;
    }
.thtopinfo {background-color: #ECECEC;font-family: Tahoma;font-size: 14px;font-weight:bold;text-align: center;padding-top: 0px;padding-right: 0px;padding-bottom: 0px;padding-left: 0px;
    overflow: hidden;color:#8A0808;width:33%;
    }
.tdcurrinfo {width: 25%;text-align: center;font-family: Tahoma;font-size: 16px;font-weight:bold;color: #003399;}
.tdcurrinfoerror {width: 25%;text-align: center;font-family: Tahoma;font-size: 16px;font-weight:bold;color: #FFFFFF;background-color: #DA0101;}

.trcurrinfo {background-color: #CCCCCC;}
.trcurrinfoodd {background-color: #ffffff;}
.tdreportsheader {background-color: #ffffff;font-family: Tahoma;font-size: 14px;font-weight:bold;text-align: center;border-top: 1px solid #000000;border-right: 1px solid #000000;
    border-bottom: 1px solid #000000;border-left: 1px solid #000000;padding-top: 0px;padding-right: 0px;padding-bottom: 0px;padding-left: 0px;overflow: hidden;color:#8A0808;
    }
.tdreports {width: 20%;text-align: center;font-family: Tahoma;font-size: 14px;font-weight:bold;color: #003399;}
.tdreportsrowdesc {width: 20%;text-align: center;font-family: Tahoma;font-size: 14px;font-weight:bold;color: #ffffff;background-color:#636363; }
.tblxainfo {width: 100%;}
.tdxaservercol {text-align: center;background-color: #CCCCCC;font-family: Tahoma;font-size: 10px;font-weight:bold;color: #003399; }
.tdxainfo {text-align: center;background-color: #8BC69B;font-family: Tahoma;font-size: 10px;font-weight:bold;color: #003399; }
.tdxainfoerror {text-align: center;background-color: #DA0101;font-family: Tahoma;font-size: 10px;font-weight:bold;color: #FFFFFF;}
.tdxainfowarn {text-align: center;background-color: #FFC300;font-family: Tahoma;font-size: 10px;font-weight:bold;color: #000000;}
.thxainfo {background-color: #797979;font-family: Tahoma;font-size: 11px;font-weight:bold;border-top: 1px solid #000000;border-right: 1px solid #000000;border-bottom: 1px solid #000000;
    border-left: 1px solid #000000;padding-top: 0px;padding-right: 0px;padding-bottom: 0px;padding-left: 0px;overflow: hidden;color:#FFFFFF;
    }
body {margin-left: 5px;margin-top: 5px;margin-right: 5px;margin-bottom: 10px;}
table {table-layout:fixed;border: thin solid #FFFFFF;}
.shadow {height: 1em;filter: Glow(Color=#000000,Direction=135,Strength=5);}
a:link, a:visited {color: #003399;text-decoration: underline;}
a:hover {color: #000000text-decoration: underline;background-color: yellow;}
a:active {color: #000000;text-decoration: underline;background-color: lightgreen;}
-->
</style>
</head>
<body>
"@
$outertable = "<table class='outertable' width='100%'><tr><td class='outertabletd'>"
$logos = @"
<table class="header" cellspacing="0" cellpadding="0">
    <tr class="headertr">
        <td class="headertdl"><p class="shadow">$MonitorName</p></td>
        <td class="headertdr"><img src="./images/$LogoImage"/></td>
    </tr>
</table>
"@
$pageinfo = @"
<table width='100%'>
    <tr>
        <th class='thtopinfo'><script type="text/javascript">
            <!--
            var currentTime = new Date()
            var month = currentTime.getMonth() + 1
            var day = currentTime.getDate()
            var year = currentTime.getFullYear()
            var hours = currentTime.getHours()
            var minutes = currentTime.getMinutes()
            if (minutes < 10){ minutes = "0" + minutes }
            var theTime = month + "/" + day + "/" + year + " " + hours + ":" + minutes + " "
            currTime = theTime
            var lastUpdate = "$date24" 
            var difference = Date.parse(currTime) - Date.parse(lastUpdate)
            var resultInMinutes = Math.round(difference / 60000)
            if(resultInMinutes > "$UpdateInterval"){ document.write("Farm Last Queried: <strong><font face='Tahoma' color='#003399' size='2'>" + lastUpdate + "<font face='Tahoma' color='#f90000' size='2'><br/>ERROR: Check Script, no update for " + resultInMinutes + " Minutes") } 
            else{ document.write("Farm Last Queried: <strong><font face='Tahoma' color='#003399' size='2'>" + lastUpdate) }
            //-->
            </script>
        </th>
        <th class='thtopinfo'>Page Last Refresfed: <strong><font face='Tahoma' color='#003399' size='2'>
        <script type="text/javascript">
            <!--
            var currentTime = new Date()
            var month = currentTime.getMonth() + 1
            var day = currentTime.getDate()
            var year = currentTime.getFullYear()
            var hours = currentTime.getHours()
            var minutes = currentTime.getMinutes()
            if (minutes < 10){ minutes = "0" + minutes }
            var theTime = month + "/" + day + "/" + year + " " + hours + ":" + minutes + " "
            document.write(theTime)
            //-->
            </script>
            </font></strong>
        </th>
        <th class='thtopinfo'>Auto-Refresh in <strong><font face='Tahoma' color='#003399' size='2'><span id="CDTimer">180</span> secs.
        </font></strong>
        <script type="text/javascript">
            /*<![CDATA[*/
            var TimerVal = 60;
            var TimerSPan = document.getElementById("CDTimer");
            function CountDown(){
                setTimeout( "CountDown()", 1000 );
                TimerSPan.innerHTML=TimerVal;
                TimerVal=TimerVal-1;
                } CountDown() /*]]>*/
                </script>
                
        </th>
    </tr>
</table>
"@
$xainfolastpolledTS = gc $XAInfoTimeStampFile 
$xainfolastpolled = @"
<table width='100%'>
    <tr>
        <th class='thtopinfo'><script type="text/javascript">
            <!--
            var currentTime = new Date()
            var month = currentTime.getMonth() + 1
            var day = currentTime.getDate()
            var year = currentTime.getFullYear()
            var hours = currentTime.getHours()
            var minutes = currentTime.getMinutes()
            if (minutes < 10){ minutes = "0" + minutes }
            var theTime = month + "/" + day + "/" + year + " " + hours + ":" + minutes + " "
            currTime = theTime
            var lastUpdate = "$xainfolastpolledTS" 
            var difference = Date.parse(currTime) - Date.parse(lastUpdate)
            var resultInMinutes = Math.round(difference / 60000)
            if(resultInMinutes > "$UpdateInterval"){ document.write("XenApp Servers Last Queried: <strong><font face='Tahoma' color='#003399' size='2'>" + lastUpdate + "<font face='Tahoma' color='#f90000' size='2'><br/>ERROR: Check Script, no XenApp Server Info update for " + resultInMinutes + " Minutes") } 
            else{ document.write("XenApp Servers Last Queried: <strong><font face='Tahoma' color='#003399' size='2'>" + lastUpdate) }
            //-->
            </script>
        </th>
    </tr>
</table>
"@
$current = @"
<table width='100%'>
    <tr class='trcurrinfo'>
        <td class='tdcurrinfo'>Total XenApp and XenDesktop Sessions: <strong><font face='Tahoma' color='#8A0808' size='4'>$TotalSessions</font></strong></td>
        <td class='tdcurrinfo'>Citrix License Usage: <strong><font face='Tahoma' color='#8A0808' size='4'>$Licenses</font></strong></td>
    </tr>
    <tr class='trcurrinfo'>
        $FailedConnCountCell
        $CompUnregisteredCountCell
    </tr>
</table>
"@
$xdlatencylastpolled = gc $XDLatencyTimeStampFile
$xalatencylastpolled = gc $XALatencyTimeStampFile
$UsageCharts = @"
<table width='100%'>
    <tr bgcolor='#ECECEC'>
        <td class='tdcurrinfo'><a title="Click here for XenDesktop Current Usage" href="./current/xdcurrent.html" onclick="window.open('./current/xdcurrent.html', 'xdcu', 'width=1000,height=750,top=25,left=25'); return false;"><img src="./graphs/$XDUsageGraphFilename"/></a></td>
        <td class='tdcurrinfo'><a title="Click here for XenApp Current Usage" href="./current/xacurrent.html" onclick="window.open('./current/xacurrent.html', 'xacu', 'width=1000,height=750,top=25,left=25'); return false;"><img src="./graphs/$XAUsageGraphFilename"/></a></td>
    </tr>
"@
if ($CollectLatency -eq $true){
$LatencyChart1 = @"
    <tr bgcolor='#ECECEC'>
        <td align='center' valign="middle">
        <script type="text/javascript">
            <!--
            var currentTime = new Date()
            var month = currentTime.getMonth() + 1
            var day = currentTime.getDate()
            var year = currentTime.getFullYear()
            var hours = currentTime.getHours()
            var minutes = currentTime.getMinutes()
            if (minutes < 10){ minutes = "0" + minutes }
            var theTime = month + "/" + day + "/" + year + " " + hours + ":" + minutes + " "
            currTime = theTime
            var lastUpdate = "$xdlatencylastpolled" 
            var difference = Date.parse(currTime) - Date.parse(lastUpdate)
            var resultInMinutes = Math.round(difference / 60000)
            if(resultInMinutes > "$UpdateInterval"){ document.write("<strong><font face='Tahoma' color='#f90000' size='2'><br/>ERROR: Check Script, no XenDesktop Latency Update for " + resultInMinutes + " Minutes </font></strong>") }
            //-->
            </script>
"@
if ($UseCustomDB -eq $true) { $SSRSLink = "<a title=`"Click here for latency historical reports`" href=`"$SSRSURL`" target=`"_blank`"><img src=`"./graphs/$DesktopGraphFilename`"/></a>" }
else { $SSRSLink = "<img src=`"./graphs/$ServerGraphFilename`"/></td>"  }
$LatencyChart1end = "$SSRSLink</td>"

$LatencyChart2 = @"
        <td align='center' valign="middle">
        <script type="text/javascript">
            <!--
            var currentTime = new Date()
            var month = currentTime.getMonth() + 1
            var day = currentTime.getDate()
            var year = currentTime.getFullYear()
            var hours = currentTime.getHours()
            var minutes = currentTime.getMinutes()
            if (minutes < 10){ minutes = "0" + minutes }
            var theTime = month + "/" + day + "/" + year + " " + hours + ":" + minutes + " "
            currTime = theTime
            var lastUpdate = "$xalatencylastpolled" 
            var difference = Date.parse(currTime) - Date.parse(lastUpdate)
            var resultInMinutes = Math.round(difference / 60000)
            if(resultInMinutes > "$UpdateInterval"){ document.write("<strong><font face='Tahoma' color='#f90000' size='2'><br/>ERROR: Check Script, no XenApp Latency Update for " + resultInMinutes + " Minutes </font></strong>") }
            //-->
            </script>
"@
if ($UseCustomDB -eq $true) { $SSRSLink = "<a title=`"Click here for latency historical reports`" href=`"$SSRSURL`" target=`"_blank`"><img src=`"./graphs/$ServerGraphFilename`"/></a>" }
else { $SSRSLink = "<img src=`"./graphs/$ServerGraphFilename`"/>"  }
$LatencyChart2end = "$SSRSLink</td></tr></table>"
}
$UsageReports = @"
<table width='100%'>
    <tr bgcolor='#ffffff'>
        <td class='tdreportsheader' colspan='4' width=100% align='center' valign="middle"><font face='Tahoma' color='#8A0808' size='3'><strong>$MonitorName Historical Reports</strong></font></td>
    </tr>
    <tr class='trcurrinfo'>
        <td class='tdreportsrowdesc'>XenDesktop Usage</td>
        <td class='tdreportsrowdesc'>XenApp Usage</td>
        <td class='tdreportsrowdesc'>Published Applications Usage</td>
        <td class='tdreportsrowdesc'>Citrix Receiver Versions</td>
    </tr>
    <tr class='trcurrinfo'>
        <td class='tdreports'><a href="./reports/xdyesterday.html" onclick="window.open('./reports/xdyesterday.html', 'yesterday', 'width=1000,height=725,top=25,left=25'); return false;">Yesterday</a></td>
        <td class='tdreports'><a href="./reports/xayesterday.html" onclick="window.open('./reports/xayesterday.html', 'yesterday', 'width=1000,height=725,top=25,left=25'); return false;">Yesterday</a></td>
        <td class='tdreports'><a href="./reports/XAPAyesterday.html" onclick="window.open('./reports/XAPAyesterday.html', 'yesterday', 'width=1000,height=725,top=25,left=25'); return false;">Yesterday</a></td>
        <td class='tdreports'><a href="./reports/clientver7days.html" onclick="window.open('./reports/clientver7days.html', 'verlast7', 'width=1000,height=725,top=25,left=25'); return false;">Last 7 Days</a></td>
    </tr>
    <tr class='trcurrinfo'>
        <td class='tdreports'><a href="./reports/xdlast7days.html" onclick="window.open('./reports/xdlast7days.html', 'last7', 'width=1000,height=725,top=25,left=25'); return false;">Last 7 Days</a></td>
        <td class='tdreports'><a href="./reports/xalast7days.html" onclick="window.open('./reports/xalast7days.html', 'last7', 'width=1000,height=725,top=25,left=25'); return false;">Last 7 Days</a></td>
        <td class='tdreports'><a href="./reports/XAPAlast7days.html" onclick="window.open('./reports/XAPAlast7days.html', 'last7', 'width=1000,height=725,top=25,left=25'); return false;">Last 7 Days</a></td>
        <td class='tdreports'><a href="./reports/clientver30days.html" onclick="window.open('./reports/clientver30days.html', 'verlast30', 'width=1000,height=725,top=25,left=25'); return false;">Last 30 Days</a></td>
    </tr>
    <tr class='trcurrinfo'>
        <td class='tdreports'><a href="./reports/xdlast30days.html" onclick="window.open('./reports/xdlast30days.html', 'last30', 'width=1000,height=725,top=25,left=25'); return false;">Last 30 Days</a></td>
        <td class='tdreports'><a href="./reports/xalast30days.html" onclick="window.open('./reports/xalast30days.html', 'last30', 'width=1000,height=725,top=25,left=25'); return false;">Last 30 Days</a></td>
        <td class='tdreports'><a href="./reports/XAPAlast30days.html" onclick="window.open('./reports/XAPAlast30days.html', 'last30', 'width=1000,height=725,top=25,left=25'); return false;">Last 30 Days</a></td>
        <td class='tdreports'></td>
    </tr>
    <tr class='trcurrinfo'>
        <td class='tdreports'><a href="./reports/XDDGs.html" onclick="window.open('./reports/XDDGs.html', 'xddgs', 'width=1400,height=725,top=25,left=25'); return false;">Delivery Groups</a></td>
        <td class='tdreports'><a href="./reports/XADDGs.html" onclick="window.open('./reports/XADDGs.html', 'xadgs', 'width=1000,height=725,top=25,left=25'); return false;">Delivery Groups</a></td>
        <td class='tdreports'><a href="./reports/XAApps.html" onclick="window.open('./reports/XAApps.html', 'xapa', 'width=1550,height=725,top=25,left=20'); return false;">All Applications</a></td>
        <td class='tdreports'></td>
    </tr>
    <tr class='trcurrinfo'>
        <td class='tdreports'><a href="./reports/XDIdle.html" onclick="window.open('./reports/XDIdle.html', 'dgsnotused', 'width=1000,height=725,top=25,left=25'); return false;">Not used in last 60 days</a></td>
        <td class='tdreports'><a href="./reports/XAIdle.html" onclick="window.open('./reports/XAIdle.html', 'xadgsnotused', 'width=1000,height=725,top=25,left=25'); return false;">Not used in last 60 days</a></td>
        <td class='tdreports'></td>
        <td class='tdreports'></td>
    </tr>
</table>
<table width='100%'>
    <tr bgcolor='#ECECEC'>
        <td><br/><br/><p style="text-align: center;"><a title="Logon to Citrix Director for more details on $MonitorName" href="$DirectorURL" target="_blank"><img src="./images/Director.png"/></a></p>
"@
$footer=("<br/><br/><font face='HP Simplified' color='#003399' size='2'><br/><em>Page last updated on {0}.<br/>Script Hosted on server {3}.<br/>Script Path: {4}</em></font><br/><br/><br/></td></tr></table></td></tr></table>" -f (Get-Date -displayhint date),$env:userdomain,$env:username,$env:COMPUTERNAME,$currentDir) 
$xainfo = gc "XAInfo.txt"
$HTMLFile = ""
$HTMLFile = $head
$HTMLFile = $HTMLFile + $outertable
$HTMLFile = $HTMLFile + $logos
$HTMLFile = $HTMLFile + $pageinfo
$HTMLFile = $HTMLFile + $xainfolastpolled
$HTMLFile = $HTMLFile + $xainfo
$HTMLFile = $HTMLFile + $current
$HTMLFile = $HTMLFile + $UsageCharts
if ($CollectLatency -eq $true){
$HTMLFile = $HTMLFile + $LatencyChart1
$HTMLFile = $HTMLFile + $LatencyChart1end
$HTMLFile = $HTMLFile + $LatencyChart2
$HTMLFile = $HTMLFile + $LatencyChart2end
}
$HTMLFile = $HTMLFile + $UsageReports
$HTMLFile = $HTMLFile + $footer
$HTMLFile = $HTMLFile + "</body></html>" | Out-File $HTMLFilePath$HTMLFileName
"Copying HTML File and images to $HTMLServer" | LogMe -displaynormal

try {copy-item $HTMLFilePath$HTMLfileName $HTMLServer$HTMLFilename -force}
    catch { "Error Copying $HTMLFilePath$HTMLfileName to $HTMLServer$HTMLFilename" | LogMe -error
        $_.Exception.Message | LogMe -error}

try {copy-item $ScriptFilePath$LogoImage $HTMLImages -force}
    catch { "Error Copying $ScriptFilePath$LogoImage to $HTMLImages" | LogMe -error
        $_.Exception.Message | LogMe -error}

try {copy-item $ScriptFilePath$ecg $HTMLImages -force}
    catch { "Error Copying $ScriptFilePath$ecg to $HTMLImages" | LogMe -error
        $_.Exception.Message | LogMe -error}

try {copy-item "Director.png" $HTMLImages -force}
    catch { "Error Copying Director.png to $HTMLImages" | LogMe -error
        $_.Exception.Message | LogMe -error}

try {copy-item $favicon $HTMLServer$fav -force}
    catch { "Error Copying $favicon to $HTMLServer$fav" | LogMe -error
        $_.Exception.Message | LogMe -error}

}

Function Send-Email($EmailSubject,$EmailBody){
    Write-host "Sending Email"
    Send-MailMessage -From $emailFrom -To $emailTo -Subject $EmailSubject -BodyAsHtml $EmailBody -SmtpServer $smtpServer
}

#endregion Script Functions

#region Parameters
if ($paramarg -eq "FirstRun") {
    LoadCitrixSnapin
    "Running XenDesktop Usage Report for Yesterday"
    GetXDUsageReports
    "Running XenDesktop Usage Report for the last 7 days"
    GetXDUsageReports -dur w
    "Running XenDesktop Usage Report for the last 30 days"
    GetXDUsageReports -dur m
    "Running XenApp Usage Report for Yesterday"
    GetXAUsageReports
    "Running XenApp Usage Report for the last 7 days"
    GetXAUsageReports -dur w
    "Running XenApp Usage Report for the last 30 days"
    GetXAUsageReports -dur m
    "Running Published Application Usage Report for Yesterday"
    GetPubAppUsageReport
    "Running Published Application Usage Report for the last 7 days"
    GetPubAppUsageReport -dur w
    "Running Published Application Usage Report for the last 30 days"
    GetPubAppUsageReport -dur m
    "Running Client Version Report for the last 30 days"
    GetClientVersionsReport
    "Running Client Version Report for the last 7 days"
    GetClientVersionsReport -dur w
    "Running Idle Desktop Report for the last 60 days"
    GetIdleDesktops
    "Running Idle Server Report for the last 60 days"
    GetIdleServers
    "Getting Inventory of XenApp Desktop Delivery Groups"
    GetXenAppDGs
    "Getting Inventory of XenDesktop Desktop Delivery Groups"
    GetXenDesktopDGs
    "Getting Inventory of XenApp Published Apps"
    GetXenAppPubApps
    "Getting Server Latency Counters"
    getServerlatency
    "Getting Desktop Latency Counters"
    getdesktoplatency
    "Collecting XenApp Server info"
    GetXAServerInfo
    }
elseif ($paramarg -eq "RunReports") {
    LoadCitrixSnapin
    "Running XenDesktop Usage Report for Yesterday"
    GetXDUsageReports
    "Running XenDesktop Usage Report for the last 7 days"
    GetXDUsageReports -dur w
    "Running XenDesktop Usage Report for the last 30 days"
    GetXDUsageReports -dur m
    "Running XenApp Usage Report for Yesterday"
    GetXAUsageReports
    "Running XenApp Usage Report for the last 7 days"
    GetXAUsageReports -dur w
    "Running XenApp Usage Report for the last 30 days"
    GetXAUsageReports -dur m
    "Running Published Application Usage Report for Yesterday"
    GetPubAppUsageReport
    "Running Published Application Usage Report for the last 7 days"
    GetPubAppUsageReport -dur w
    "Running Published Application Usage Report for the last 30 days"
    GetPubAppUsageReport -dur m
    "Running Client Version Report for Yesterday"
    GetClientVersionsReport
    "Running Client Version Report for the last 7 days"
    GetClientVersionsReport -dur w
    "Running Idle Desktop Report for the last 60 days"
    GetIdleDesktops
    "Running Idle Servers Report for the last 60 days"
    GetIdleServers
    "Getting Inventory of XenApp Desktop Delivery Groups"
    GetXenAppDGs
    "Getting Inventory of XenDesktop Desktop Delivery Groups"
    GetXenDesktopDGs
    "Getting Inventory of XenApp Published Apps"
    GetXenAppPubApps
    }
elseif ($paramarg -eq "XDLatency") {
    LoadCitrixSnapin
    CreatePreviousLatencyRunLog
    "Getting Desktop Latency Counters"
    getdesktoplatency
    }
elseif ($paramarg -eq "XALatency") {
    LoadCitrixSnapin
    CreatePreviousLatencyRunLog
    "Getting Server Latency Counters"
    getServerlatency
    }
elseif ($paramarg -eq "Current") {
    "Loading Citrix Snapin"
    LoadCitrixSnapin
    }
elseif ($paramarg -eq "XAInfo") { 
    LoadCitrixSnapin
    "Collecting XenApp Server info"
    GetXAServerInfo
    }
elseif ($paramarg -eq "Test") {
    LoadCitrixSnapin
    "Running XenDesktop Usage Report for Yesterday"
    GetXDUsageReports
 }
elseif (!$paramarg) { ArgHelp; Break }
else { ArgHelp; Break }

#endregion Parameters

#region Default Functions to run, regardless of Argument

"Getting Current info"
"Checking XenApp Usage"
GetXACurrentUsage
"Checking Current XenDesktop Usage"
GetXDCurrentUsage
"Getting Current Citrix Licenses in Use"
get-licenses
$TotalSessions = $Script:xdcurrcount + $Script:xacurrcount
$Licenses = $Script:citrixlicenses
"Checking Failed Logons for the past $FailedConnMinutestoCheck minutes"
$FailedConnectionsLastMinutes = Get-BrokerConnectionLog -AdminAddress $DDCName  -MaxRecordCount 2500 -Filter {BrokeringTime -gt $FailedConnectionTimeSpanMinutes -and ConnectionFailureReason -ne 'None'}
[int]$FailedConnCount = $FailedConnectionsLastMinutes.count
If ($FailedConnCount -ge 5) {
    $FailedConnCountCell = "<td class='tdcurrinfoerror'>Session connection failures for the last $FailedConnMinutestoCheck minutes: <strong><font face='Tahoma' color='#FFFFFF' size='4'>$FailedConnCount</font></strong></td>"
    $EmailLogFileExists = Test-Path $LogEmailFile
    if ($EmailLogFileExists -eq $true) { $EmailFileTime = gc $LogEmailFile -TotalCount 1
        $AlertTimeSpan = New-TimeSpan $EmailFileTime $(Get-Date -format g)
        $AlertTimeDifference = $AlertTimeSpan.Minutes
        if ($AlertTimeDifference -ge $FailedConnMinutestoCheck) { 
            Send-Email "Failed Session Count Alert for $MonitorName" "Failed Session Count equals $FailedConnCount for the last $FailedConnMinutestoCheck minutes"
            $CurrTime = get-date
            $CurrTime.ToString("g")  | Out-File $LogEmailFile
            }
        elseif (!$AlertTimeDifference) { 
            Send-Email "Failed Session Count Alert for $MonitorName" "Failed Session Count equals $FailedConnCount for the last $FailedConnMinutestoCheck minutes"
            $CurrTime = get-date
            $CurrTime.ToString("g")  | Out-File $LogEmailFile
            }
        }
    elseif ($EmailLogFileExists -eq $false) { 
        Send-Email "Failed Session Count Alert for $MonitorName" "Failed Session Count equals $FailedConnCount for the last $FailedConnMinutestoCheck minutes"
        $CurrTime = get-date
        $CurrTime.ToString("g")  | Out-File $LogEmailFile
        }
    }
else { $FailedConnCountCell = "<td class='tdcurrinfo'>Session connection failures for the last $FailedConnMinutestoCheck minutes: <strong><font face='Tahoma' color='#003399' size='4'>$FailedConnCount</font></strong></td>" }
"Checking for unregistered Computers that are not in Maintenance mode"
$CompUnregsitered = Get-BrokerMachine | where {($_.RegistrationState -eq "Unregistered") -and ($_.InMaintenanceMode -ne "True")}
[int]$CompUnregisteredCount = $CompUnregsitered.count
If ($CompUnregisteredCount -ge 1) { $CompUnregisteredCountCell = "<td class='tdcurrinfoerror'>Computers Unregistered: <strong><font face='Tahoma' color='#FFFFFF' size='4'>$CompUnregisteredCount</font></strong></td>" }
else { $CompUnregisteredCountCell = "<td class='tdcurrinfo'>Computers Unregistered: <strong><font face='Tahoma' color='#003399' size='4'>$CompUnregisteredCount</font></strong></td>" }
"Creating HTML"
CreateHTML

#endregion Default Functions to run, regardless of Argument