#########################################################
# Adjustable variables (5)
# Manual safe guard to ensure scope is limited
# Set the name of the Active Directory domain to query, eg $DomainDNS = "danovich.com.au"
# $DomainDNS value is ignored if the $SpecificDC value is set to anything after from null
$DomainDNS = "contoso.com"
# If this value is set, only a specific Domain Controller is queried
# For example, to only query TESTDC01, set $SpecificDC = "TESTDC01"
# Otherwise set $SpecificDC = $null and all Domain Controllers will be discovered and queried
$SpecificDC = $null
# Set the timeout in milliseconds for ping command, eg $Timeout = 500
# Some enviornments may need a larger value
$Timeout = 5000
# Set the type of output required, either:
#  "Excel" - Nicely formatted Excel spreadsheet formatting - requires Excel installed on host
#  "CSV" - Basic text output to CSV file for further manipulation
$OutputFormat = "CSV"
# If using CSV, specify the name and location of the CSV file to be created, eg $FileLocation = "c:tempExportDCs.csv"
$FileLocation = "c:\temp\ExportDCs.csv"
##########################################################
# AUTHOR: blog.danovich.com.au
# DATE:  16/05/2013
# NAME:  Domain_Controllers_Audit.ps1
# VERSION: 1.7
# PURPOSE: Audit Active Directory Domain Controllers
# COMMENT: 1.0 Initial release after testing
#   1.1 Adjusted $DomainControllers = Get-ADDomainController -Domain $DomainDNS query
#   1.2 Added support for Windows 2003 Server queries - previously was only 2008 OS and above
#       Ability to query only one particular Domain Controller
#   1.3 Fixed incorrect CPU core counting
#   1.4 Added WMI connectivity check
#   1.5 Added checks for disk space, SCCM & SCOM agents
#   1.6 Added progress bar
#   1.7 Fixed incorrect virtual machine query
##########################################################
# REQUIREMENTS:
# To run the PowerShell script, the correct Execution Policy level must be set (Set-ExecutionPolicy)
# The user account running this script must have permission to:
#               - Remotely query WMI on Domain Controllers
#               - Query Active Directory attributes of Domain Contollers
# Network connectivity from the host where script is running to all Domain Controllers in the domain including:
#               - ICMP (for ping)
#               - TCP & UDP ports for remote WMI queries
#               - Consider both hardware and software firewalls
# Micorosft Office Excel must be installed on the host running this script if $OutputFormat = "Excel"
# Run from within the Active Directory Module for PowerShell (or import the Active Directory PS module into the PS session)
##########################################################
# Check if $SpecificDC value is not null
if ($SpecificDC)
{
$DomainControllers = Get-ADDomainController $SpecificDC
}
else
{
# Domain Controller discovery
clear
write-host "" `r
Write-host "Discovering all DCs in the domain $DomainDNS ....." `r
Write-host "`n"
$DomainControllers = Get-ADDomainController -Filter * -Server $DomainDNS
}
# Output list of Domain Controllers to screen
ForEach ($DC in $DomainControllers)
{
Write-Host $DC.Name
}
##########################################################
# Start of Excel section
##########################################################
if ($OutputFormat -eq "Excel")
{
# Set up Excel spreadsheet
$erroractionpreference = "SilentlyContinue"
$a = New-Object -comobject Excel.Application
$a.visible = $True
$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)
$c.Cells.Item(1,1) = "Name"
$c.Cells.Item(1,2) = "Description"
$c.Cells.Item(1,3) = "Ping Status"
$c.Cells.Item(1,4) = "FQDN"
$c.Cells.Item(1,5) = "IP Address"
$c.Cells.Item(1,6) = "Operating System"
$c.Cells.Item(1,7) = "Service Pack"
$c.Cells.Item(1,8) = "Domain"
$c.Cells.Item(1,9) = "GC"
$c.Cells.Item(1,10) = "FSMO Roles"
$c.Cells.Item(1,11) = "AD Site"
$c.Cells.Item(1,12) = "Read Only"
$c.Cells.Item(1,13) = "LDAP Port"
$c.Cells.Item(1,14) = "SSL Port"
$c.Cells.Item(1,15) = "Roles Installed"
$c.Cells.Item(1,16) = "Last Boot Time"
$c.Cells.Item(1,17) = "Virtual"
$c.Cells.Item(1,18) = "DNS Servers"
$c.Cells.Item(1,19) = "RAM (MB)"
$c.Cells.Item(1,20) = "CPU Speed (MHz)"
$c.Cells.Item(1,21) = "CPU Cores"
$c.Cells.Item(1,22) = "Logical CPUs"
$c.Cells.Item(1,23) = "Timezone"
$c.Cells.Item(1,24) = "Free space (C: GB)"
$c.Cells.Item(1,25) = "SCCM Client"
$c.Cells.Item(1,26) = "SCOM Client"
$c.Cells.Item(1,27) = "Query Time"
$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True
$d.EntireColumn.AutoFit($True)
$intRow = 2
# Start querying each Domain Controller
ForEach ($DC in $DomainControllers)
{
# Output progress bar to the screen
$i++
$numberofDCs = $DomainControllers.count
$ProgressName = $DC.Name
Write-Progress -Activity "Collecting Domain Controller information" -status "Contacting $ProgressName [$i out of $numberofDCs].  Overall percentage complete:" -percentComplete ($i / $DomainControllers.count*100)
# Write Domain Controller name in capitals
$c.Cells.Item($intRow, 1) = $DC.Name.ToUpper()
# Test connectivity from host to each Domain Controller
$ping = new-object System.Net.NetworkInformation.Ping
$Reply = $ping.send($DC.Name,$Timeout)
if ($Reply.status -eq "Success")
{
$DCPing = "Resolved & active"
$Online = $true
}
elseif ($Reply.status -eq "TimedOut")
{
$DCPing = "Resolved host but timed out"
$Online = $false
}
else
{
$DCPing = "Unable to resolve"
$Online = $false
}
$Reply = ""
# Write ping status
$c.Cells.Item($intRow, 3) = $DCPing
# Check WMI connectivity
$wmi = $null
$wmi = Get-WmiObject -class Win32_ComputerSystem -ComputerName $DC.Name -ErrorAction SilentlyContinue
if ($wmi)
{
# Able to connect to WMI
$ConnectViaWmi = $True
}
else
{
# Unable to connect to WMI
$ConnectViaWmi = $False
}
# Query and write computer description retrieved from AD
$DCDesc = (Get-ADComputer -Properties * -Filter {name -like $DC.Name}).Description
$c.Cells.Item($intRow, 2) = $DCDesc
# Query and write computer FQDN retrieved from AD
$DCFQDN = (Get-ADComputer -Properties * -Filter {name -like $DC.Name}).DNSHostName
$c.Cells.Item($intRow, 4) = $DCFQDN
# Query and write computer IPv4 address retrieved from AD
$DCIPv4 = (Get-ADComputer -Properties * -Filter {name -like $DC.Name}).IPv4Address
$c.Cells.Item($intRow, 5) = $DCIPv4
# Query and write computer Operating System retrieved from AD
$DCOS = (Get-ADComputer -Properties * -Filter {name -like $DC.Name}).OperatingSystem
$c.Cells.Item($intRow, 6) = $DCOS
# Query and write computer Operating System Service Pack retrieved from AD
$DCOSSP = (Get-ADComputer -Properties * -Filter {name -like $DC.Name}).OperatingSystemServicePack
$c.Cells.Item($intRow, 7) = $DCOSSP
# Query and write computer domain retrieved from AD
$DCDomain = (Get-ADDomainController -Filter {name -like $DC.Name}).Domain
$c.Cells.Item($intRow, 8) = $DCDomain
# Query and write computer global catalog boolean retrieved from AD
$DCGC = (Get-ADDomainController -Filter {name -like $DC.Name}).IsGlobalCatalog
$c.Cells.Item($intRow, 9) = $DCGC
# Query and write computer FSMO roles retrieved from AD
$DCFSMOOutput = ((Get-ADDomainController -Filter {name -like $DC.Name}).OperationMasterRoles | Out-String)
$DCFSMO = ($DCFSMOOutput).Replace("`n",'  ')
$c.Cells.Item($intRow, 10) = $DCFSMO
# Query and write computer AD site info retrieved from AD
$DCSite = (Get-ADDomainController -Filter {name -like $DC.Name}).Site
$c.Cells.Item($intRow, 11) = $DCSite
# Query and write computer read-only domain controller info retrieved from AD
$DCRO = (Get-ADDomainController -Filter {name -like $DC.Name}).IsReadOnly
$c.Cells.Item($intRow, 12) = $DCRO
# Query and write computer LDAP port retrieved from AD
$DCLDAP = (Get-ADDomainController -Filter {name -like $DC.Name}).LdapPort
$c.Cells.Item($intRow, 13) = $DCLDAP
# Query and write computer SSL port retrieved from AD
$DCSLDAP = (Get-ADDomainController -Filter {name -like $DC.Name}).SSLPort
$c.Cells.Item($intRow, 14) = $DCSLDAP
# Query the Server Roles that are installed on the Domain Controller (eg DNS, DHCP, ADDS)
# Assuming role ID is less than 30 - http://msdn.microsoft.com/en-gb/library/windows/desktop/cc280268(v=vs.85).aspx
if ($Online -eq $true -and $DCOS -notlike "*2003*" -and $ConnectViaWmi -eq $true)
{
$DCRole = (gwmi win32_ServerFeature -filter "ID<30" -computername $DC.Name | Select-Object "Name")
$k = @()
foreach ($j in $DCRole)
{$k += $j.Name
$c.Cells.Item($intRow, 15) = ($k -Join ', ')}
}
elseif ($Online -eq $false)
{
$c.Cells.Item($intRow, 15) = "Cannot query - Ping timeout"
}
elseif ($ConnectViaWmi -eq $false -and $DCOS -notlike "*2003*")
{
$c.Cells.Item($intRow, 15) = "Cannot connect to WMI"
}
elseif ($DCOS -like "*2003*")
{
$c.Cells.Item($intRow, 15) = "N/A - Windows 2003 Server OS"
}
# Query last boot time
if ($Online -eq $true -and $ConnectViaWmi -eq $true)
{
$date = new-object -com WbemScripting.SWbemDateTime
$z = get-wmiobject Win32_OperatingSystem -computername $DC.Name
foreach ($k in $z)
{$date.value = $k.lastBootupTime
If ($k.Version -eq "*" )
{$c.Cells.Item($intRow, 16) = $Date.GetVarDate($True)}
Else
{$c.Cells.Item($intRow, 16) = $Date.GetVarDate($False)}
}
}
elseif ($Online -eq $false)
{
$c.Cells.Item($intRow, 16) = "Cannot query - Ping timeout"
}
elseif ($ConnectViaWmi -eq $false)
{
$c.Cells.Item($intRow, 16) = "Cannot connect to WMI"
}
# Query if virtual machine / virtual hardware
$DCVM = $null
if ($Online -eq $true -and $ConnectViaWmi -eq $true)
{
$bios = gwmi Win32_BIOS -computername $DC.Name | Select-Object "version","serialnumber"
$compsys = gwmi Win32_ComputerSystem -computername $DC.Name | Select-Object "model","manufacturer"
if($bios.Version -match "VRTUAL") {$DCVM = "Virtual - Hyper-V"}
elseif($bios.Version -match "A M I") {$DCVM = "Virtual -  Virtual PC"}
elseif($bios.Version -like "*Xen*") {$DCVM = "Virtual - Xen"}
elseif($bios.SerialNumber -like "*VMware*") {$DCVM = "Virtual - VMWare"}
elseif($compsys.manufacturer -like "*Microsoft*") {$DCVM = "Virtual - Hyper-V"}
elseif($compsys.manufacturer -like "*VMWare*") {$DCVM = "Virtual - VMWare"}
elseif($compsys.model -like "*Virtual*") {$DCVM = "Virtual"}
else {$DCVM = "Physical"}
}
elseif ($Online -eq $false)
{
$DCVM = "Cannot query - Ping timeout"
}
elseif ($ConnectViaWmi -eq $false)
{
$DCVM = "Cannot connect to WMI"
}
# Query machine network cards to see which DNS Servers are used for queries
if ($Online -eq $true -and $ConnectViaWmi -eq $true)
{
$Adapters = Get-Wmiobject Win32_NetworkAdapterConfiguration -Computername $DC.Name | Where-Object{$_.IPEnabled -eq $True}
ForEach($Adapter In $Adapters)
{
[String]$DNSServers = ""
$Adapters2 = Get-Wmiobject Win32_NetworkAdapter -Computername $DC.Name | Where-Object{$_.Caption -eq $Adapter.Caption}
[String]$NetID = $Adapters2.NetConnectionID
If($Adapter.DNSServerSearchOrder -ne $Null)
{ForEach($Address In $Adapter.DNSServerSearchOrder)
{$DNSServers += $Address + "  "}
}
}
}
elseif ($Online -eq $false)
{
$DNSServers = "Cannot query - Ping timeout"
}
elseif ($ConnectViaWmi -eq $false)
{
$DNSServers = "Cannot connect to WMI"
}
# Output time of query
$datetime = get-date -uformat "%d/%m/%Y %H:%M:%S"
$UTC = get-date -uformat "%Z"
$querytime = $datetime + " UTC " + $UTC
# Check amount of installed RAM
if ($Online -eq $true -and $ConnectViaWmi -eq $true)
{
$colItems = get-wmiobject -class "Win32_ComputerSystem" -namespace "rootCIMV2" -computername $DC.Name
foreach ($objItem in $colItems)
{$DCRAM = [math]::round($objItem.TotalPhysicalMemory/1024/1024, 0)}
}
elseif ($Online -eq $false)
{
$DCRAM = "Cannot query - Ping timeout"
}
elseif ($ConnectViaWmi -eq $false)
{
$DCRAM = "Cannot connect to WMI"
}
# Query CPU information
if ($Online -eq $true -and $DCOS -notlike "*2003*" -and $ConnectViaWmi -eq $true)
{
$CPUproperty = "maxclockspeed", "numberOfCores", "NumberOfLogicalProcessors"
$DCCPUSpeed = Get-WmiObject -class "win32_processor" -Property $CPUproperty -computername $DC.Name -filter "deviceid='CPU0'" | Select-Object -expand "maxclockspeed"
$Win32_cpu = Get-WmiObject -class win32_processor -computername $DC.Name
$DCCPULogical = ($Win32_cpu | measure-object).count
$DCCPUCores = ($Win32_cpu | measure-object NumberOfCores -sum).sum
}
elseif ($Online -eq $true -and $DCOS -like "*2003*" -and $ConnectViaWmi -eq $true)
{
$CPUproperty = "maxclockspeed"
$DCCPUSpeed = Get-WmiObject -class "win32_processor" -Property $CPUproperty -computername $DC.Name -filter "deviceid='CPU0'" | Select-Object -expand "maxclockspeed"
$physCount = new-object hashtable
$Win32_cpu = Get-WmiObject -class win32_processor -computername $DC.Name
$Win32_cpu |%{$physCount[$_.SocketDesignation] = 1}
$DCCPULogical = $physCount.count
$DCCPUCores = ($Win32_cpu | measure-object).count
}
elseif ($Online -eq $false)
{
$DCCPUSpeed = "Cannot query - Ping timeout"
$DCCPUCores = "Cannot query - Ping timeout"
$DCCPULogical = "Cannot query - Ping timeout"
}
elseif ($ConnectViaWmi -eq $false)
{
$DCCPUSpeed = "Cannot connect to WMI"
$DCCPUCores = "Cannot connect to WMI"
$DCCPULogical = "Cannot connect to WMI"
}
# Query timezone information
if ($Online -eq $true -and $ConnectViaWmi -eq $true)
{
$DCTZ = Get-WmiObject -class "Win32_TimeZone" -computername $DC.Name | Select-Object -expand "Caption"
}
elseif ($Online -eq $false)
{
$DCTZ = "Cannot query - Ping timeout"
}
elseif ($ConnectViaWmi -eq $false)
{
$DCTZ = "Cannot connect to WMI"
}
# Query free disk space on C drive
if ($Online -eq $true -and $ConnectViaWmi -eq $true)
{
$Cdisk = Get-WmiObject Win32_LogicalDisk -ComputerName $DC.Name -Filter "DeviceID='C:'" | Select-Object FreeSpace
$Cdisk.FreeSpace = $([Math]::Round($Cdisk.FreeSpace/1073741824,1))
$DCDiskFree = $Cdisk.FreeSpace
}
elseif ($Online -eq $false)
{
$DCDiskFree = "Cannot query - Ping timeout"
}
elseif ($ConnectViaWmi -eq $false)
{
$DCDiskFree = "Cannot connect to WMI"
}
# Check if SCCM client is installed
if ($Online -eq $true -and $ConnectViaWmi -eq $true)
{
$SCCMCheck = get-wmiobject -namespace rootcimv2 -computername $DC.Name -class win32_process -filter 'Name="ccmexec.exe"'
if (-not $SCCMCheck)
{
$DCSCCMClient = "No"
}
else
{
$siteCode = (Get-WmiObject -computername $DC.Name -namespace rootccmpolicymachine -Class CCM_SystemHealthClientConfig).SiteCode
$DCSCCMClient = "Yes - " +$siteCode
}
}
elseif ($Online -eq $false)
{
$DCSCCMClient = "Cannot query - Ping timeout"
}
elseif ($ConnectViaWmi -eq $false)
{
$DCSCCMClient = "Cannot connect to WMI"
}
# Check if SCOM client is installed
if ($Online -eq $true -and $ConnectViaWmi -eq $true)
{
$SCOMCheck = get-wmiobject -namespace rootcimv2 -computername $DC.Name -class win32_process -filter 'Name="HealthService.exe"'
if (-not $SCOMCheck)
{
$DCSCOMClient = "No"
}
else
{
$Path = "hklm:SOFTWAREMICROSOFTMICROSOFT OPERATIONS MANAGER3.0AGENT MANAGEMENT GROUPS"
$baseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine", $DC.Name)
$s = ""
## Open the key
$key = $baseKey.OpenSubKey("SOFTWAREMICROSOFTMICROSOFT OPERATIONS MANAGER3.0AGENT MANAGEMENT GROUPS")
## Retrieve all of its children
foreach($subkeyName in $key.GetSubKeyNames())
{
## Open the subkey
$subkey = $key.OpenSubKey($subkeyName)
$returnObject = [PsObject] $subKey
$returnObject | Add-Member NoteProperty PsChildName $subkeyName | Select PSChildName
## Output the key
$s += $returnObject.PsChildName + " "
## Close the child key
$subkey.Close()
}
## Close the key and base keys
$key.Close()
$baseKey.Close()
$DCSCOMClient = "Yes - " +$s
}
}
elseif ($Online -eq $false)
{
$DCSCOMClient = "Cannot query - Ping timeout"
}
elseif ($ConnectViaWmi -eq $false)
{
$DCSCOMClient = "Cannot connect to WMI"
}
# Write out Excel rows for direct machine queries
$c.Cells.Item($intRow, 17) = $DCVM
$c.Cells.Item($intRow, 18) = $DNSServers
$c.Cells.Item($intRow, 19) = $DCRAM
$c.Cells.Item($intRow, 20) = $DCCPUSpeed
$c.Cells.Item($intRow, 21) = $DCCPUCores
$c.Cells.Item($intRow, 22) = $DCCPULogical
$c.Cells.Item($intRow, 23) = $DCTZ
$c.Cells.Item($intRow, 24) = $DCDiskFree
$c.Cells.Item($intRow, 25) = $DCSCCMClient
$c.Cells.Item($intRow, 26) = $DCSCOMClient
$c.Cells.Item($intRow, 27) = $querytime
# Next Excel row
$intRow = $intRow + 1
[array] $DomainDCs += $DC.HostName
$Online = $false
# End of foreach DC
}
# Configure Excel Autofit for rows and columns
$d.EntireColumn.AutoFit()|out-null
$d.EntireRow.AutoFit()|out-null
# Configure Excel Filters - uncomment if required
#$d.EntireColumn.AutoFilter()
Write-host "`n"
Write-Host "Excel file creation complete...."
}
##########################################################
# End of Excel section
##########################################################
##########################################################
# Start of CSV section
##########################################################
if ($OutputFormat -eq "CSV")
{
# Create empty array
$report = @()
$erroractionpreference = "SilentlyContinue"
# Start querying each Domain Controller
ForEach ($DC in $DomainControllers)
{
# Output progress bar to the screen
$i++
$numberofDCs = $DomainControllers.count
$ProgressName = $DC.Name
Write-Progress -Activity "Collecting Domain Controller information" -status "Contacting $ProgressName [$i out of $numberofDCs].  Overall percentage complete:" -percentComplete ($i / $DomainControllers.count*100)
# Test connectivity from host to each Domain Controller
$ping = new-object System.Net.NetworkInformation.Ping
$Reply = $ping.send($DC.Name,$Timeout)
if ($Reply.status -eq "Success")
{
$DCPing = "Resolved & active"
$Online = $true
}
elseif ($Reply.status -eq "TimedOut")
{
$DCPing = "Resolved host but timed out"
$Online = $false
}
else
{
$DCPing = "Unable to resolve"
$Online = $false
}
$Reply = ""
# Check WMI connectivity
$wmi = $null
$wmi = Get-WmiObject -class Win32_ComputerSystem -ComputerName $DC.Name -ErrorAction SilentlyContinue
if ($wmi)
{
# Able to connect to WMI
$ConnectViaWmi = $True
}
else
{
# Unable to connect to WMI
$ConnectViaWmi = $False
}
# Query computer description retrieved from AD
$DCDesc = (Get-ADComputer -Properties * -Filter {name -like $DC.Name}).Description
# Query computer FQDN retrieved from AD
$DCFQDN = (Get-ADComputer -Properties * -Filter {name -like $DC.Name}).DNSHostName
# Query IPv4 address retrieved from AD
$DCIPv4 = (Get-ADComputer -Properties * -Filter {name -like $DC.Name}).IPv4Address
# Query Operating System retrieved from AD
$DCOS = (Get-ADComputer -Properties * -Filter {name -like $DC.Name}).OperatingSystem
# Query Operating System Service Pack retrieved from AD
$DCOSSP = (Get-ADComputer -Properties * -Filter {name -like $DC.Name}).OperatingSystemServicePack
# Query computer domain retrieved from AD
$DCDomain = (Get-ADDomainController -Filter {name -like $DC.Name}).Domain
# Query computer global catalog boolean retrieved from AD
$DCGC = (Get-ADDomainController -Filter {name -like $DC.Name}).IsGlobalCatalog
# Query computer FSMO roles retrieved from AD
$DCFSMOOutput = ((Get-ADDomainController -Filter {name -like $DC.Name}).OperationMasterRoles | Out-String)
$DCFSMO = ($DCFSMOOutput).Replace("`n",'  ')
# Query computer AD site info retrieved from AD
$DCSite = (Get-ADDomainController -Filter {name -like $DC.Name}).Site
# Query read-only domain controller info retrieved from AD
$DCRO = (Get-ADDomainController -Filter {name -like $DC.Name}).IsReadOnly
# Query computer LDAP port retrieved from AD
$DCLDAP = (Get-ADDomainController -Filter {name -like $DC.Name}).LdapPort
# Query computer SSL port retrieved from AD
$DCSLDAP = (Get-ADDomainController -Filter {name -like $DC.Name}).SSLPort
# Query the Server Roles that are installed on the Domain Controller (eg DNS, DHCP, ADDS)
# Assuming role ID is less than 30 - http://msdn.microsoft.com/en-gb/library/windows/desktop/cc280268(v=vs.85).aspx
if ($Online -eq $true -and $DCOS -notlike "*2003*" -and $ConnectViaWmi -eq $true)
{
$DCRole = (gwmi win32_ServerFeature -filter "ID<30" -computername $DC.Name | Select-Object "Name")
$k = @()
foreach ($j in $DCRole)
{$k += $j.Name
$DCRoleOutput = ($k -Join ', ')
}
}
elseif ($Online -eq $false)
{
$DCRoleOutput = "Cannot query - Ping timeout"
}
elseif ($DCOS -like "*2003*")
{
$DCRoleOutput = "N/A - Windows 2003 Server OS"
}
elseif ($ConnectViaWmi -eq $false)
{
$DCRoleOutput = "Cannot connect to WMI"
}
# Query last boot time
if ($Online -eq $true -and $ConnectViaWmi -eq $true)
{
$date = new-object -com WbemScripting.SWbemDateTime
$z = get-wmiobject Win32_OperatingSystem -computername $DC.Name
foreach ($k in $z)
{
$date.value = $k.lastBootupTime
If ($k.Version -eq "*" )
{
$LastBoot = $Date.GetVarDate($True)
}
Else
{
$LastBoot = $Date.GetVarDate($False)
}
}
}
elseif ($Online -eq $false)
{
$LastBoot = "Cannot query - Ping timeout"
}
elseif ($ConnectViaWmi -eq $false)
{
$LastBoot = "Cannot connect to WMI"
}
# Query if virtual machine / virtual hardware
$DCVM = $null
if ($Online -eq $true -and $ConnectViaWmi -eq $true)
{
$bios = gwmi Win32_BIOS -computername $DC.Name | Select-Object "version","serialnumber"
$compsys = gwmi Win32_ComputerSystem -computername $DC.Name | Select-Object "model","manufacturer"
if($bios.Version -match "VRTUAL") {$DCVM = "Virtual - Hyper-V"}
elseif($bios.Version -match "A M I") {$DCVM = "Virtual -  Virtual PC"}
elseif($bios.Version -like "*Xen*") {$DCVM = "Virtual - Xen"}
elseif($bios.SerialNumber -like "*VMware*") {$DCVM = "Virtual - VMWare"}
elseif($compsys.manufacturer -like "*Microsoft*") {$DCVM = "Virtual - Hyper-V"}
elseif($compsys.manufacturer -like "*VMWare*") {$DCVM = "Virtual - VMWare"}
elseif($compsys.model -like "*Virtual*") {$DCVM = "Virtual"}
else {$DCVM = "Physical"}
}
elseif ($Online -eq $false)
{
$DCVM = "Cannot query - Ping timeout"
}
elseif ($ConnectViaWmi -eq $false)
{
$DCVM = "Cannot connect to WMI"
}
# Query machine network cards to see which DNS Servers are used for queries
if ($Online -eq $true -and $ConnectViaWmi -eq $true)
{
$Adapters = Get-Wmiobject Win32_NetworkAdapterConfiguration -Computername $DC.Name | Where-Object{$_.IPEnabled -eq $True}
ForEach($Adapter In $Adapters)
{
[String]$DNSServers = ""
$Adapters2 = Get-Wmiobject Win32_NetworkAdapter -Computername $DC.Name | Where-Object{$_.Caption -eq $Adapter.Caption}
[String]$NetID = $Adapters2.NetConnectionID
If($Adapter.DNSServerSearchOrder -ne $Null)
{ForEach($Address In $Adapter.DNSServerSearchOrder)
{$DNSServers += $Address + "  "}
}
}
}
elseif ($Online -eq $false)
{
$DNSServers = "Cannot query - Ping timeout"
}
elseif ($ConnectViaWmi -eq $false)
{
$DNSServers = "Cannot connect to WMI"
}
# Output time of query
$datetime = get-date -uformat "%d/%m/%Y %H:%M:%S"
$UTC = get-date -uformat "%Z"
$querytime = $datetime + " UTC " + $UTC
# Check amount of installed RAM
if ($Online -eq $true -and $ConnectViaWmi -eq $true)
{
$colItems = get-wmiobject -class "Win32_ComputerSystem" -namespace "rootCIMV2" -computername $DC.Name
foreach ($objItem in $colItems)
{
$DCRAM = [math]::round($objItem.TotalPhysicalMemory/1024/1024, 0)
}
}
elseif ($Online -eq $false)
{
$DCRAM = "Cannot query - Ping timeout"
}
elseif ($ConnectViaWmi -eq $false)
{
$DCRAM = "Cannot connect to WMI"
}
# Query CPU information
if ($Online -eq $true -and $DCOS -notlike "*2003*" -and $ConnectViaWmi -eq $true)
{
$CPUproperty = "maxclockspeed", "numberOfCores", "NumberOfLogicalProcessors"
$DCCPUSpeed = Get-WmiObject -class "win32_processor" -Property $CPUproperty -computername $DC.Name -filter "deviceid='CPU0'" | Select-Object -expand "maxclockspeed"
$Win32_cpu = Get-WmiObject -class win32_processor -computername $DC.Name
$DCCPULogical = ($Win32_cpu | measure-object).count
$DCCPUCores = ($Win32_cpu | measure-object NumberOfCores -sum).sum
}
elseif ($Online -eq $true -and $DCOS -like "*2003*" -and $ConnectViaWmi -eq $true)
{
$CPUproperty = "maxclockspeed"
$DCCPUSpeed = Get-WmiObject -class "win32_processor" -Property $CPUproperty -computername $DC.Name -filter "deviceid='CPU0'" | Select-Object -expand "maxclockspeed"
$physCount = new-object hashtable
$Win32_cpu = Get-WmiObject -class win32_processor -computername $DC.Name
$Win32_cpu |%{$physCount[$_.SocketDesignation] = 1}
$DCCPULogical = $physCount.count
$DCCPUCores = ($Win32_cpu | measure-object).count
}
elseif ($Online -eq $false)
{
$DCCPUSpeed = "Cannot query - Ping timeout"
$DCCPUCores = "Cannot query - Ping timeout"
$DCCPULogical = "Cannot query - Ping timeout"
}
elseif ($ConnectViaWmi -eq $false)
{
$DCCPUSpeed = "Cannot connect to WMI"
$DCCPUCores = "Cannot connect to WMI"
$DCCPULogical = "Cannot connect to WMI"
}
# Query timezone information
if ($Online -eq $true -and $ConnectViaWmi -eq $true)
{
$DCTZ = Get-WmiObject -class "Win32_TimeZone" -computername $DC.Name | Select-Object -expand "Caption"
}
elseif ($Online -eq $false)
{
$DCTZ = "Cannot query - Ping timeout"
}
elseif ($ConnectViaWmi -eq $false)
{
$DCTZ = "Cannot connect to WMI"
}
# Query free disk space on C drive
if ($Online -eq $true -and $ConnectViaWmi -eq $true)
{
$Cdisk = Get-WmiObject Win32_LogicalDisk -ComputerName $DC.Name -Filter "DeviceID='C:'" | Select-Object FreeSpace
$Cdisk.FreeSpace = $([Math]::Round($Cdisk.FreeSpace/1073741824,1))
$DCDiskFree = $Cdisk.FreeSpace
}
elseif ($Online -eq $false)
{
$DCDiskFree = "Cannot query - Ping timeout"
}
elseif ($ConnectViaWmi -eq $false)
{
$DCDiskFree = "Cannot connect to WMI"
}
# Check if SCCM client is installed
if ($Online -eq $true -and $ConnectViaWmi -eq $true)
{
$SCCMCheck = get-wmiobject -namespace rootcimv2 -computername $DC.Name -class win32_process -filter 'Name="ccmexec.exe"'
if (-not $SCCMCheck)
{
$DCSCCMClient = "No"
}
else
{
$siteCode = (Get-WmiObject -computername $DC.Name -namespace rootccmpolicymachine -Class CCM_SystemHealthClientConfig).SiteCode
$DCSCCMClient = "Yes - " +$siteCode
}
}
elseif ($Online -eq $false)
{
$DCSCCMClient = "Cannot query - Ping timeout"
}
elseif ($ConnectViaWmi -eq $false)
{
$DCSCCMClient = "Cannot connect to WMI"
}
# Check if SCOM client is installed
if ($Online -eq $true -and $ConnectViaWmi -eq $true)
{
$SCOMCheck = get-wmiobject -namespace rootcimv2 -computername $DC.Name -class win32_process -filter 'Name="HealthService.exe"'
if (-not $SCOMCheck)
{
$DCSCOMClient = "No"
}
else
{
$Path = "hklm:SOFTWAREMICROSOFTMICROSOFT OPERATIONS MANAGER3.0AGENT MANAGEMENT GROUPS"
$baseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine", $DC.Name)
$s = ""
## Open the key
$key = $baseKey.OpenSubKey("SOFTWAREMICROSOFTMICROSOFT OPERATIONS MANAGER3.0AGENT MANAGEMENT GROUPS")
## Retrieve all of its children
foreach($subkeyName in $key.GetSubKeyNames())
{
## Open the subkey
$subkey = $key.OpenSubKey($subkeyName)
$returnObject = [PsObject] $subKey
$returnObject | Add-Member NoteProperty PsChildName $subkeyName | Select PSChildName
## Output the key
$s += $returnObject.PsChildName + " "
## Close the child key
$subkey.Close()
}
## Close the key and base keys
$key.Close()
$baseKey.Close()
$DCSCOMClient = "Yes - " +$s
}
}
elseif ($Online -eq $false)
{
$DCSCOMClient = "Cannot query - Ping timeout"
}
elseif ($ConnectViaWmi -eq $false)
{
$DCSCOMClient = "Cannot connect to WMI"
}
[array] $DomainDCs += $DC.HostName
$Online = $false
# Create object to collect elements
$OutputObj = New-Object -TypeName PSObject -Property @{
"Name" = $DC.Name.ToUpper()
"Description" = $DCDesc
"Ping Status" = $DCPing
"FQDN" = $DCFQDN
"IP Address" = $DCIPv4
"Operating System" = $DCOS
"Service Pack" = $DCOSSP
"Domain" = $DCDomain
"GC" = $DCGC
"FSMO Roles" = $DCFSMO
"AD Site" = $DCSite
"Read Only" = $DCRO
"LDAP Port" = $DCLDAP
"SSL Port" = $DCSLDAP
"Roles Installed" = $DCRoleOutput
"Last Boot Time" = $LastBoot
"Virtual" = $DCVM
"DNS Servers" = $DNSServers
"RAM (MB)" = $DCRAM
"CPU Speed (MHz)" = $DCCPUSpeed
"CPU Cores" = $DCCPUCores
"Logical CPUs" = $DCCPULogical
"Timezone" = $DCTZ
"Free space (C: GB)" = $DCDiskFree
"SCCM Client" = $DCSCCMClient
"SCOM Client" = $DCSCOMClient
"Query Time" = $querytime
} | Select-Object "Name","Description","Ping Status","FQDN","IP Address","Operating System","Service Pack","Domain","GC","FSMO Roles","AD Site","Read Only","LDAP Port","SSL Port","Roles Installed","Last Boot Time","Virtual","DNS Servers","RAM (MB)","CPU Speed (MHz)","CPU Cores","Logical CPUs","Timezone","Free space (C: GB)","SCCM Client","SCOM Client","Query Time"
# Add item to report
$report += $OutputObj
$Progress
# End of for each Domain Controller
}
# Export information to specified CSV file
$report | Export-Csv $FileLocation -Force -NoTypeInformation
Write-host "`n"
Write-Host "CSV file creation complete...."
# End of if CSV section
}
##########################################################
# End of CSV section
##########################################################
# Output domain controller count to screen
if ($SpecificDC -ne $null)
{
Write-host "`n"
$DCCount = $DomainDCs.Count
$DomainDCs = $DomainDCs | sort
Write-Host "Found $DCCount DCs in $DomainDNS. Displaying all DCs in the domain :" `r
Write-host "`n"
$DomainDCs
Write-host "`n"
}
