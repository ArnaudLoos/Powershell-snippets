<#
.SYNOPSIS
Script that will list the logon information of AD users.

.DESCRIPTION
This script will list the AD users logon information with their logged on computers by inspecting the Kerberos TGT Request
Events(EventID 4768) from domain controllers. Not Only User account Name is fetched, but also users OU path and Computer 
Accounts are retrieved. You can also list the history of last logged on users. In Environment where Exchange Servers are 
used, the exchange servers authentication request for users will also be logged since it also uses EventID (4768) to for 
TGT Request. You can also export the result to CSV file format. Powershell version 3.0 is needed to use the script.
You can Define the following parameters to suite your need:																					
-MaxEvent		Specify the number of all (4768) events to search for TGT Requests. Default is 1000.											
-LastLogonOnly	Display only the history of last logon users.																				
-OuOnly			Do not display the full path of users/computers. Only OU is displayed.														
Author: phyoepaing3.142@gmail.com																											
Country: Myanmar(Burma)																														
Released Date: 08/29/2016
https://gallery.technet.microsoft.com/Get-All-AD-Users-Logon-9e721a89

Example usage:																					  
.\Get_AD_Users_Logon_History.ps1 -MaxEvent 800 -LastLogonOnly -OuOnly

.EXAMPLE
.\Get_AD_Users_Logon_History.ps1 | Format-Table * -Auto
This command will retrieve AD users logon within default 1000 EventID-4768 events and display the results as table.

.EXAMPLE
.\Get_AD_Users_Logon_History.ps1 -MaxEvent 500 -LastLogonOnly -OuOnly
This command will retrieve AD users logon within 500 EventID-4768 events and show only the last logged users with their
related logged on computers. Only OU name is displayed in results.

.EXAMPLE
.\Get_AD_Users_Logon_History.ps1 | Export-Csv Users_Loggedon_History.csv
This command will retrieve AD users logon within default 1000 EventID-4768 events and export the result to CSV file.

.PARAMETER MaxEvent
This paraemeter will specify the number of EventID-4768 events to look for.

.PARAMETER LastLogonOnly
This paraemeter will display the history of last logged on users in descending order.

.PARAMETER OuOnly
This paraemeter will show only the OU names for users/computers but not the full path.

.LINK
You can find this script and more at: https://www.sysadminplus.blogspot.com/
#>

param( [switch]$LastLogonOnly,[switch]$OuOnly,[int]$MaxEvent=1000)
############## Find EventID 4768 with user's requesting Kerberos TGT, skipping Exchange Health Mailbox request and extracting  Users/Client names,IP Addresses ####
$Domain = (Get-WmiObject Win32_Computersystem).domain
$read_log={
Param ($MaxEvent,$OuOnly,$Domain)																					## Define parameter to pass maxevent to scripblock
$EventInfo=Get-WinEvent -FilterHashTable @{LogName="Security";ID=4768} -MaxEvents 200000  | select -first $MaxEvent | where {$_.Message -notmatch "SM_" } | where { $_.Message -notmatch "\$" } | 
 select @{N="Authenticated DC";Exp={$_.MachineName}}, 
 @{N="LoggedOn Time";Exp={$_.TimeCreated}}, 
 @{N="User"; Exp={ $SplitAccountName=(($_.Message -Split "\n") -match "Account Name") -split ':';$SplitAccountName[$SplitAccountName.Length-1].Trim() }}, 
 @{N="User Location";Exp={}}, @{N="Workstation";Exp={}}, 
 @{N="IP Address"; Exp={if((($_.Message -split "\n") -match "Client Address:").Trim() -match "::1" ) {"localhost"} 
 else { $SplitClientAddress=(($_.Message -Split "\n") -match "Client Address") -split ':'; $splitClientAddress[$splitClientAddress.Length-1].Trim() }}}, 
 @{N="Computer Location";Exp={}}

$EventInfo | foreach {
$IPAddress=$_."Client Address"

if ($_."IP Address" -eq "localhost")
{ $Client_Name=[system.net.dns]::GetHostbyName($env:computername).hostname }
else
{ 
###### Resolve the PTR record to find AD computer information ################
		#$Client_Name=(Resolve-DnsName $_."IP Address").NameHost
		if ((Resolve-dnsname $_."IP Address" -Type PTR -TcpOnly -DnsOnly -ErrorAction "SilentlyContinue").Type -eq "PTR")
		{ 
		$Client_Name = (Resolve-dnsname $_."IP Address" -Type PTR -TcpOnly -DnsOnly).NameHost  
		
		}
		else
		{ $Client_Name = "NOT FOUND" }

}
## $_."Authenticated DC"=($_."Authenticated DC" -split "."+$Domain)[0]  ##Uncomment this line if you want to strip off domain name in "Authenticated DC" list
$user=$_.user
############# Find the User account in AD and if not found, throw and exception ###########
$Full_User_Property=0
	Try																									## Need Try statement to test and surpress error
	{
	$Full_User_Property = (Get-AdUser $_.user -Properties *)
	$_."User Location" = $Full_User_Property.CanonicalName.TrimStart($Domain).SubString(1)
	}
	catch
		{ } 																							## The $_."User Location" is not passed to catch statement thus needing another below statement to set value"
If (!$Full_User_Property)
	{ $_."User Location"="NOT FOUND"  }
$Full_User_Property=0

If($OuOnly -AND ($_."User Location" -ne "NOT FOUND"))
{
$_."User Location"= $_."User Location".Remove($_."User Location".LastIndexOf("/"))						##Trim the last user name part if -OuOnly flag is set
}

$_."Workstation"=($Client_Name -split "."+$Domain)[0]													## remove the domain suffix
########## Find the Computer account in AD and if not found, throw an exception ###########
$Full_Workstation_Property = 0
	Try
		{
		$Full_Workstation_Property = Get-AdComputer $_."Workstation" -Properties *
		$_."Computer Location"= $Full_Workstation_Property.CanonicalName.TrimStart($Domain).SubString(1)
		}
	catch
		{ }
########## Here the catch exception does not work in Invoke session so we need to manually set the "NOT FOUND" value ######
If (!$Full_Workstation_Property)
	{ $_."Computer Location" = "NOT FOUND"}
$Full_Workstation_Property=0

If ($OuOnly -AND ($_."Computer Location" -ne "NOT FOUND"))																							##Trim the last computer part if -OuOnly flag is set
	{
		$_."Computer Location"=$_."Computer Location".Remove($_."Computer Location".LastIndexOf("/"))
	}
Return $_
}
}

########### Job starts to query replica domain controllers #############
$result=@()
$RemoteJob=@()                                                           						       ## Make array of remote jobs

$DomainControllers = (Get-ADDomainController -Filter  { isGlobalCatalog -eq $true -or isGlobalCatalog -eq $false}).Name
############### Start the Local Job and Remote Job to find the event id ################
$LocalJobExists=0
If ($DomainControllers -contains $(hostname))															## Check if the computer running the script is Domain Controller itself
{
$LocalJob = Start-Job -scriptblock $read_log -ArgumentList $MaxEvent,$OuOnly,$Domain;$LocalJobExists=1	## If so, start job to query local domain controller
}

$DomainControllers | where {$_ -ne $(hostname)} | foreach {												## Start remote jobs on each other domain controllers
	$RemoteJob+= Invoke-Command -ComputerName $_ -ScriptBlock $read_log -ArgumentList $MaxEvent,$OuOnly,$Domain -AsJob  
	}

If ($LocalJobExists)
	{
	$result = $LocalJob | Wait-Job | Receive-Job; Remove-Job $LocalJob									## If the computer running the script is not a domain controller(may be RSAT installed), then all jobs will be remote jobs
	}

$RemoteJob | foreach { $result+=$_ | Wait-Job | Receive-Job ; Remove-Job $_}							## Wait and Receive remote jobs on each remote DCs and add to Local job result

If ($LastLogonOnly)
	{
		$result | Sort-Object "LoggedOn Time" -Descending | Group-Object User | foreach { $_ | Select -ExpandProperty Group | select * -First 1 -ExcludeProperty PsComputerName,RunSpaceID,PsShowComputerName }   ## the Last LoggedOn time of Each User
	}
else
	{
		$result | Sort-Object "LoggedOn Time" -Descending | Select * -ExcludeProperty PsComputerName,RunSpaceID,PsShowComputerName  ## Normal Results
	}


