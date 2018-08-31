<#
 
.DESCRIPTION
 *** THIS SCRIPT IS PROVIDED WITHOUT WARRANTY, USE AT YOUR OWN RISK ***
    
    Forest Information
        - Forest Root Domain
        - Forest Functional Level
        - Domains in the forest
        - AD Recycle BIN status
    Domain Information
        - Domain Functional Level
        - NETBIOS name
    FSMO Roles
        - Domain Naming Master
        - Schema Master
        - PDC Emulator
        - RID Master
        - Infrastructure Master
    Domain Controller Information
        - Domain
        - Forest
        - Computer Name
        - IP Address
        - Global Catalog
        - Read Only
        - Operating System
        - Operating System Version
        - Site
    DNS Information
        - Primary Zones
        - NS Records
        - MX Records
        - Forwards
        - Scavenging Enabled
        - Aging Enabled
    DHCP Information
        - Computer Name
        - IP Address
    Site Information
        - Site Names
        - Intersite Links
            - Name
            - Site Included
            - Site Cost
            - Site Replication Frequency
    GPO Information
        - Domain Name
        - Display Name
        - Creation Time
        - Modification Time
    Privileged Account Information
        - Enterprise Admin Group Members
        - Domain Admin Group Members
        - Schema Admin Group Members
        - Accounts that Passwords Never Expire
    Exchange Information
        - Organization Management Group Members
        - Exchange Server 
 
.NOTES
 File Name: Get-ForestInfo.ps1
 Author: David Hall
 Contact Info: 
 Website: www.signalwarrant.com
 Twitter: @signalwarrant
 Facebook: facebook.com/signalwarrant/
 Google +: plus.google.com/113307879414407675617
 YouTube Subscribe link: https://www.youtube.com/c/SignalWarrant1?sub_confirmation=1
 Requires: 
        Proper permissions to execute the script in the forest
        Execute the script from a Domain Controller, or the preferred
        method, on a client with the RSAT installed
 Tested: Windows Server 2012 R2, Windows Server 2016, Windows 10, PowerShell v3-5 
 
.PARAMETER(s)
    None
 
.EXAMPLE
    Get-ForestInfo.ps1
 
Inspired by: 
    Zachary Loeber's script 
    https://gallery.technet.microsoft.com/office/Active-Directory-Audit-7754a877
 
#>
 
#region Variables
 
######################################
# Variables
######################################
# Get the date for the filename 
$date = (Get-Date -Format d_MMMM_yyyy).toString()
# Where to ouput the html file
$filePATH = "$env:userprofile\Desktop\"
# Define the filename
$fileNAME = 'AD_Info_' + $date + '.html'
$File = $filePATH + $fileNAME
 
$forestInfo = Get-ADForest
$AllDomains = (Get-ADForest).Domains
$domainInfo = Get-ADDomain
$PDCEmulator = (Get-ADDomain).PDCEmulator
$DNSRoot = $domainInfo.dnsroot
$ADsiteLinks = Get-ADReplicationSiteLink -Filter *
#endregion
 
#region Forest Info
 
######################################
# Forest Information
######################################
# Forest Root Domain  
  $RootDomain = $forestInfo.RootDomain
# Forest Functional Level
  $ForestMode = $forestInfo.ForestMode
# Forest Domains
  $Domains = ($forestInfo | 
             Select-Object -ExpandProperty Domains) -join ' | '
# AD Recycle BIN Status
$ADRecycleBIN = Get-ADOptionalFeature -filter {Name -eq 'Recycle Bin Feature'} | 
              Select-Object -ExpandProperty EnabledScopes
 
  If (!$ADRecycleBIN){
      $ADRecycleBIN = 'Disabled'
  } else {
      $ADRecycleBIN = 'Enabled'
  }   
 
# Forest Information Output Object 
$ForestOutputObj  = New-Object -TypeName PSObject
  $ForestOutputObj | Add-Member -MemberType NoteProperty -Name ForestRootDomain -Value $RootDomain
  $ForestOutputObj | Add-Member -MemberType NoteProperty -Name ForestFunctionalLevel -Value $ForestMode
  $ForestOutputObj | Add-Member -MemberType NoteProperty -Name ForestDomains -Value $Domains
  $ForestOutputObj | Add-Member -MemberType NoteProperty -Name ADRecycleBIN -Value $ADRecycleBIN
  $ForestOutputObjHTML = $ForestOutputObj | ConvertTo-Html
#endregion
 
#region Domain Info
 
######################################
# Domain Information
######################################
# Get the Domain Functional Level 
$DomainMode = ($DNSRoot | foreach { Get-ADDomain -Identity $_ }  | 
              Select-Object -ExpandProperty DomainMode) -join ' | '
# Get the Domain NetBIOS Name
$NetBIOSName = $domainInfo.netBIOSName
 
# Domain Information Output Object 
$DomainOutputObj  = New-Object -TypeName PSObject
  $DomainOutputObj | Add-Member -MemberType NoteProperty -Name ForestFunctionalLevel -Value $DomainMode
  $DomainOutputObj | Add-Member -MemberType NoteProperty -Name NetBIOS_Name -Value $NetBIOSName
  $DomainOutputObjHTML = $DomainOutputObj | ConvertTo-Html
#endregion
 
#region FSMO Info
 
######################################
# FSMO Role Information
######################################
# Forest FSMO Roles
$ForestFSMO = $forestInfo | Select-Object -Property DomainNamingMaster,SchemaMaster | 
              ConvertTo-Html -Fragment
# Domain FSMO Roles
$DomainFSMO = $DNSRoot | foreach { Get-ADDomain -Identity $_ }  | 
              Select-Object -Property PDCEmulator,RIDMaster,InfrastructureMaster | 
              ConvertTo-Html -Fragment
#endregion
 
#region DC Info
 
######################################
# Domain Controllers Information
######################################
# Domain Controller Information
$DCs = Get-ADDomainController -Filter * | 
        Select-Object -Property Domain,Forest,Name,IPv4Address,IsGlobalCatalog,IsReadOnly,OperatingSystem,OperatingSystemVersion,Site
$DCOutputObjHTML = $DCs | ConvertTo-Html 
#endregion
 
#region DNS Info
 
######################################
# DNS Information
######################################
# Primary Zone Information
$PrimaryZones =  (Get-DnsServerZone -ComputerName $PDCEmulator | 
              Where-Object {$_.IsReverseLookupZone -eq $False} | 
              Select-Object -ExpandProperty ZoneName) -join '<br/>'
# NS records
$NSRecords =  (Resolve-DnsName -Name $DNSRoot -type ns | 
              Where-Object {$_.QueryType -eq 'NS'} | 
              Select-Object -ExpandProperty Server) -join '<br/>'
# MX Records
$MXRecords =  (Resolve-DnsName -Name $DNSRoot -type MX | 
              Where-Object {$_.QueryType -eq 'MX'} | 
              Select-Object -ExpandProperty Exchange) -join '<br/>'
# Forwarders
$DNSForwarders = (Get-DnsServerForwarder -ComputerName $PDCEmulator | 
              Select-Object -ExpandProperty IPAddress) -join '<br/>'
# Scavenging (Returns True or False)
$DNSScavenging = (Get-DnsServerScavenging -ComputerName $PDCEmulator).scavengingState
# Aging (Returns True or False)
$DNSAging = (Get-DnsServerZoneAging -Name $DNSRoot -ComputerName $PDCEmulator).AgingEnabled
#endregion
 
#region DHCP Info
 
######################################
# DHCP Information
######################################
$DHCP = Get-WindowsFeature -name DHCP | 
        Where-Object {$_.Installed -eq $True}
 
If($DHCP){
  $DHCPServers =  Get-DhcpServerInDC
 
  $DHCPOutputObj  = New-Object -TypeName PSObject
  $DHCPOutputObj | Add-Member -MemberType NoteProperty -Name Name -Value $DHCPServers.DNSName
  $DHCPOutputObj | Add-Member -MemberType NoteProperty -Name IPAddress -Value $DHCPServers.IPAddress
  $DHCPOutputObjHTML = $DHCPOutputObj | ConvertTo-Html
} Else {
  $DHCPServers =  'No DHCP Found'
    
  $DHCPOutputObj  = New-Object -TypeName PSObject
  $DHCPOutputObj | Add-Member -MemberType NoteProperty -Name -- -Value $DHCPServers
  $DHCPOutputObjHTML = $DHCPOutputObj | ConvertTo-Html
}
#endregion
 
#region Site Info
 
######################################
# Site Information
######################################
# All Forest Sites
$Sites = ($forestInfo | 
         Select-Object -ExpandProperty Sites) -join '<br/>'
 
 # Inter-Site Transport
  ####### Need a foreachloop for each sites info
  $SiteLinkNames = $ADSiteLinks.Name
  $SitesInlcuded = ($ADSiteLinks | Select-Object -ExpandProperty SitesIncluded) -join ' | '
  $SiteCost = ($ADSiteLinks | Select-Object -ExpandProperty Cost) -join '<br/>'
  $SiteReplicationFreq = ($ADSiteLinks | Select-Object -ExpandProperty ReplicationFrequencyInMinutes) -join '<br/>' 
  
  # Create a custom object from the values above and convert it to an html table
  $SiteLinkObj  = New-Object -TypeName PSObject
  $SiteLinkObj | Add-Member -MemberType NoteProperty -Name SiteName -Value $SiteLinkNames
  $SiteLinkObj | Add-Member -MemberType NoteProperty -Name SitesIncluded -Value $SitesInlcuded
  $SiteLinkObj | Add-Member -MemberType NoteProperty -Name SiteCost -Value $SiteCost
  $SiteLinkObj | Add-Member -MemberType NoteProperty -Name SiteReplicationFreq -Value $SiteReplicationFreq
  $SiteLinkObjHTML = $SiteLinkObj | ConvertTo-Html
#endregion
 
#region GPO Info
 
######################################
# GPO Information
######################################
$DomainGPOs = Get-GPO -all | Select-Object -Property DomainName,DisplayName,CreationTime,ModificationTime
$GPOInfo = $DomainGPOs | ConvertTo-Html 
#endregion
 
#region Priviledged Account Info 
 
######################################
# Priviledged Account Information
######################################
# Priviledge Group Membership
  $DomainAdmins = (Get-ADGroupMember -Identity 'Domain Admins' | Select-Object -ExpandProperty SamAccountName) -join '<br/>'
  $EnterpriseAdmins = (Get-ADGroupMember -Identity 'Enterprise Admins'  | Select-Object -ExpandProperty SamAccountName) -join '<br/>'
  $SchemaAdmins = (Get-ADGroupMember -Identity 'Schema Admins'  | Select-Object -ExpandProperty SamAccountName) -join '<br/>'
#endregion
 
#region Exchange Info
 
######################################
# Exchange Information
######################################
# Get all Org Management Users
  $OrgManagement = (Get-ADGroupMember -Identity 'Organization Management' | Select-Object -ExpandProperty SamAccountName) -join '<br/>'
# Get all Exchange Servers
  $ExchangeSVRs = (Get-ADGroupMember -Identity 'Exchange Servers' | Select-Object -ExpandProperty SamAccountName) -join '<br/>'
#endregion
 
#region User Info
 
######################################
# User Information
######################################
# Users with Passwords set to never expire
  $NeverExpire = (Get-ADUser -Filter {PasswordNeverExpires -eq $true} | Select-Object -ExpandProperty SamAccountName) -join '<br/>'
#endregion
 
#region HTML Output
 
######################################
# HTML Output
######################################
$Create_HTML_doc = "<!DOCTYPE html>
  <head>
  <title>Active Directory Information</title>
  <style>
 
  BODY{
    font-family: Arial, Verdana;
    background-color:#D8D8D8;
  }
  TABLE{
    border=1; 
    border-color:black; 
    border-width:1px; 
    border-style:solid;
    border-collapse: collapse; 
    empty-cells:show
  }
  TH{
    font-size: 15px;
    border-width:1px; 
    padding:5px; 
    border-style:solid; 
    font-weight:bold; 
    text-align:left;
    border-color:black;
    background-color:#ffffff;
    empty-cells:show
  }
  TD{
    font-size: 12px;
    color:black; 
    colspan=1; 
    border-width:1px; 
    padding:5px; 
    font-weight:normal; 
    border-style:solid;
    border-color:black;
    background-color:#ffffff;
    vertical-align: top;
    empty-cells:show
  }
  h1{
    font-size: 35px;
  }
  h2{
    font-size: 30px;
    text-decoration: underline;
    
  }
  h3{
    font-size: 15px;
  }
  </style>
  </head>
  <h1> Active Directory Information for :  $DNSRoot </h1>
 
  <h2> Forest Information </h2> 
  $ForestOutputObjHTML
  <br/>
 
  <h2> Domain Information </h2> 
  $DomainOutputObjHTML
  <br/>
 
  <h2> FSMO Information </h2> 
  <table>
  <tr>
    <td><h3>Forest FSMO Roles</h3></td>
    <td><h3>Domain FSMO Roles</h3></td>
  </tr>
  <tr>
    <td>$ForestFSMO</td>
    <td>$DomainFSMO</td>
  </tr>
  </table>
 
  <h2> Domain Controller Information </h2>
  $DCOutputObjHTML
  <br/>
 
  <h2> DNS Information </h2>
  <table>
  <tr>
    <td><h3>Primary Zones</h3></td>
    <td><h3>NS Records</h3></td>
    <td><h3>MX Records</h3></td>
    <td><h3>Forwarders</h3></td>
    <td><h3>Scavenging Enabled?</h3></td>
    <td><h3>Aging Enabled?</h3></td>
  </tr>
  <tr>
    <td>$PrimaryZones</td>
    <td>$NSRecords</td>
    <td>$MXRecords</td>
    <td>$DNSForwarders</td>
    <td>$DNSScavenging</td>
    <td>$DNSAging</td>
  </tr>
  </table>
 
  <h2> DHCP Information </h2> 
  $DHCPOutputObjHTML
 
  <h2> AD Site Information </h2>
  <table>
  <tr>
    <td><h3>Forest Wide Sites</h3></td>
    <td><h3>Site Links</h3></td>
  </tr>
  <tr>
    <td>$Sites</td>
    <td>$SiteLinkObjHTML</td>
  </tr>
  </table>
 
  <h2> GPO Information </h2> 
  $GPOInfo
 
  <h2>Priviledged Accounts</h2>
  <table>
  <tr>
    <td><h3>Enterprise Admin Group Members</h3></td>
    <td><h3>Domain Admin Group Members</h3></td>
    <td><h3>Schema Admin Group Members</h3></td>
    <td><h3>Password Never Expire</h3></td>
  </tr>
  <tr>
    <td>$EnterpriseAdmins</td>
    <td>$DomainAdmins</td>
    <td>$SchemaAdmins</td>
    <td>$NeverExpire</td>
  </tr>
  </table>
 
  <h2>Exchange Information</h2>
  <table>
  <tr>
    <td><h3>Organization Management Group Members</h3></td>
    <td><h3>Exchange Servers</h3></td>
  </tr>
  <tr>
    <td>$OrgManagement</td>
    <td>$ExchangeSVRs</td>
  </tr>
  </table>
"
$Create_HTML_doc > $File
#endregion
 
# This is optional, it just opens the html file after the script runs
Invoke-Item -Path $File