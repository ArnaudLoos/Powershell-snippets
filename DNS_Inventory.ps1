#Requires -Version 3.0
#Requires -Module dnsserver
#This File is in Unicode format.  Do not edit in an ASCII editor.
# This is version 1.10 from April 2018 from carlwebster.com

#region help text

<#
.SYNOPSIS
	Creates an inventory of Microsoft DNS using Microsoft Word, PDF, formatted text or HTML.
.DESCRIPTION
	Creates an inventory of Microsoft DNS using Microsoft Word, PDF, formatted text or HTML.
	Creates a document named DNS.docx (or .PDF or .TXT or .HTML).
	Word and PDF documents include a Cover Page, Table of Contents and Footer.
	Includes support for the following language versions of Microsoft Word:
		Catalan
		Chinese
		Danish
		Dutch
		English
		Finnish
		French
		German
		Norwegian
		Portuguese
		Spanish
		Swedish

	To run the script from a workstation, RSAT is required.
	
	Remote Server Administration Tools for Windows 7 with Service Pack 1 (SP1)
		http://www.microsoft.com/en-us/download/details.aspx?id=7887
		
	Remote Server Administration Tools for Windows 8 
		http://www.microsoft.com/en-us/download/details.aspx?id=28972
		
	Remote Server Administration Tools for Windows 8.1 
		http://www.microsoft.com/en-us/download/details.aspx?id=39296
		
	Remote Server Administration Tools for Windows 10
		http://www.microsoft.com/en-us/download/details.aspx?id=45520
		
.PARAMETER CompanyAddress
	Company Address to use for the Cover Page, if the Cover Page has the Address field.
	
	The following Cover Pages have an Address field:
		Banded (Word 2013/2016)
		Contrast (Word 2010)
		Exposure (Word 2010)
		Filigree (Word 2013/2016)
		Ion (Dark) (Word 2013/2016)
		Retrospect (Word 2013/2016)
		Semaphore (Word 2013/2016)
		Tiles (Word 2010)
		ViewMaster (Word 2013/2016)
		
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CA.
.PARAMETER CompanyEmail
	Company Email to use for the Cover Page, if the Cover Page has the Email field.  
	
	The following Cover Pages have an Email field:
		Facet (Word 2013/2016)
	
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CE.
.PARAMETER CompanyFax
	Company Fax to use for the Cover Page, if the Cover Page has the Fax field.  
	
	The following Cover Pages have a Fax field:
		Contrast (Word 2010)
		Exposure (Word 2010)
	
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CF.
.PARAMETER CompanyName
	Company Name to use for the Cover Page.  
	The default value is contained in 
	HKCU:\Software\Microsoft\Office\Common\UserInfo\CompanyName or
	HKCU:\Software\Microsoft\Office\Common\UserInfo\Company, whichever is populated 
	on the computer running the script.

	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CN.
.PARAMETER CompanyPhone
	Company Phone to use for the Cover Page, if the Cover Page has the Phone field.  
	
	The following Cover Pages have a Phone field:
		Contrast (Word 2010)
		Exposure (Word 2010)
	
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CPh.
.PARAMETER CoverPage
	What Microsoft Word Cover Page to use.
	Only Word 2010, 2013 and 2016 are supported.
	(default cover pages in Word en-US)
	
	Valid input is:
		Alphabet (Word 2010. Works)
		Annual (Word 2010. Doesn't work well for this report)
		Austere (Word 2010. Works)
		Austin (Word 2010/2013/2016. Doesn't work in 2013 or 2016, mostly 
		works in 2010 but Subtitle/Subject & Author fields need to be moved 
		after title box is moved up)
		Banded (Word 2013/2016. Works)
		Conservative (Word 2010. Works)
		Contrast (Word 2010. Works)
		Cubicles (Word 2010. Works)
		Exposure (Word 2010. Works if you like looking sideways)
		Facet (Word 2013/2016. Works)
		Filigree (Word 2013/2016. Works)
		Grid (Word 2010/2013/2016. Works in 2010)
		Integral (Word 2013/2016. Works)
		Ion (Dark) (Word 2013/2016. Top date doesn't fit; box needs to be 
		manually resized or font changed to 8 point)
		Ion (Light) (Word 2013/2016. Top date doesn't fit; box needs to be 
		manually resized or font changed to 8 point)
		Mod (Word 2010. Works)
		Motion (Word 2010/2013/2016. Works if top date is manually changed to 
		36 point)
		Newsprint (Word 2010. Works but date is not populated)
		Perspective (Word 2010. Works)
		Pinstripes (Word 2010. Works)
		Puzzle (Word 2010. Top date doesn't fit; box needs to be manually 
		resized or font changed to 14 point)
		Retrospect (Word 2013/2016. Works)
		Semaphore (Word 2013/2016. Works)
		Sideline (Word 2010/2013/2016. Doesn't work in 2013 or 2016, works in 
		2010)
		Slice (Dark) (Word 2013/2016. Doesn't work)
		Slice (Light) (Word 2013/2016. Doesn't work)
		Stacks (Word 2010. Works)
		Tiles (Word 2010. Date doesn't fit unless changed to 26 point)
		Transcend (Word 2010. Works)
		ViewMaster (Word 2013/2016. Works)
		Whisp (Word 2013/2016. Works)
		
	Default value is Sideline.
	This parameter has an alias of CP.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER UserName
	User name to use for the Cover Page and Footer.
	Default value is contained in $env:username
	This parameter has an alias of UN.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER HTML
	Creates an HTML file with an .html extension.
	This parameter is disabled by default.
.PARAMETER MSWord
	SaveAs DOCX file
	This parameter is set True if no other output format is selected.
.PARAMETER PDF
	SaveAs PDF file instead of DOCX file.
	This parameter is disabled by default.
	The PDF file is roughly 5X to 10X larger than the DOCX file.
	This parameter requires Microsoft Word to be installed.
	This parameter uses the Word SaveAs PDF capability.
.PARAMETER Text
	Creates a formatted text file with a .txt extension.
	This parameter is disabled by default.
.PARAMETER AddDateTime
	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2018 at 6PM is 2018-06-01_1800.
	Output filename will be DomainName_DNS_2018-06-01_1800.docx (or .pdf).
	This parameter is disabled by default.
.PARAMETER ComputerName
	Specifies a computer to use to run the script against.
	ComputerName can be entered as the NetBIOS name, FQDN, localhost or IP Address.
	If entered as localhost, the actual computer name is determined and used.
	If entered as an IP address, an attempt is made to determine and use the actual 
	computer name.
	Default is localhost.
.PARAMETER Folder
	Specifies the optional output folder to save the output report. 
.PARAMETER SmtpServer
	Specifies the optional email server to send the output report. 
.PARAMETER SmtpPort
	Specifies the SMTP port. 
	Default is 25.
.PARAMETER UseSSL
	Specifies whether to use SSL for the SmtpServer.
	Default is False.
.PARAMETER From
	Specifies the username for the From email address.
	If SmtpServer is used, this is a required parameter.
.PARAMETER To
	Specifies the username for the To email address.
	If SmtpServer is used, this is a required parameter.
.PARAMETER Dev
	Clears errors at the beginning of the script.
	Outputs all errors to a text file at the end of the script.
	
	This is used when the script developer requests more troubleshooting data.
	Text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
.PARAMETER ScriptInfo
	Outputs information about the script to a text file.
	Text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
	This parameter has an alias of SI.
.PARAMETER Details
	Include Resource Record data for both Forward and Reverse lookup zones.
	Default is to not include Resource Record information.
.PARAMETER Log
	Generates a log file for troubleshooting.
.EXAMPLE
	PS C:\PSScript > .\DNS_Inventory.ps1
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Tests to see if the computer, localhost, is a DNS server. 
	If it is, the script runs. If not, the script aborts.
.EXAMPLE
	PS C:\PSScript > .\DNS_Inventory.ps1 -ComputerName DNS01
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Runs the script against the DNS server named DNS01.
.EXAMPLE
	PS C:\PSScript > .\DNS_Inventory.ps1 -PDF
	
	Will use all default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\DNS_Inventory.ps1 -TEXT

	Will use all default values and save the document as a formatted text file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\DNS_Inventory.ps1 -HTML

	Will use all default values and save the document as an HTML file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript .\DNS_Inventory.ps1 -CompanyName "Carl Webster Consulting" 
	-CoverPage "Mod" -UserName "Carl Webster"

	Will use:
		Carl Webster Consulting for the Company Name.
		Mod for the Cover Page format.
		Carl Webster for the User Name.
.EXAMPLE
	PS C:\PSScript .\DNS_Inventory.ps1 -CN "Carl Webster Consulting" -CP "Mod" -UN 
	"Carl Webster"

	Will use:
		Carl Webster Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Carl Webster for the User Name (alias UN).
.EXAMPLE
	PS C:\PSScript > .\DNS_Inventory.ps1 -AddDateTime
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	July 25, 2018 at 6PM is 2018-07-25_1800.
	Output filename will be DomainName_DNS_2018-07-25_1800.docx
.EXAMPLE
	PS C:\PSScript > .\DNS_Inventory.ps1 -PDF -AddDateTime
	
	Will use all default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	July 25, 2018 at 6PM is 2018-07-25_1800.
	Output filename will be DomainName_DNS_2018-07-25_1800.PDF
.EXAMPLE
	PS C:\PSScript > .\DNS_Inventory.ps1 -Folder \\FileServer\ShareName
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Output file is saved in the path \\FileServer\ShareName
.EXAMPLE
	PS C:\PSScript > .\DNS_Inventory.ps1 -SmtpServer mail.domain.tld -From 
	XDAdmin@domain.tld -To ITGroup@domain.tld -ComputerName Server01
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Script will be run remotely against DNS server Server01.
	
	Script will use the email server mail.domain.tld, sending from XDAdmin@domain.tld, 
	sending to ITGroup@domain.tld.
	Script will use the default SMTP port 25 and will not use SSL.
	If the current user's credentials are not valid to send email, the user will be prompted 
	to enter valid credentials.
.EXAMPLE
	PS C:\PSScript > .\DNS_Inventory.ps1 -SmtpServer smtp.office365.com -SmtpPort 587 
	-UseSSL -From Webster@CarlWebster.com -To ITGroup@CarlWebster.com
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Script will use the email server smtp.office365.com on port 587 using SSL, sending from 
	webster@carlwebster.com, sending to ITGroup@carlwebster.com.
	If the current user's credentials are not valid to send email, the user will be prompted
	to enter valid credentials.
.EXAMPLE
	PS C:\PSScript > .\DNS_Inventory.ps1 -Details
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Includes details for all Resource Records for both Forward and Reverse lookup zones.
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  
	This script creates a Word, PDF, Formatted Text or HTML document.
.NOTES
	NAME: DNS_Inventory.ps1
	VERSION: 1.10
	AUTHOR: Carl Webster - Sr. Solutions Architect - Choice Solutions, LLC
	LASTEDIT: April 6, 2018
#>

#endregion

#region script parameters
#thanks to @jeffwouters and Michael B. Smith for helping me with these parameters
[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "Word") ]

Param(
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Alias("CA")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyAddress="",
    
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Alias("CE")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyEmail="",
    
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Alias("CF")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyFax="",
    
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Alias("CN")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyName="",
    
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Alias("CPh")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyPhone="",
    
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Alias("CP")]
	[ValidateNotNullOrEmpty()]
	[string]$CoverPage="Sideline", 

	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Alias("UN")]
	[ValidateNotNullOrEmpty()]
	[string]$UserName=$env:username,

	[parameter(ParameterSetName="HTML",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Switch]$HTML=$False,

	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Switch]$MSWord=$False,

	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Switch]$PDF=$False,

	[parameter(ParameterSetName="Text",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Switch]$Text=$False,

	[parameter(Mandatory=$False)] 
	[Switch]$AddDateTime=$False,
	
	[parameter(Mandatory=$False)] 
	[string]$ComputerName="LocalHost",

	[parameter(Mandatory=$False)] 
	[string]$Folder="",
	
	[parameter(ParameterSetName="SMTP",Mandatory=$True)] 
	[string]$SmtpServer="",

	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[int]$SmtpPort=25,

	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[switch]$UseSSL=$False,

	[parameter(ParameterSetName="SMTP",Mandatory=$True)] 
	[string]$From="",

	[parameter(ParameterSetName="SMTP",Mandatory=$True)] 
	[string]$To="",

	[parameter(Mandatory=$False)] 
	[Switch]$Dev=$False,
	
	[parameter(Mandatory=$False)] 
	[Alias("SI")]
	[Switch]$ScriptInfo=$False,
	
	[parameter(Mandatory=$False)] 
	[Switch]$Details=$False,
	
	[parameter(Mandatory=$False)] 
	[Switch]$Log=$False
	
	)
#endregion

#region script change log	
#Created by Carl Webster
#Sr. Solutions Architect, Choice Solutions, LLC
#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com
#Created on February 10, 2016
#Version 1.00 released to the community on July 25, 2016

#Version 1.01 16-Aug-2016
#	Added support for the four Record Types created by implementing DNSSEC
#		NSec
#		NSec3
#		NSec3Param
#		RRSig
#
#Version 1.02 19-Aug-2016
#	Fixed several misspelled words
#
#Version 1.03 19-Oct-2016
#	Fixed formatting issues with HTML headings output
#
#Version 1.04 22-Oct-2016
#	More refinement of HTML output
#
#Version 1.05 7-Nov-2016
#	Added Chinese language support
#
#Version 1.06 13-Feb-2017
#	Fixed French wording for Table of Contents 2 (Thanks to David Rouquier)
#
#Version 1.07 13-Nov-2017
#	Added Scavenge Server(s) to Zone Properties General section
#	Added the domain name of the computer used for -ComputerName to the output filename
#	Fixed output of Name Server IP address(es) in Zone properties
#	For Word/PDF output added the domain name of the computer used for -ComputerName to the report title
#	General code cleanup
#	In Text output, fixed alignment of "Scavenging period" in DNS Server Properties
#	Removed code that made sure all Parameters were set to default values if for some reason they did exist or values were $Null
#	Reordered the parameters in the help text and parameter list so they match and are grouped better
#	Replaced _SetDocumentProperty function with Jim Moyle's Set-DocumentProperty function
#	Updated Function ProcessScriptEnd for the new Cover Page properties and Parameters
#	Updated Function ShowScriptOptions for the new Cover Page properties and Parameters
#	Updated Function UpdateDocumentProperties for the new Cover Page properties and Parameters
#	Updated help text
#
#Version 1.08 8-Dec-2017
#	Updated Function WriteHTMLLine with fixes from the script template
#
#Version 1.09 2-Mar-2018
#	Added Log switch to create a transcript log
#	I found two "If($Var = something)" which are now "If($Var -eq something)"
#	In the function OutputLookupZoneDetails, with the "=" changed to "-eq" fix, the hostname was now always blank. Fixed.
#	Many Switch bocks I never added "; break" to. Those are now fixed.
#	Update functions ShowScriptOutput and ProcessScriptEnd for new Log parameter
#	Updated help text
#	Updated the WriteWordLine function 
#
#Version 1.10 6-Apr-2018
#	Code clean up from Visual Studio Code
#
#HTML functions contributed by Ken Avram October 2014
#HTML Functions FormatHTMLTable and AddHTMLTable modified by Jake Rutski May 2015
#endregion

#region initial variable testing and setup
Set-StrictMode -Version 2

#force  on
$PSDefaultParameterValues = @{"*:Verbose"=$True}
$SaveEAPreference = $ErrorActionPreference
$ErrorActionPreference = 'SilentlyContinue'

#V1.09 added
If($Log) 
{
	#start transcript logging
	$Script:ThisScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
	$Script:LogPath = "$Script:ThisScriptPath\DNSDocScriptTranscript_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
	
	try 
	{
		Start-Transcript -Path $Script:LogPath -Force -Verbose:$false | Out-Null
		Write-Verbose "$(Get-Date): Transcript/log started at $Script:LogPath"
		$Script:StartLog = $true
	} 
	catch 
	{
		Write-Verbose "$(Get-Date): Transcript/log failed at $Script:LogPath"
		$Script:StartLog = $false
	}
}

If($Dev)
{
	$Error.Clear()
	$Script:DevErrorFile = "$($pwd.Path)\DNSInventoryScriptErrors_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
}

If($Null -eq $MSWord)
{
	If($Text -or $HTML -or $PDF)
	{
		$MSWord = $False
	}
	Else
	{
		$MSWord = $True
	}
}

If($MSWord -eq $False -and $PDF -eq $False -and $Text -eq $False -and $HTML -eq $False)
{
	$MSWord = $True
}

Write-Verbose "$(Get-Date): Testing output parameters"

If($MSWord)
{
	Write-Verbose "$(Get-Date): MSWord is set"
}
ElseIf($PDF)
{
	Write-Verbose "$(Get-Date): PDF is set"
}
ElseIf($Text)
{
	Write-Verbose "$(Get-Date): Text is set"
}
ElseIf($HTML)
{
	Write-Verbose "$(Get-Date): HTML is set"
}
Else
{
	$ErrorActionPreference = $SaveEAPreference
	Write-Verbose "$(Get-Date): Unable to determine output parameter"
	If($Null -eq $MSWord)
	{
		Write-Verbose "$(Get-Date): MSWord is Null"
	}
	ElseIf($Null -eq $PDF)
	{
		Write-Verbose "$(Get-Date): PDF is Null"
	}
	ElseIf($Null -eq $Text)
	{
		Write-Verbose "$(Get-Date): Text is Null"
	}
	ElseIf($Null -eq $HTML)
	{
		Write-Verbose "$(Get-Date): HTML is Null"
	}
	Else
	{
		Write-Verbose "$(Get-Date): MSWord is $MSWord"
		Write-Verbose "$(Get-Date): PDF is $PDF"
		Write-Verbose "$(Get-Date): Text is $Text"
		Write-Verbose "$(Get-Date): HTML is $HTML"
	}
	Write-Error "Unable to determine output parameter.  Script cannot continue"
	Exit
}

If($Folder -ne "")
{
	Write-Verbose "$(Get-Date): Testing folder path"
	#does it exist
	If(Test-Path $Folder -EA 0)
	{
		#it exists, now check to see if it is a folder and not a file
		If(Test-Path $Folder -pathType Container -EA 0)
		{
			#it exists and it is a folder
			Write-Verbose "$(Get-Date): Folder path $Folder exists and is a folder"
		}
		Else
		{
			#it exists but it is a file not a folder
			Write-Error "Folder $Folder is a file, not a folder.  Script cannot continue"
			Exit
		}
	}
	Else
	{
		#does not exist
		Write-Error "Folder $Folder does not exist.  Script cannot continue"
		Exit
	}
}

#endregion

#region initialize variables for word html and text
[string]$Script:RunningOS = (Get-WmiObject -class Win32_OperatingSystem -EA 0).Caption

If($MSWord -or $PDF)
{
	#try and fix the issue with the $CompanyName variable
	$Script:CoName = $CompanyName
	Write-Verbose "$(Get-Date): CoName is $Script:CoName"
	
	#the following values were attained from 
	#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/
	#http://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx
	[int]$wdAlignPageNumberRight = 2
	[long]$wdColorGray15 = 14277081
	[long]$wdColorGray05 = 15987699 
	[int]$wdMove = 0
	[int]$wdSeekMainDocument = 0
	[int]$wdSeekPrimaryFooter = 4
	[int]$wdStory = 6
	[long]$wdColorRed = 255
	[int]$wdColorBlack = 0
	[int]$wdWord2007 = 12
	[int]$wdWord2010 = 14
	[int]$wdWord2013 = 15
	[int]$wdWord2016 = 16
	[int]$wdFormatDocumentDefault = 16
	[int]$wdFormatPDF = 17
	#http://blogs.technet.com/b/heyscriptingguy/archive/2006/03/01/how-can-i-right-align-a-single-column-in-a-word-table.aspx
	#http://msdn.microsoft.com/en-us/library/office/ff835817%28v=office.15%29.aspx
	[int]$wdAlignParagraphLeft = 0
	[int]$wdAlignParagraphCenter = 1
	[int]$wdAlignParagraphRight = 2
	#http://msdn.microsoft.com/en-us/library/office/ff193345%28v=office.15%29.aspx
	[int]$wdCellAlignVerticalTop = 0
	[int]$wdCellAlignVerticalCenter = 1
	[int]$wdCellAlignVerticalBottom = 2
	#http://msdn.microsoft.com/en-us/library/office/ff844856%28v=office.15%29.aspx
	[int]$wdAutoFitFixed = 0
	[int]$wdAutoFitContent = 1
	[int]$wdAutoFitWindow = 2
	#http://msdn.microsoft.com/en-us/library/office/ff821928%28v=office.15%29.aspx
	[int]$wdAdjustNone = 0
	[int]$wdAdjustProportional = 1
	[int]$wdAdjustFirstColumn = 2
	[int]$wdAdjustSameWidth = 3

	[int]$PointsPerTabStop = 36
	[int]$Indent0TabStops = 0 * $PointsPerTabStop
	[int]$Indent1TabStops = 1 * $PointsPerTabStop
	[int]$Indent2TabStops = 2 * $PointsPerTabStop
	[int]$Indent3TabStops = 3 * $PointsPerTabStop
	[int]$Indent4TabStops = 4 * $PointsPerTabStop

	# http://www.thedoctools.com/index.php?show=wt_style_names_english_danish_german_french
	[int]$wdStyleHeading1 = -2
	[int]$wdStyleHeading2 = -3
	[int]$wdStyleHeading3 = -4
	[int]$wdStyleHeading4 = -5
	[int]$wdStyleNoSpacing = -158
	[int]$wdTableGrid = -155

	#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/org/codehaus/groovy/scriptom/tlb/office/word/WdLineStyle.html
	[int]$wdLineStyleNone = 0
	[int]$wdLineStyleSingle = 1

	[int]$wdHeadingFormatTrue = -1
	[int]$wdHeadingFormatFalse = 0 
}

If($HTML)
{
    Set-Variable htmlredmask         -Option AllScope -Value "#FF0000" 4>$Null
    Set-Variable htmlcyanmask        -Option AllScope -Value "#00FFFF" 4>$Null
    Set-Variable htmlbluemask        -Option AllScope -Value "#0000FF" 4>$Null
    Set-Variable htmldarkbluemask    -Option AllScope -Value "#0000A0" 4>$Null
    Set-Variable htmllightbluemask   -Option AllScope -Value "#ADD8E6" 4>$Null
    Set-Variable htmlpurplemask      -Option AllScope -Value "#800080" 4>$Null
    Set-Variable htmlyellowmask      -Option AllScope -Value "#FFFF00" 4>$Null
    Set-Variable htmllimemask        -Option AllScope -Value "#00FF00" 4>$Null
    Set-Variable htmlmagentamask     -Option AllScope -Value "#FF00FF" 4>$Null
    Set-Variable htmlwhitemask       -Option AllScope -Value "#FFFFFF" 4>$Null
    Set-Variable htmlsilvermask      -Option AllScope -Value "#C0C0C0" 4>$Null
    Set-Variable htmlgraymask        -Option AllScope -Value "#808080" 4>$Null
    Set-Variable htmlblackmask       -Option AllScope -Value "#000000" 4>$Null
    Set-Variable htmlorangemask      -Option AllScope -Value "#FFA500" 4>$Null
    Set-Variable htmlmaroonmask      -Option AllScope -Value "#800000" 4>$Null
    Set-Variable htmlgreenmask       -Option AllScope -Value "#008000" 4>$Null
    Set-Variable htmlolivemask       -Option AllScope -Value "#808000" 4>$Null

    Set-Variable htmlbold        -Option AllScope -Value 1 4>$Null
    Set-Variable htmlitalics     -Option AllScope -Value 2 4>$Null
    Set-Variable htmlred         -Option AllScope -Value 4 4>$Null
    Set-Variable htmlcyan        -Option AllScope -Value 8 4>$Null
    Set-Variable htmlblue        -Option AllScope -Value 16 4>$Null
    Set-Variable htmldarkblue    -Option AllScope -Value 32 4>$Null
    Set-Variable htmllightblue   -Option AllScope -Value 64 4>$Null
    Set-Variable htmlpurple      -Option AllScope -Value 128 4>$Null
    Set-Variable htmlyellow      -Option AllScope -Value 256 4>$Null
    Set-Variable htmllime        -Option AllScope -Value 512 4>$Null
    Set-Variable htmlmagenta     -Option AllScope -Value 1024 4>$Null
    Set-Variable htmlwhite       -Option AllScope -Value 2048 4>$Null
    Set-Variable htmlsilver      -Option AllScope -Value 4096 4>$Null
    Set-Variable htmlgray        -Option AllScope -Value 8192 4>$Null
    Set-Variable htmlolive       -Option AllScope -Value 16384 4>$Null
    Set-Variable htmlorange      -Option AllScope -Value 32768 4>$Null
    Set-Variable htmlmaroon      -Option AllScope -Value 65536 4>$Null
    Set-Variable htmlgreen       -Option AllScope -Value 131072 4>$Null
    Set-Variable htmlblack       -Option AllScope -Value 262144 4>$Null
}

If($TEXT)
{
	$global:output = ""
}
#endregion

#region word specific functions
Function SetWordHashTable
{
	Param([string]$CultureCode)

	#optimized by Michael B. SMith
	
	# DE and FR translations for Word 2010 by Vladimir Radojevic
	# Vladimir.Radojevic@Commerzreal.com

	# DA translations for Word 2010 by Thomas Daugaard
	# Citrix Infrastructure Specialist at edgemo A/S

	# CA translations by Javier Sanchez 
	# CEO & Founder 101 Consulting

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish
	#zh - Chinese
	
	[string]$toc = $(
		Switch ($CultureCode)
		{
			'ca-'	{ 'Taula automática 2'; Break }
			'da-'	{ 'Automatisk tabel 2'; Break }
			'de-'	{ 'Automatische Tabelle 2'; Break }
			'en-'	{ 'Automatic Table 2'; Break }
			'es-'	{ 'Tabla automática 2'; Break }
			'fi-'	{ 'Automaattinen taulukko 2'; Break }
			'fr-'	{ 'Table automatique 2'; Break } #changed 13-feb-2017 david roquier and samuel legrand
			'nb-'	{ 'Automatisk tabell 2'; Break }
			'nl-'	{ 'Automatische inhoudsopgave 2'; Break }
			'pt-'	{ 'Sumário Automático 2'; Break }
			'sv-'	{ 'Automatisk innehållsförteckning2'; Break }
			'zh-'	{ '自动目录 2'; Break }
		}
	)

	$Script:myHash                      = @{}
	$Script:myHash.Word_TableOfContents = $toc
	$Script:myHash.Word_NoSpacing       = $wdStyleNoSpacing
	$Script:myHash.Word_Heading1        = $wdStyleheading1
	$Script:myHash.Word_Heading2        = $wdStyleheading2
	$Script:myHash.Word_Heading3        = $wdStyleheading3
	$Script:myHash.Word_Heading4        = $wdStyleheading4
	$Script:myHash.Word_TableGrid       = $wdTableGrid
}

Function GetCulture
{
	Param([int]$WordValue)
	
	#codes obtained from http://support.microsoft.com/kb/221435
	#http://msdn.microsoft.com/en-us/library/bb213877(v=office.12).aspx
	$CatalanArray = 1027
	$ChineseArray = 2052,3076,5124,4100
	$DanishArray = 1030
	$DutchArray = 2067, 1043
	$EnglishArray = 3081, 10249, 4105, 9225, 6153, 8201, 5129, 13321, 7177, 11273, 2057, 1033, 12297
	$FinnishArray = 1035
	$FrenchArray = 2060, 1036, 11276, 3084, 12300, 5132, 13324, 6156, 8204, 10252, 7180, 9228, 4108
	$GermanArray = 1031, 3079, 5127, 4103, 2055
	$NorwegianArray = 1044, 2068
	$PortugueseArray = 1046, 2070
	$SpanishArray = 1034, 11274, 16394, 13322, 9226, 5130, 7178, 12298, 17418, 4106, 18442, 19466, 6154, 15370, 10250, 20490, 3082, 14346, 8202
	$SwedishArray = 1053, 2077

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish
	#zh - Chinese

	Switch ($WordValue)
	{
		{$CatalanArray -contains $_} {$CultureCode = "ca-"}
		{$ChineseArray -contains $_} {$CultureCode = "zh-"}
		{$DanishArray -contains $_} {$CultureCode = "da-"}
		{$DutchArray -contains $_} {$CultureCode = "nl-"}
		{$EnglishArray -contains $_} {$CultureCode = "en-"}
		{$FinnishArray -contains $_} {$CultureCode = "fi-"}
		{$FrenchArray -contains $_} {$CultureCode = "fr-"}
		{$GermanArray -contains $_} {$CultureCode = "de-"}
		{$NorwegianArray -contains $_} {$CultureCode = "nb-"}
		{$PortugueseArray -contains $_} {$CultureCode = "pt-"}
		{$SpanishArray -contains $_} {$CultureCode = "es-"}
		{$SwedishArray -contains $_} {$CultureCode = "sv-"}
		Default {$CultureCode = "en-"}
	}
	
	Return $CultureCode
}

Function ValidateCoverPage
{
	Param([int]$xWordVersion, [string]$xCP, [string]$CultureCode)
	
	$xArray = ""
	
	Switch ($CultureCode)
	{
		'ca-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "Semàfor", "Visualització principal", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "Semàfor", "Visualització", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Anual", "Austin", "Conservador",
					"Contrast", "Cubicles", "Diplomàtic", "Exposició",
					"Línia lateral", "Mod", "Mosiac", "Moviment", "Paper de diari",
					"Perspectiva", "Piles", "Quadrícula", "Sobri",
					"Transcendir", "Trencaclosques")
				}
			}

		'da-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "BevægElse", "Brusen", "Facet", "Filigran", 
					"Gitter", "Integral", "Ion (lys)", "Ion (mørk)", 
					"Retro", "Semafor", "Sidelinje", "Stribet", 
					"Udsnit (lys)", "Udsnit (mørk)", "Visningsmaster")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("BevægElse", "Brusen", "Ion (lys)", "Filigran",
					"Retro", "Semafor", "Visningsmaster", "Integral",
					"Facet", "Gitter", "Stribet", "Sidelinje", "Udsnit (lys)",
					"Udsnit (mørk)", "Ion (mørk)", "Austin")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("BevægElse", "Moderat", "Perspektiv", "Firkanter",
					"Overskrid", "Alfabet", "Kontrast", "Stakke", "Fliser", "Gåde",
					"Gitter", "Austin", "Eksponering", "Sidelinje", "Enkel",
					"Nålestribet", "Årlig", "Avispapir", "Tradionel")
				}
			}

		'de-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Bewegung", "Facette", "Filigran", 
					"Gebändert", "Integral", "Ion (dunkel)", "Ion (hell)", 
					"Pfiff", "Randlinie", "Raster", "Rückblick", 
					"Segment (dunkel)", "Segment (hell)", "Semaphor", 
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Semaphor", "Segment (hell)", "Ion (hell)",
					"Raster", "Ion (dunkel)", "Filigran", "Rückblick", "Pfiff",
					"ViewMaster", "Segment (dunkel)", "Verbunden", "Bewegung",
					"Randlinie", "Austin", "Integral", "Facette")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Austin", "Bewegung", "Durchscheinend",
					"Herausgestellt", "Jährlich", "Kacheln", "Kontrast", "Kubistisch",
					"Modern", "Nadelstreifen", "Perspektive", "Puzzle", "Randlinie",
					"Raster", "Schlicht", "Stapel", "Traditionell", "Zeitungspapier")
				}
			}

		'en-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
					"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
					"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
					"Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
					"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
					"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
				}
			}

		'es-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Con bandas", "Cortar (oscuro)", "Cuadrícula", 
					"Whisp", "Faceta", "Filigrana", "Integral", "Ion (claro)", 
					"Ion (oscuro)", "Línea lateral", "Movimiento", "Retrospectiva", 
					"Semáforo", "Slice (luz)", "Vista principal", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Whisp", "Vista principal", "Filigrana", "Austin",
					"Slice (luz)", "Faceta", "Semáforo", "Retrospectiva", "Cuadrícula",
					"Movimiento", "Cortar (oscuro)", "Línea lateral", "Ion (oscuro)",
					"Ion (claro)", "Integral", "Con bandas")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Anual", "Austero", "Austin", "Conservador",
					"Contraste", "Cuadrícula", "Cubículos", "Exposición", "Línea lateral",
					"Moderno", "Mosaicos", "Movimiento", "Papel periódico",
					"Perspectiva", "Pilas", "Puzzle", "Rayas", "Sobrepasar")
				}
			}

		'fi-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
					"Kuiskaus", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
					"Kiehkura", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aakkoset", "Askeettinen", "Austin", "Kontrasti",
					"Laatikot", "Liike", "Liituraita", "Mod", "Osittain peitossa",
					"Palapeli", "Perinteinen", "Perspektiivi", "Pinot", "Ruudukko",
					"Ruudut", "Sanomalehtipaperi", "Sivussa", "Vuotuinen", "Ylitys")
				}
			}

		'fr-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("À bandes", "Austin", "Facette", "Filigrane", 
					"Guide", "Intégrale", "Ion (clair)", "Ion (foncé)", 
					"Lignes latérales", "Quadrillage", "Rétrospective", "Secteur (clair)", 
					"Secteur (foncé)", "Sémaphore", "ViewMaster", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annuel", "Austère", "Austin", 
					"Blocs empilés", "Classique", "Contraste", "Emplacements de bureau", 
					"Exposition", "Guide", "Ligne latérale", "Moderne", 
					"Mosaïques", "Mots croisés", "Papier journal", "Perspective",
					"Quadrillage", "Rayures fines", "Transcendant")
				}
			}

		'nb-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "BevegElse", "Dempet", "Fasett", "Filigran",
					"Integral", "Ion (lys)", "Ion (mørk)", "Retrospekt", "Rutenett",
					"Sektor (lys)", "Sektor (mørk)", "Semafor", "Sidelinje", "Stripet",
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Årlig", "Avistrykk", "Austin", "Avlukker",
					"BevegElse", "Engasjement", "Enkel", "Fliser", "Konservativ",
					"Kontrast", "Mod", "Perspektiv", "Puslespill", "Rutenett", "Sidelinje",
					"Smale striper", "Stabler", "Transcenderende")
				}
			}

		'nl-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Beweging", "Facet", "Filigraan", "Gestreept",
					"Integraal", "Ion (donker)", "Ion (licht)", "Raster",
					"Segment (Light)", "Semafoor", "Slice (donker)", "Spriet",
					"Terugblik", "Terzijde", "ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aantrekkelijk", "Alfabet", "Austin", "Bescheiden",
					"Beweging", "Blikvanger", "Contrast", "Eenvoudig", "Jaarlijks",
					"Krantenpapier", "Krijtstreep", "Kubussen", "Mod", "Perspectief",
					"Puzzel", "Raster", "Stapels",
					"Tegels", "Terzijde")
				}
			}

		'pt-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Animação", "Austin", "Em Tiras", "Exibição Mestra",
					"Faceta", "Fatia (Clara)", "Fatia (Escura)", "Filete", "Filigrana", 
					"Grade", "Integral", "Íon (Claro)", "Íon (Escuro)", "Linha Lateral",
					"Retrospectiva", "Semáforo")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Animação", "Anual", "Austero", "Austin", "Baias",
					"Conservador", "Contraste", "Exposição", "Grade", "Ladrilhos",
					"Linha Lateral", "Listras", "Mod", "Papel Jornal", "Perspectiva", "Pilhas",
					"Quebra-cabeça", "Transcend")
				}
			}

		'sv-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Band", "Fasett", "Filigran", "Integrerad", "Jon (ljust)",
					"Jon (mörkt)", "Knippe", "Rutnät", "RörElse", "Sektor (ljus)", "Sektor (mörk)",
					"Semafor", "Sidlinje", "VisaHuvudsida", "Återblick")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabetmönster", "Austin", "Enkelt", "Exponering", "Konservativt",
					"Kontrast", "Kritstreck", "Kuber", "Perspektiv", "Plattor", "Pussel", "Rutnät",
					"RörElse", "Sidlinje", "Sobert", "Staplat", "Tidningspapper", "Årligt",
					"Övergående")
				}
			}

		'zh-'	{
				If($xWordVersion -eq $wdWord2010 -or $xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ('奥斯汀', '边线型', '花丝', '怀旧', '积分',
					'离子(浅色)', '离子(深色)', '母版型', '平面', '切片(浅色)',
					'切片(深色)', '丝状', '网格', '镶边', '信号灯',
					'运动型')
				}
			}

		Default	{
					If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
					{
						$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
						"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
						"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
						"Whisp")
					}
					ElseIf($xWordVersion -eq $wdWord2010)
					{
						$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
						"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
						"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
					}
				}
	}
	
	If($xArray -contains $xCP)
	{
		$xArray = $Null
		Return $True
	}
	Else
	{
		$xArray = $Null
		Return $False
	}
}

Function CheckWordPrereq
{
	If((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Word.Application) -eq $False)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Host "`n`n`t`tThis script directly outputs to Microsoft Word, please install Microsoft Word`n`n"
		Exit
	}

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId
	
	#Find out if winword is running in our session
	[bool]$wordrunning = ((Get-Process 'WinWord' -ea 0) | Where-Object {$_.SessionId -eq $SessionID}) -ne $Null
	If($wordrunning)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Host "`n`n`tPlease close all instances of Microsoft Word before running this report.`n`n"
		Exit
	}
}

Function ValidateCompanyName
{
	[bool]$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	If($xResult)
	{
		Return Get-LocalRegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	}
	Else
	{
		$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		If($xResult)
		{
			Return Get-LocalRegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		}
		Else
		{
			Return ""
		}
	}
}

Function Set-DocumentProperty {
    <#
	.SYNOPSIS
	Function to set the Title Page document properties in MS Word
	.DESCRIPTION
	Long description
	.PARAMETER Document
	Current Document Object
	.PARAMETER DocProperty
	Parameter description
	.PARAMETER Value
	Parameter description
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Title -Value 'MyTitle'
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Company -Value 'MyCompany'
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Author -Value 'Jim Moyle'
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Subject -Value 'MySubjectTitle'
	.NOTES
	Function Created by Jim Moyle June 2017
	Twitter : @JimMoyle
	#>
    param (
        [object]$Document,
        [String]$DocProperty,
        [string]$Value
    )
    try {
        $binding = "System.Reflection.BindingFlags" -as [type]
        $builtInProperties = $Document.BuiltInDocumentProperties
        $property = [System.__ComObject].invokemember("item", $binding::GetProperty, $null, $BuiltinProperties, $DocProperty)
        [System.__ComObject].invokemember("value", $binding::SetProperty, $null, $property, $Value)
    }
    catch {
        Write-Warning "Failed to set $DocProperty to $Value"
    }
}

Function FindWordDocumentEnd
{
	#return focus to main document    
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument
	#move to the end of the current document
	$Script:Selection.EndKey($wdStory,$wdMove) | Out-Null
}

Function SetupWord
{
	Write-Verbose "$(Get-Date): Setting up Word"
    
	# Setup word for output
	Write-Verbose "$(Get-Date): Create Word comObject."
	$Script:Word = New-Object -comobject "Word.Application" -EA 0 4>$Null
	
	If(!$? -or $Null -eq $Script:Word)
	{
		Write-Warning "The Word object could not be created.  You may need to repair your Word installation."
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tThe Word object could not be created.  You may need to repair your Word installation.`n`n`t`tScript cannot continue.`n`n"
		Exit
	}

	Write-Verbose "$(Get-Date): Determine Word language value"
	If( ( validStateProp $Script:Word Language Value__ ) )
	{
		[int]$Script:WordLanguageValue = [int]$Script:Word.Language.Value__
	}
	Else
	{
		[int]$Script:WordLanguageValue = [int]$Script:Word.Language
	}

	If(!($Script:WordLanguageValue -gt -1))
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tUnable to determine the Word language value.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}
	Write-Verbose "$(Get-Date): Word language value is $Script:WordLanguageValue"
	
	$Script:WordCultureCode = GetCulture $Script:WordLanguageValue
	
	SetWordHashTable $Script:WordCultureCode
	
	[int]$Script:WordVersion = [int]$Script:Word.Version
	If($Script:WordVersion -eq $wdWord2016)
	{
		$Script:WordProduct = "Word 2016"
	}
	ElseIf($Script:WordVersion -eq $wdWord2013)
	{
		$Script:WordProduct = "Word 2013"
	}
	ElseIf($Script:WordVersion -eq $wdWord2010)
	{
		$Script:WordProduct = "Word 2010"
	}
	ElseIf($Script:WordVersion -eq $wdWord2007)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tMicrosoft Word 2007 is no longer supported.`n`n`t`tScript will end.`n`n"
		AbortScript
	}
	Else
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tYou are running an untested or unsupported version of Microsoft Word.`n`n`t`tScript will end.`n`n`t`tPlease send info on your version of Word to webster@carlwebster.com`n`n"
		AbortScript
	}

	#only validate CompanyName if the field is blank
	If([String]::IsNullOrEmpty($Script:CoName))
	{
		Write-Verbose "$(Get-Date): Company name is blank.  Retrieve company name from registry."
		$TmpName = ValidateCompanyName
		
		If([String]::IsNullOrEmpty($TmpName))
		{
			Write-Warning "`n`n`t`tCompany Name is blank so Cover Page will not show a Company Name."
			Write-Warning "`n`t`tCheck HKCU:\Software\Microsoft\Office\Common\UserInfo for Company or CompanyName value."
			Write-Warning "`n`t`tYou may want to use the -CompanyName parameter if you need a Company Name on the cover page.`n`n"
		}
		Else
		{
			$Script:CoName = $TmpName
			Write-Verbose "$(Get-Date): Updated company name to $Script:CoName"
		}
	}

	If($Script:WordCultureCode -ne "en-")
	{
		Write-Verbose "$(Get-Date): Check Default Cover Page for $WordCultureCode"
		[bool]$CPChanged = $False
		Switch ($Script:WordCultureCode)
		{
			'ca-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Línia lateral"
						$CPChanged = $True
					}
				}

			'da-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidelinje"
						$CPChanged = $True
					}
				}

			'de-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Randlinie"
						$CPChanged = $True
					}
				}

			'es-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Línea lateral"
						$CPChanged = $True
					}
				}

			'fi-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sivussa"
						$CPChanged = $True
					}
				}

			'fr-'	{
					If($CoverPage -eq "Sideline")
					{
						If($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016)
						{
							$CoverPage = "Lignes latérales"
							$CPChanged = $True
						}
						Else
						{
							$CoverPage = "Ligne latérale"
							$CPChanged = $True
						}
					}
				}

			'nb-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidelinje"
						$CPChanged = $True
					}
				}

			'nl-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Terzijde"
						$CPChanged = $True
					}
				}

			'pt-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Linha Lateral"
						$CPChanged = $True
					}
				}

			'sv-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidlinje"
						$CPChanged = $True
					}
				}

			'zh-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "边线型"
						$CPChanged = $True
					}
				}
		}

		If($CPChanged)
		{
			Write-Verbose "$(Get-Date): Changed Default Cover Page from Sideline to $CoverPage"
		}
	}

	Write-Verbose "$(Get-Date): Validate cover page $($CoverPage) for culture code $Script:WordCultureCode"
	[bool]$ValidCP = $False
	
	$ValidCP = ValidateCoverPage $Script:WordVersion $CoverPage $Script:WordCultureCode
	
	If(!$ValidCP)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Verbose "$(Get-Date): Word language value $Script:WordLanguageValue"
		Write-Verbose "$(Get-Date): Culture code $Script:WordCultureCode"
		Write-Error "`n`n`t`tFor $Script:WordProduct, $CoverPage is not a valid Cover Page option.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}

	ShowScriptOptions

	$Script:Word.Visible = $False

	#http://jdhitsolutions.com/blog/2012/05/san-diego-2012-powershell-deep-dive-slides-and-demos/
	#using Jeff's Demo-WordReport.ps1 file for examples
	Write-Verbose "$(Get-Date): Load Word Templates"

	[bool]$Script:CoverPagesExist = $False
	[bool]$BuildingBlocksExist = $False

	$Script:Word.Templates.LoadBuildingBlocks()
	#word 2010/2013/2016
	$BuildingBlocksCollection = $Script:Word.Templates | Where-Object {$_.name -eq "Built-In Building Blocks.dotx"}

	Write-Verbose "$(Get-Date): Attempt to load cover page $CoverPage"
	$part = $Null

	$BuildingBlocksCollection | 
	ForEach-Object{
		If ($_.BuildingBlockEntries.Item($CoverPage).Name -eq $CoverPage) 
		{
			$BuildingBlocks = $_
		}
	}        

	If($Null -ne $BuildingBlocks)
	{
		$BuildingBlocksExist = $True

		Try 
		{
			$part = $BuildingBlocks.BuildingBlockEntries.Item($CoverPage)
		}

		Catch
		{
			$part = $Null
		}

		If($Null -ne $part)
		{
			$Script:CoverPagesExist = $True
		}
	}

	If(!$Script:CoverPagesExist)
	{
		Write-Verbose "$(Get-Date): Cover Pages are not installed or the Cover Page $CoverPage does not exist."
		Write-Warning "Cover Pages are not installed or the Cover Page $CoverPage does not exist."
		Write-Warning "This report will not have a Cover Page."
	}

	Write-Verbose "$(Get-Date): Create empty word doc"
	$Script:Doc = $Script:Word.Documents.Add()
	If($Null -eq $Script:Doc)
	{
		Write-Verbose "$(Get-Date): "
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tAn empty Word document could not be created.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}

	$Script:Selection = $Script:Word.Selection
	If($Null -eq $Script:Selection)
	{
		Write-Verbose "$(Get-Date): "
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tAn unknown error happened selecting the entire Word document for default formatting options.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}

	#set Default tab stops to 1/2 inch (this line is not from Jeff Hicks)
	#36 = .50"
	$Script:Word.ActiveDocument.DefaultTabStop = 36

	#Disable Spell and Grammar Check to resolve issue and improve performance (from Pat Coughlin)
	Write-Verbose "$(Get-Date): Disable grammar and spell checking"
	#bug reported 1-Apr-2014 by Tim Mangan
	#save current options first before turning them off
	$Script:CurrentGrammarOption = $Script:Word.Options.CheckGrammarAsYouType
	$Script:CurrentSpellingOption = $Script:Word.Options.CheckSpellingAsYouType
	$Script:Word.Options.CheckGrammarAsYouType = $False
	$Script:Word.Options.CheckSpellingAsYouType = $False

	If($BuildingBlocksExist)
	{
		#insert new page, getting ready for table of contents
		Write-Verbose "$(Get-Date): Insert new page, getting ready for table of contents"
		$part.Insert($Script:Selection.Range,$True) | Out-Null
		$Script:Selection.InsertNewPage()

		#table of contents
		Write-Verbose "$(Get-Date): Table of Contents - $Script:MyHash.Word_TableOfContents"
		$toc = $BuildingBlocks.BuildingBlockEntries.Item($Script:MyHash.Word_TableOfContents)
		If($Null -eq $toc)
		{
			Write-Verbose "$(Get-Date): "
			Write-Verbose "$(Get-Date): Table of Content - $Script:MyHash.Word_TableOfContents could not be retrieved."
			Write-Warning "This report will not have a Table of Contents."
		}
		Else
		{
			$toc.insert($Script:Selection.Range,$True) | Out-Null
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): Table of Contents are not installed."
		Write-Warning "Table of Contents are not installed so this report will not have a Table of Contents."
	}

	#set the footer
	Write-Verbose "$(Get-Date): Set the footer"
	[string]$footertext = "Report created by $username"

	#get the footer
	Write-Verbose "$(Get-Date): Get the footer and format font"
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekPrimaryFooter
	#get the footer and format font
	$footers = $Script:Doc.Sections.Last.Footers
	ForEach ($footer in $footers) 
	{
		If($footer.exists) 
		{
			$footer.range.Font.name = "Calibri"
			$footer.range.Font.size = 8
			$footer.range.Font.Italic = $True
			$footer.range.Font.Bold = $True
		}
	} #end ForEach
	Write-Verbose "$(Get-Date): Footer text"
	$Script:Selection.HeaderFooter.Range.Text = $footerText

	#add page numbering
	Write-Verbose "$(Get-Date): Add page numbering"
	$Script:Selection.HeaderFooter.PageNumbers.Add($wdAlignPageNumberRight) | Out-Null

	FindWordDocumentEnd
	Write-Verbose "$(Get-Date):"
	#end of Jeff Hicks 
}

Function UpdateDocumentProperties
{
	Param([string]$AbstractTitle, [string]$SubjectTitle)
	#updated 12-Nov-2017 with additional cover page fields
	#Update document properties
	If($MSWORD -or $PDF)
	{
		If($Script:CoverPagesExist)
		{
			Write-Verbose "$(Get-Date): Set Cover Page Properties"
			#8-Jun-2017 put these 4 items in alpha order
            Set-DocumentProperty -Document $Script:Doc -DocProperty Author -Value $UserName
            Set-DocumentProperty -Document $Script:Doc -DocProperty Company -Value $Script:CoName
            Set-DocumentProperty -Document $Script:Doc -DocProperty Subject -Value $SubjectTitle
            Set-DocumentProperty -Document $Script:Doc -DocProperty Title -Value $Script:title

			#Get the Coverpage XML part
			$cp = $Script:Doc.CustomXMLParts | Where-Object {$_.NamespaceURI -match "coverPageProps$"}

			#get the abstract XML part
			$ab = $cp.documentelement.ChildNodes | Where-Object {$_.basename -eq "Abstract"}
			#set the text
			If([String]::IsNullOrEmpty($Script:CoName))
			{
				[string]$abstract = $AbstractTitle
			}
			Else
			{
				[string]$abstract = "$($AbstractTitle) for $($Script:CoName)"
			}
			$ab.Text = $abstract

			#added 8-Jun-2017
			$ab = $cp.documentelement.ChildNodes | Where-Object {$_.basename -eq "CompanyAddress"}
			#set the text
			[string]$abstract = $CompanyAddress
			$ab.Text = $abstract

			#added 8-Jun-2017
			$ab = $cp.documentelement.ChildNodes | Where-Object {$_.basename -eq "CompanyEmail"}
			#set the text
			[string]$abstract = $CompanyEmail
			$ab.Text = $abstract

			#added 8-Jun-2017
			$ab = $cp.documentelement.ChildNodes | Where-Object {$_.basename -eq "CompanyFax"}
			#set the text
			[string]$abstract = $CompanyFax
			$ab.Text = $abstract

			#added 8-Jun-2017
			$ab = $cp.documentelement.ChildNodes | Where-Object {$_.basename -eq "CompanyPhone"}
			#set the text
			[string]$abstract = $CompanyPhone
			$ab.Text = $abstract

			$ab = $cp.documentelement.ChildNodes | Where-Object {$_.basename -eq "PublishDate"}
			#set the text
			[string]$abstract = (Get-Date -Format d).ToString()
			$ab.Text = $abstract

			Write-Verbose "$(Get-Date): Update the Table of Contents"
			#update the Table of Contents
			$Script:Doc.TablesOfContents.item(1).Update()
			$cp = $Null
			$ab = $Null
			$abstract = $Null
		}
	}
}
#endregion

#region registry functions
#http://stackoverflow.com/questions/5648931/test-if-registry-value-exists
# This Function just gets $True or $False
Function Test-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	$key -and $Null -ne $key.GetValue($name, $Null)
}

# Gets the specified local registry value or $Null if it is missing
Function Get-LocalRegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	If($key)
	{
		$key.GetValue($name, $Null)
	}
	Else
	{
		$Null
	}
}

Function Get-RegistryValue
{
	# Gets the specified registry value or $Null if it is missing
	[CmdletBinding()]
	Param([string]$path, [string]$name, [string]$ComputerName)
	If($ComputerName -eq $env:computername -or $ComputerName -eq "LocalHost")
	{
		$key = Get-Item -LiteralPath $path -EA 0
		If($key)
		{
			Return $key.GetValue($name, $Null)
		}
		Else
		{
			Return $Null
		}
	}
	Else
	{
		#path needed here is different for remote registry access
		$path1 = $path.SubString(6)
		$path2 = $path1.Replace('\','\\')
		$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ComputerName)
		$RegKey= $Reg.OpenSubKey($path2)
		$Results = $RegKey.GetValue($name)
		If($Null -ne $Results)
		{
			Return $Results
		}
		Else
		{
			Return $Null
		}
	}
}
#endregion

#region word, text and html line output functions
Function line
#function created by Michael B. Smith, Exchange MVP
#@essentialexchange on Twitter
#http://TheEssentialExchange.com
#for creating the formatted text report
#created March 2011
#updated March 2014
{
	Param( [int]$tabs = 0, [string]$name = '', [string]$value = '', [string]$newline = "`r`n", [switch]$nonewline )
	While( $tabs -gt 0 ) { $Global:Output += "`t"; $tabs--; }
	If( $nonewline )
	{
		$Global:Output += $name + $value
	}
	Else
	{
		$Global:Output += $name + $value + $newline
	}
}
	
Function WriteWordLine
#Function created by Ryan Revord
#@rsrevord on Twitter
#Function created to make output to Word easy in this script
#updated 27-Mar-2014 to include font name, font size, italics and bold options
{
	Param([int]$style=0, 
	[int]$tabs = 0, 
	[string]$name = '', 
	[string]$value = '', 
	[string]$fontName=$Null,
	[int]$fontSize=0,
	[bool]$italics=$False,
	[bool]$boldface=$False,
	[Switch]$nonewline)
	
	#Build output style
	[string]$output = ""
	Switch ($style)
	{
		0 {$Script:Selection.Style = $Script:MyHash.Word_NoSpacing; Break}
		1 {$Script:Selection.Style = $Script:MyHash.Word_Heading1; Break}
		2 {$Script:Selection.Style = $Script:MyHash.Word_Heading2; Break}
		3 {$Script:Selection.Style = $Script:MyHash.Word_Heading3; Break}
		4 {$Script:Selection.Style = $Script:MyHash.Word_Heading4; Break}
		Default {$Script:Selection.Style = $Script:MyHash.Word_NoSpacing; Break}
	}
	
	#build # of tabs
	While($tabs -gt 0)
	{ 
		$output += "`t"; $tabs--; 
	}
 
	If(![String]::IsNullOrEmpty($fontName)) 
	{
		$Script:Selection.Font.name = $fontName
	} 

	If($fontSize -ne 0) 
	{
		$Script:Selection.Font.size = $fontSize
	} 
 
	If($italics -eq $True) 
	{
		$Script:Selection.Font.Italic = $True
	} 
 
	If($boldface -eq $True) 
	{
		$Script:Selection.Font.Bold = $True
	} 

	#output the rest of the parameters.
	$output += $name + $value
	$Script:Selection.TypeText($output)
 
	#test for new WriteWordLine 0.
	If($nonewline)
	{
		# Do nothing.
	} 
	Else 
	{
		$Script:Selection.TypeParagraph()
	}
}

#***********************************************************************************************************
# WriteHTMLLine
#***********************************************************************************************************

<#
.Synopsis
	Writes a line of output for HTML output
.DESCRIPTION
	This function formats an HTML line
.USAGE
	WriteHTMLLine <Style> <Tabs> <Name> <Value> <Font Name> <Font Size> <Options>

	0 for Font Size denotes using the default font size of 2 or 10 point

.EXAMPLE
	WriteHTMLLine 0 0 " "

	Writes a blank line with no style or tab stops, obviously none needed.

.EXAMPLE
	WriteHTMLLine 0 1 "This is a regular line of text indented 1 tab stops"

	Writes a line with 1 tab stop.

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in italics" "" $Null 0 $htmlitalics

	Writes a line omitting font and font size and setting the italics attribute

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in bold" "" $Null 0 $htmlbold

	Writes a line omitting font and font size and setting the bold attribute

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in bold italics" "" $Null 0 ($htmlbold -bor $htmlitalics)

	Writes a line omitting font and font size and setting both italics and bold options

.EXAMPLE	
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in 10 point" "" $Null 2  # 10 point font

	Writes a line using 10 point font

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in Courier New font" "" "Courier New" 0 

	Writes a line using Courier New Font and 0 font point size (default = 2 if set to 0)

.EXAMPLE	
	WriteHTMLLine 0 0 "This is a regular line of RED text indented 0 tab stops with the computer name as data in 10 point Courier New bold italics: " $env:computername "Courier New" 2 ($htmlbold -bor $htmlred -bor $htmlitalics)

	Writes a line using Courier New Font with first and second string values to be used, also uses 10 point font with bold, italics and red color options set.

.NOTES

	Font Size - Unlike word, there is a limited set of font sizes that can be used in HTML.  They are:
		0 - default which actually gives it a 2 or 10 point.
		1 - 7.5 point font size
		2 - 10 point
		3 - 13.5 point
		4 - 15 point
		5 - 18 point
		6 - 24 point
		7 - 36 point
	Any number larger than 7 defaults to 7

	Style - Refers to the headers that are used with output and resemble the headers in word, 
	HTML supports headers h1-h6 and h1-h4 are more commonly used.  Unlike word, H1 will not 
	give you a blue colored font, you will have to set that yourself.

	Colors and Bold/Italics Flags are:

		htmlbold       
		htmlitalics    
		htmlred        
		htmlcyan        
		htmlblue       
		htmldarkblue   
		htmllightblue   
		htmlpurple      
		htmlyellow      
		htmllime       
		htmlmagenta     
		htmlwhite       
		htmlsilver      
		htmlgray       
		htmlolive       
		htmlorange      
		htmlmaroon      
		htmlgreen       
		htmlblack       
#>

Function WriteHTMLLine
#Function created by Ken Avram
#Function created to make output to HTML easy in this script
#headings fixed 12-Oct-2016 by Webster
#errors with $HTMLStyle fixed 7-Dec-2017 by Webster
{
	Param([int]$style=0, 
	[int]$tabs = 0, 
	[string]$name = '', 
	[string]$value = '', 
	[string]$fontName="Calibri",
	[int]$fontSize=1,
	[int]$options=$htmlblack)


	#Build output style
	[string]$output = ""

	If([String]::IsNullOrEmpty($Name))	
	{
		$HTMLBody = "<p></p>"
	}
	Else
	{
		$color = CheckHTMLColor $options

		#build # of tabs

		While($tabs -gt 0)
		{ 
			$output += "&nbsp;&nbsp;&nbsp;&nbsp;"; $tabs--; 
		}

		$HTMLFontName = $fontName		

		$HTMLBody = ""

		If($options -band $htmlitalics) 
		{
			$HTMLBody += "<i>"
		} 

		If($options -band $htmlbold) 
		{
			$HTMLBody += "<b>"
		} 

		#output the rest of the parameters.
		$output += $name + $value

		Switch ($style)
		{
			1 {$HTMLStyle = "<h1>"; Break}
			2 {$HTMLStyle = "<h2>"; Break}
			3 {$HTMLStyle = "<h3>"; Break}
			4 {$HTMLStyle = "<h4>"; Break}
			Default {$HTMLStyle = ""; Break}
		}

		$HTMLBody += $HTMLStyle + $output

		Switch ($style)
		{
			1 {$HTMLStyle = "</h1>"; Break}
			2 {$HTMLStyle = "</h2>"; Break}
			3 {$HTMLStyle = "</h3>"; Break}
			4 {$HTMLStyle = "</h4>"; Break}
			Default {$HTMLStyle = ""; Break}
		}

		#added by webster 12-oct-2016
		#if a heading, don't add the <br>
		#moved to after the two switch statements on 7-Dec-2017 to fix $HTMLStyle has not been set error
		If($HTMLStyle -eq "")
		{
			$HTMLBody += "<br><font face='" + $HTMLFontName + "' " + "color='" + $color + "' size='"  + $fontsize + "'>"
		}
		Else
		{
			$HTMLBody += "<font face='" + $HTMLFontName + "' " + "color='" + $color + "' size='"  + $fontsize + "'>"
		}
		
		$HTMLBody += $HTMLStyle +  "</font>"

		If($options -band $htmlitalics) 
		{
			$HTMLBody += "</i>"
		} 

		If($options -band $htmlbold) 
		{
			$HTMLBody += "</b>"
		} 

		#added by webster 12-oct-2016
		#if a heading, don't add the <br />
		#moved to inside the Else statement on 7-Dec-2017 to fix $HTMLStyle has not been set error
		If($HTMLStyle -eq "")
		{
			$HTMLBody += "<br />"
		}
	}

	out-file -FilePath $Script:FileName1 -Append -InputObject $HTMLBody 4>$Null
}
#endregion

#region HTML table functions
#***********************************************************************************************************
# AddHTMLTable - Called from FormatHTMLTable function
# Created by Ken Avram
# modified by Jake Rutski
#***********************************************************************************************************
Function AddHTMLTable
{
	Param([string]$fontName="Calibri",
	[int]$fontSize=2,
	[int]$colCount=0,
	[int]$rowCount=0,
	[object[]]$rowInfo=@(),
	[object[]]$fixedInfo=@())

	For($rowidx = $RowIndex;$rowidx -le $rowCount;$rowidx++)
	{
		$rd = @($rowInfo[$rowidx - 2])
		$htmlbody = $htmlbody + "<tr>"
		For($columnIndex = 0; $columnIndex -lt $colCount; $columnindex+=2)
		{
			$tmp = CheckHTMLColor $rd[$columnIndex+1]

			If($fixedInfo.Length -eq 0)
			{
				$htmlbody += "<td style=""background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}
			Else
			{
				$htmlbody += "<td style=""width:$($fixedInfo[$columnIndex/2]); background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}

			If($rd[$columnIndex+1] -band $htmlbold)
			{
				$htmlbody += "<b>"
			}
			If($rd[$columnIndex+1] -band $htmlitalics)
			{
				$htmlbody += "<i>"
			}
			If($Null -ne $rd[$columnIndex])
			{
				$cell = $rd[$columnIndex].tostring()
				If($cell -eq " " -or $cell.length -eq 0)
				{
					$htmlbody += "&nbsp;&nbsp;&nbsp;"
				}
				Else
				{
					For($i=0;$i -lt $cell.length;$i++)
					{
						If($cell[$i] -eq " ")
						{
							$htmlbody += "&nbsp;"
						}
						If($cell[$i] -ne " ")
						{
							Break
						}
					}
					$htmlbody += $cell
				}
			}
			Else
			{
				$htmlbody += "&nbsp;&nbsp;&nbsp;"
			}
			If($rd[$columnIndex+1] -band $htmlbold)
			{
				$htmlbody += "</b>"
			}
			If($rd[$columnIndex+1] -band $htmlitalics)
			{
				$htmlbody += "</i>"
			}
			$htmlbody += "</font></td>"
		}
		$htmlbody += "</tr>"
	}
	out-file -FilePath $Script:FileName1 -Append -InputObject $HTMLBody 4>$Null 
}

#***********************************************************************************************************
# FormatHTMLTable 
# Created by Ken Avram
# modified by Jake Rutski
#***********************************************************************************************************

<#
.Synopsis
	Format table for HTML output document
.DESCRIPTION
	This function formats a table for HTML from an array of strings
.PARAMETER noBorder
	If set to $true, a table will be generated without a border (border='0')
.PARAMETER noHeadCols
	This parameter should be used when generating tables without column headers
	Set this parameter equal to the number of columns in the table
.PARAMETER rowArray
	This parameter contains the row data array for the table
.PARAMETER columnArray
	This parameter contains column header data for the table
.PARAMETER fixedWidth
	This parameter contains widths for columns in pixel format ("100px") to override auto column widths
	The variable should contain a width for each column you wish to override the auto-size setting
	For example: $columnWidths = @("100px","110px","120px","130px","140px")

.USAGE
	FormatHTMLTable <Table Header> <Table Format> <Font Name> <Font Size>

.EXAMPLE
	FormatHTMLTable "Table Heading" "auto" "Calibri" 3

	This example formats a table and writes it out into an html file.  All of the parameters are optional
	defaults are used if not supplied.

	for <Table format>, the default is auto which will autofit the text into the columns and adjust to the longest text in that column.  You can also use percentage i.e. 25%
	which will take only 25% of the line and will auto word wrap the text to the next line in the column.  Also, instead of using a percentage, you can use pixels i.e. 400px.

	FormatHTMLTable "Table Heading" "auto" -rowArray $rowData -columnArray $columnData

	This example creates an HTML table with a heading of 'Table Heading', auto column spacing, column header data from $columnData and row data from $rowData

	FormatHTMLTable "Table Heading" -rowArray $rowData -noHeadCols 3

	This example creates an HTML table with a heading of 'Table Heading', auto column spacing, no header, and row data from $rowData

	FormatHTMLTable "Table Heading" -rowArray $rowData -fixedWidth $fixedColumns

	This example creates an HTML table with a heading of 'Table Heading, no header, row data from $rowData, and fixed columns defined by $fixedColumns

.NOTES
	In order to use the formatted table it first has to be loaded with data.  Examples below will show how to load the table:

	First, initialize the table array

	$rowdata = @()

	Then Load the array.  If you are using column headers then load those into the column headers array, otherwise the first line of the table goes into the column headers array
	and the second and subsequent lines go into the $rowdata table as shown below:

	$columnHeaders = @('Display Name',($htmlsilver -bor $htmlbold),'Status',($htmlsilver -bor $htmlbold),'Startup Type',($htmlsilver -bor $htmlbold))

	The first column is the actual name to display, the second are the attributes of the column i.e. color anded with bold or italics.  For the anding, parens are required or it will
	not format correctly.

	This is following by adding rowdata as shown below.  As more columns are added the columns will auto adjust to fit the size of the page.

	$rowdata = @()
	$columnHeaders = @("User Name",($htmlsilver -bor $htmlbold),$UserName,$htmlwhite)
	$rowdata += @(,('Save as PDF',($htmlsilver -bor $htmlbold),$PDF.ToString(),$htmlwhite))
	$rowdata += @(,('Save as TEXT',($htmlsilver -bor $htmlbold),$TEXT.ToString(),$htmlwhite))
	$rowdata += @(,('Save as WORD',($htmlsilver -bor $htmlbold),$MSWORD.ToString(),$htmlwhite))
	$rowdata += @(,('Save as HTML',($htmlsilver -bor $htmlbold),$HTML.ToString(),$htmlwhite))
	$rowdata += @(,('Add DateTime',($htmlsilver -bor $htmlbold),$AddDateTime.ToString(),$htmlwhite))
	$rowdata += @(,('Hardware Inventory',($htmlsilver -bor $htmlbold),$Hardware.ToString(),$htmlwhite))
	$rowdata += @(,('Computer Name',($htmlsilver -bor $htmlbold),$ComputerName,$htmlwhite))
	$rowdata += @(,('Filename1',($htmlsilver -bor $htmlbold),$Script:FileName1,$htmlwhite))
	$rowdata += @(,('OS Detected',($htmlsilver -bor $htmlbold),$Script:RunningOS,$htmlwhite))
	$rowdata += @(,('PSUICulture',($htmlsilver -bor $htmlbold),$PSCulture,$htmlwhite))
	$rowdata += @(,('PoSH version',($htmlsilver -bor $htmlbold),$Host.Version.ToString(),$htmlwhite))
	FormatHTMLTable "Example of Horizontal AutoFitContents HTML Table" -rowArray $rowdata

	The 'rowArray' paramater is mandatory to build the table, but it is not set as such in the function - if nothing is passed, the table will be empty.

	Colors and Bold/Italics Flags are shown below:

		htmlbold       
		htmlitalics    
		htmlred        
		htmlcyan        
		htmlblue       
		htmldarkblue   
		htmllightblue   
		htmlpurple      
		htmlyellow      
		htmllime       
		htmlmagenta     
		htmlwhite       
		htmlsilver      
		htmlgray       
		htmlolive       
		htmlorange      
		htmlmaroon      
		htmlgreen       
		htmlblack     

#>

Function FormatHTMLTable
{
	Param([string]$tableheader,
	[string]$tablewidth="auto",
	[string]$fontName="Calibri",
	[int]$fontSize=2,
	[switch]$noBorder=$false,
	[int]$noHeadCols=1,
	[object[]]$rowArray=@(),
	[object[]]$fixedWidth=@(),
	[object[]]$columnArray=@())

	$HTMLBody = "<b><font face='" + $fontname + "' size='" + ($fontsize + 1) + "'>" + $tableheader + "</font></b>"

	If($columnArray.Length -eq 0)
	{
		$NumCols = $noHeadCols + 1
	}  # means we have no column headers, just a table
	Else
	{
		$NumCols = $columnArray.Length
	}  # need to add one for the color attrib

	If($Null -ne $rowArray)
	{
		$NumRows = $rowArray.length + 1
	}
	Else
	{
		$NumRows = 1
	}

	If($noBorder)
	{
		$htmlbody += "<table border='0' width='" + $tablewidth + "'>"
	}
	Else
	{
		$htmlbody += "<table border='1' width='" + $tablewidth + "'>"
	}

	If(!($columnArray.Length -eq 0))
	{
		$htmlbody += "<tr>"

		For($columnIndex = 0; $columnIndex -lt $NumCols; $columnindex+=2)
		{
			$tmp = CheckHTMLColor $columnArray[$columnIndex+1]
			If($fixedWidth.Length -eq 0)
			{
				$htmlbody += "<td style=""background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}
			Else
			{
				$htmlbody += "<td style=""width:$($fixedWidth[$columnIndex/2]); background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}

			If($columnArray[$columnIndex+1] -band $htmlbold)
			{
				$htmlbody += "<b>"
			}
			If($columnArray[$columnIndex+1] -band $htmlitalics)
			{
				$htmlbody += "<i>"
			}
			If($Null -ne $columnArray[$columnIndex])
			{
				If($columnArray[$columnIndex] -eq " " -or $columnArray[$columnIndex].length -eq 0)
				{
					$htmlbody += "&nbsp;&nbsp;&nbsp;"
				}
				Else
				{
					For($i=0;$i -lt $columnArray[$columnIndex].length;$i+=2)
					{
						If($columnArray[$columnIndex][$i] -eq " ")
						{
							$htmlbody += "&nbsp;"
						}
						If($columnArray[$columnIndex][$i] -ne " ")
						{
							Break
						}
					}
					$htmlbody += $columnArray[$columnIndex]
				}
			}
			Else
			{
				$htmlbody += "&nbsp;&nbsp;&nbsp;"
			}
			If($columnArray[$columnIndex+1] -band $htmlbold)
			{
				$htmlbody += "</b>"
			}
			If($columnArray[$columnIndex+1] -band $htmlitalics)
			{
				$htmlbody += "</i>"
			}
			$htmlbody += "</font></td>"
		}
		$htmlbody += "</tr>"
	}
	$rowindex = 2
	If($Null -ne $rowArray)
	{
		AddHTMLTable $fontName $fontSize -colCount $numCols -rowCount $NumRows -rowInfo $rowArray -fixedInfo $fixedWidth
		$rowArray = @()
		$htmlbody = "</table>"
	}
	Else
	{
		$HTMLBody += "</table>"
	}	
	out-file -FilePath $Script:FileName1 -Append -InputObject $HTMLBody 4>$Null 
}
#endregion

#region other HTML functions
#***********************************************************************************************************
# CheckHTMLColor - Called from AddHTMLTable WriteHTMLLine and FormatHTMLTable
#***********************************************************************************************************
Function CheckHTMLColor
{
	Param($hash)

	If($hash -band $htmlwhite)
	{
		Return $htmlwhitemask
	}
	If($hash -band $htmlred)
	{
		Return $htmlredmask
	}
	If($hash -band $htmlcyan)
	{
		Return $htmlcyanmask
	}
	If($hash -band $htmlblue)
	{
		Return $htmlbluemask
	}
	If($hash -band $htmldarkblue)
	{
		Return $htmldarkbluemask
	}
	If($hash -band $htmllightblue)
	{
		Return $htmllightbluemask
	}
	If($hash -band $htmlpurple)
	{
		Return $htmlpurplemask
	}
	If($hash -band $htmlyellow)
	{
		Return $htmlyellowmask
	}
	If($hash -band $htmllime)
	{
		Return $htmllimemask
	}
	If($hash -band $htmlmagenta)
	{
		Return $htmlmagentamask
	}
	If($hash -band $htmlsilver)
	{
		Return $htmlsilvermask
	}
	If($hash -band $htmlgray)
	{
		Return $htmlgraymask
	}
	If($hash -band $htmlblack)
	{
		Return $htmlblackmask
	}
	If($hash -band $htmlorange)
	{
		Return $htmlorangemask
	}
	If($hash -band $htmlmaroon)
	{
		Return $htmlmaroonmask
	}
	If($hash -band $htmlgreen)
	{
		Return $htmlgreenmask
	}
	If($hash -band $htmlolive)
	{
		Return $htmlolivemask
	}
}

Function SetupHTML
{
	Write-Verbose "$(Get-Date): Setting up HTML"
	If(!$AddDateTime)
	{
		[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName).html"
	}
	ElseIf($AddDateTime)
	{
		[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName)_$(Get-Date -f yyyy-MM-dd_HHmm).html"
	}

	$htmlhead = "<html><head><meta http-equiv='Content-Language' content='da'><title>" + $Script:Title + "</title></head><body>"
	out-file -FilePath $Script:Filename1 -Force -InputObject $HTMLHead 4>$Null
}
#endregion

#region Iain's Word table functions

<#
.Synopsis
	Add a table to a Microsoft Word document
.DESCRIPTION
	This function adds a table to a Microsoft Word document from either an array of
	Hashtables or an array of PSCustomObjects.

	Using this function is quicker than setting each table cell individually but can
	only utilise the built-in MS Word table autoformats. Individual tables cells can
	be altered after the table has been appended to the document (a table reference
	is returned).
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. Column headers will display the key names as defined.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -List

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. No column headers will be added, in a ListView format.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray

	This example adds table to the MS Word document, utilising all note property names
	the array of PSCustomObjects. Column headers will display the note property names.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -Columns FirstName,LastName,EmailAddress

	This example adds a table to the MS Word document, but only using the specified
	key names: FirstName, LastName and EmailAddress. If other keys are present in the
	array of Hashtables they will be ignored.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray -Columns FirstName,LastName,EmailAddress -Headers "First Name","Last Name","Email Address"

	This example adds a table to the MS Word document, but only using the specified
	PSCustomObject note properties: FirstName, LastName and EmailAddress. If other note
	properties are present in the array of PSCustomObjects they will be ignored. The
	display names for each specified column header has been overridden to display a
	custom header. Note: the order of the header names must match the specified columns.
#>

Function AddWordTable
{
	[CmdletBinding()]
	Param
	(
		# Array of Hashtable (including table headers)
		[Parameter(Mandatory=$True, ValueFromPipelineByPropertyName=$True, ParameterSetName='Hashtable', Position=0)]
		[ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Hashtable,
		# Array of PSCustomObjects
		[Parameter(Mandatory=$True, ValueFromPipelineByPropertyName=$True, ParameterSetName='CustomObject', Position=0)]
		[ValidateNotNullOrEmpty()] [PSCustomObject[]] $CustomObject,
		# Array of Hashtable key names or PSCustomObject property names to include, in display order.
		# If not supplied then all Hashtable keys or all PSCustomObject properties will be displayed.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [string[]] $Columns = $Null,
		# Array of custom table header strings in display order.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [string[]] $Headers = $Null,
		# AutoFit table behavior.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [int] $AutoFit = -1,
		# List view (no headers)
		[Switch] $List,
		# Grid lines
		[Switch] $NoGridLines,
		[Switch] $NoInternalGridLines,
		# Built-in Word table formatting style constant
		# Would recommend only $wdTableFormatContempory for normal usage (possibly $wdTableFormatList5 for List view)
		[Parameter(ValueFromPipelineByPropertyName=$True)] [int] $Format = 0
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'" -f $PSCmdlet.ParameterSetName);
		## Check if -Columns wasn't specified but -Headers were (saves some additional parameter sets!)
		If(($Null -eq $Columns) -and ($Null -ne $Headers)) 
		{
			Write-Warning "No columns specified and therefore, specified headers will be ignored.";
			$Columns = $Null;
		}
		ElseIf(($Null -ne $Columns) -and ($Null -ne $Headers)) 
		{
			## Check if number of specified -Columns matches number of specified -Headers
			If($Columns.Length -ne $Headers.Length) 
			{
				Write-Error "The specified number of columns does not match the specified number of headers.";
			}
		} ## end elseif
	} ## end Begin

	Process
	{
		## Build the Word table data string to be converted to a range and then a table later.
		[System.Text.StringBuilder] $WordRangeString = New-Object System.Text.StringBuilder;

		Switch ($PSCmdlet.ParameterSetName) 
		{
			'CustomObject' 
			{
				If($Null -eq $Columns) 
				{
					## Build the available columns from all available PSCustomObject note properties
					[string[]] $Columns = @();
					## Add each NoteProperty name to the array
					ForEach($Property in ($CustomObject | Get-Member -MemberType NoteProperty)) 
					{ 
						$Columns += $Property.Name; 
					}
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Null -ne $Headers) 
					{
                        [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{ 
                        [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}

				## Iterate through each PSCustomObject
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Object in $CustomObject) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Object.$Column; 
					}
					## Use the ordered list to add each column in specified order
					[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end foreach
				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f ($CustomObject.Count));
			} ## end CustomObject

			Default 
			{   ## Hashtable
				If($Null -eq $Columns) 
				{
					## Build the available columns from all available hashtable keys. Hopefully
					## all Hashtables have the same keys (they should for a table).
					$Columns = $Hashtable[0].Keys;
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Null -ne $Headers) 
					{ 
						[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{
						[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}
                
				## Iterate through each Hashtable
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Hash in $Hashtable) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Hash.$Column; 
					}
					## Use the ordered list to add each column in specified order
					[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end foreach

				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f $Hashtable.Count);
			} ## end default
		} ## end switch

		## Create a MS Word range and set its text to our tab-delimited, concatenated string
		Write-Debug ("$(Get-Date): `t`tBuilding table range");
		$WordRange = $Script:Doc.Application.Selection.Range;
		$WordRange.Text = $WordRangeString.ToString();

		## Create hash table of named arguments to pass to the ConvertToTable method
		$ConvertToTableArguments = @{ Separator = [Microsoft.Office.Interop.Word.WdTableFieldSeparator]::wdSeparateByTabs; }

		## Negative built-in styles are not supported by the ConvertToTable method
		If($Format -ge 0) 
		{
			$ConvertToTableArguments.Add("Format", $Format);
			$ConvertToTableArguments.Add("ApplyBorders", $True);
			$ConvertToTableArguments.Add("ApplyShading", $True);
			$ConvertToTableArguments.Add("ApplyFont", $True);
			$ConvertToTableArguments.Add("ApplyColor", $True);
			If(!$List) 
			{ 
				$ConvertToTableArguments.Add("ApplyHeadingRows", $True); 
			}
			$ConvertToTableArguments.Add("ApplyLastRow", $True);
			$ConvertToTableArguments.Add("ApplyFirstColumn", $True);
			$ConvertToTableArguments.Add("ApplyLastColumn", $True);
		}

		## Invoke ConvertToTable method - with named arguments - to convert Word range to a table
		## See http://msdn.microsoft.com/en-us/library/office/aa171893(v=office.11).aspx
		Write-Debug ("$(Get-Date): `t`tConverting range to table");
		## Store the table reference just in case we need to set alternate row coloring
		$WordTable = $WordRange.GetType().InvokeMember(
			"ConvertToTable",                               # Method name
			[System.Reflection.BindingFlags]::InvokeMethod, # Flags
			$Null,                                          # Binder
			$WordRange,                                     # Target (self!)
			([Object[]]($ConvertToTableArguments.Values)),  ## Named argument values
			$Null,                                          # Modifiers
			$Null,                                          # Culture
			([String[]]($ConvertToTableArguments.Keys))     ## Named argument names
		);

		## Implement grid lines (will wipe out any existing formatting
		If($Format -lt 0) 
		{
			Write-Debug ("$(Get-Date): `t`tSetting table format");
			$WordTable.Style = $Format;
		}

		## Set the table autofit behavior
		If($AutoFit -ne -1) 
		{ 
			$WordTable.AutoFitBehavior($AutoFit); 
		}

		If(!$List)
		{
			#the next line causes the heading row to flow across page breaks
			$WordTable.Rows.First.Headingformat = $wdHeadingFormatTrue;
		}

		If(!$NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleSingle;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}
		If($NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleNone;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleNone;
		}
		If($NoInternalGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleNone;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}

		Return $WordTable;

	} ## end Process
}

<#
.Synopsis
	Sets the format of one or more Word table cells
.DESCRIPTION
	This function sets the format of one or more table cells, either from a collection
	of Word COM object cell references, an individual Word COM object cell reference or
	a hashtable containing Row and Column information.

	The font name, font size, bold, italic , underline and shading values can be used.
.EXAMPLE
	SetWordCellFormat -Hashtable $Coordinates -Table $TableReference -Bold

	This example sets all text to bold that is contained within the $TableReference
	Word table, using an array of hashtables. Each hashtable contain a pair of co-
	ordinates that is used to select the required cells. Note: the hashtable must
	contain the .Row and .Column key names. For example:
	@ { Row = 7; Column = 3 } to set the cell at row 7 and column 3 to bold.
.EXAMPLE
	$RowCollection = $Table.Rows.First.Cells
	SetWordCellFormat -Collection $RowCollection -Bold -Size 10

	This example sets all text to size 8 and bold for all cells that are contained
	within the first row of the table.
	Note: the $Table.Rows.First.Cells returns a collection of Word COM cells objects
	that are in the first table row.
.EXAMPLE
	$ColumnCollection = $Table.Columns.Item(2).Cells
	SetWordCellFormat -Collection $ColumnCollection -BackgroundColor 255

	This example sets the background (shading) of all cells in the table's second
	column to red.
	Note: the $Table.Columns.Item(2).Cells returns a collection of Word COM cells objects
	that are in the table's second column.
.EXAMPLE
	SetWordCellFormat -Cell $Table.Cell(17,3) -Font "Tahoma" -Color 16711680

	This example sets the font to Tahoma and the text color to blue for the cell located
	in the table's 17th row and 3rd column.
	Note: the $Table.Cell(17,3) returns a single Word COM cells object.
#>

Function SetWordCellFormat 
{
	[CmdletBinding(DefaultParameterSetName='Collection')]
	Param (
		# Word COM object cell collection reference
		[Parameter(Mandatory=$true, ValueFromPipeline=$true, ParameterSetName='Collection', Position=0)] [ValidateNotNullOrEmpty()] $Collection,
		# Word COM object individual cell reference
		[Parameter(Mandatory=$true, ParameterSetName='Cell', Position=0)] [ValidateNotNullOrEmpty()] $Cell,
		# Hashtable of cell co-ordinates
		[Parameter(Mandatory=$true, ParameterSetName='Hashtable', Position=0)] [ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Coordinates,
		# Word COM object table reference
		[Parameter(Mandatory=$true, ParameterSetName='Hashtable', Position=1)] [ValidateNotNullOrEmpty()] $Table,
		# Font name
		[Parameter()] [AllowNull()] [string] $Font = $Null,
		# Font color
		[Parameter()] [AllowNull()] $Color = $Null,
		# Font size
		[Parameter()] [ValidateNotNullOrEmpty()] [int] $Size = 0,
		# Cell background color
		[Parameter()] [AllowNull()] $BackgroundColor = $Null,
		# Force solid background color
		[Switch] $Solid,
		[Switch] $Bold,
		[Switch] $Italic,
		[Switch] $Underline
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'." -f $PSCmdlet.ParameterSetName);
	}

	Process 
	{
		Switch ($PSCmdlet.ParameterSetName) 
		{
			'Collection' {
				ForEach($Cell in $Collection) 
				{
					If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Bold) { $Cell.Range.Font.Bold = $true; }
					If($Italic) { $Cell.Range.Font.Italic = $true; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
					If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				} # end foreach
			} # end Collection
			'Cell' 
			{
				If($Bold) { $Cell.Range.Font.Bold = $true; }
				If($Italic) { $Cell.Range.Font.Italic = $true; }
				If($Underline) { $Cell.Range.Font.Underline = 1; }
				If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
				If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
				If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
				If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
				If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
			} # end Cell
			'Hashtable' 
			{
				ForEach($Coordinate in $Coordinates) 
				{
					$Cell = $Table.Cell($Coordinate.Row, $Coordinate.Column);
					If($Bold) { $Cell.Range.Font.Bold = $true; }
					If($Italic) { $Cell.Range.Font.Italic = $true; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
					If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				}
			} # end Hashtable
		} # end switch
	} # end process
}

<#
.Synopsis
	Sets alternate row colors in a Word table
.DESCRIPTION
	This function sets the format of alternate rows within a Word table using the
	specified $BackgroundColor. This function is expensive (in performance terms) as
	it recursively sets the format on alternate rows. It would be better to pick one
	of the predefined table formats (if one exists)? Obviously the more rows, the
	longer it takes :'(

	Note: this function is called by the AddWordTable function if an alternate row
	format is specified.
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 255

	This example sets every-other table (starting with the first) row and sets the
	background color to red (wdColorRed).
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 39423 -Seed Second

	This example sets every other table (starting with the second) row and sets the
	background color to light orange (weColorLightOrange).
#>

Function SetWordTableAlternateRowColor 
{
	[CmdletBinding()]
	Param (
		# Word COM object table reference
		[Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=0)] [ValidateNotNullOrEmpty()] $Table,
		# Alternate row background color
		[Parameter(Mandatory=$true, Position=1)] [ValidateNotNull()] [int] $BackgroundColor,
		# Alternate row starting seed
		[Parameter(ValueFromPipelineByPropertyName=$true, Position=2)] [ValidateSet('First','Second')] [string] $Seed = 'First'
	)

	Process 
	{
		$StartDateTime = Get-Date;
		Write-Debug ("{0}: `t`tSetting alternate table row colors.." -f $StartDateTime);

		## Determine the row seed (only really need to check for 'Second' and default to 'First' otherwise
		If($Seed.ToLower() -eq 'second') 
		{ 
			$StartRowIndex = 2; 
		}
		Else 
		{ 
			$StartRowIndex = 1; 
		}

		For($AlternateRowIndex = $StartRowIndex; $AlternateRowIndex -lt $Table.Rows.Count; $AlternateRowIndex += 2) 
		{ 
			$Table.Rows.Item($AlternateRowIndex).Shading.BackgroundPatternColor = $BackgroundColor;
		}

		## I've put verbose calls in here we can see how expensive this functionality actually is.
		$EndDateTime = Get-Date;
		$ExecutionTime = New-TimeSpan -Start $StartDateTime -End $EndDateTime;
		Write-Debug ("{0}: `t`tDone setting alternate row style color in '{1}' seconds" -f $EndDateTime, $ExecutionTime.TotalSeconds);
	}
}
#endregion

#region general script functions
Function validStateProp( [object] $object, [string] $topLevel, [string] $secondLevel )
{
	#function created 8-jan-2014 by Michael B. Smith
	If( $object )
	{
		If((Get-Member -Name $topLevel -InputObject $object))
		{
			If((Get-Member -Name $secondLevel -InputObject $object.$topLevel))
			{
				Return $True
			}
		}
	}
	Return $False
}

Function validObject( [object] $object, [string] $topLevel )
{
	#function created 8-jan-2014 by Michael B. Smith
	If( $object )
	{
		If((Get-Member -Name $topLevel -InputObject $object))
		{
			Return $True
		}
	}
	Return $False
}

Function ShowScriptOptions
{
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): Add DateTime    : $AddDateTime"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date): Company Name    : $Script:CoName"
	}
	Write-Verbose "$(Get-Date): Computer Name   : $ComputerName"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date): Company Address : $($CompanyAddress)"
		Write-Verbose "$(Get-Date): Company Email   : $($CompanyEmail)"
		Write-Verbose "$(Get-Date): Company Fax     : $($CompanyFax)"
		Write-Verbose "$(Get-Date): Company Phone   : $($CompanyPhone)"
		Write-Verbose "$(Get-Date): Cover Page      : $CoverPage"
	}
	Write-Verbose "$(Get-Date): Details         : $Details"
	Write-Verbose "$(Get-Date): Dev             : $Dev"
	If($Dev)
	{
		Write-Verbose "$(Get-Date): DevErrorFile    : $Script:DevErrorFile"
	}
	Write-Verbose "$(Get-Date): Filename1       : $Script:Filename1"
	If($PDF)
	{
		Write-Verbose "$(Get-Date): Filename2       : $Script:Filename2"
	}
	Write-Verbose "$(Get-Date): Folder          : $Folder"
	Write-Verbose "$(Get-Date): From            : $From"
	Write-Verbose "$(Get-Date): Log             : $($Log)"
	Write-Verbose "$(Get-Date): Save As HTML    : $HTML"
	Write-Verbose "$(Get-Date): Save As PDF     : $PDF"
	Write-Verbose "$(Get-Date): Save As Text    : $Text"
	Write-Verbose "$(Get-Date): Save As Word    : $MSWord"
	Write-Verbose "$(Get-Date): Script Info     : $ScriptInfo"
	Write-Verbose "$(Get-Date): Smtp Port       : $SmtpPort"
	Write-Verbose "$(Get-Date): Smtp Server     : $SmtpServer"
	Write-Verbose "$(Get-Date): Title           : $Script:Title"
	Write-Verbose "$(Get-Date): To              : $To"
	Write-Verbose "$(Get-Date): Use SSL         : $UseSSL"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date): User Name       : $UserName"
	}
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): OS Detected     : $Script:RunningOS"
	Write-Verbose "$(Get-Date): PSUICulture     : $PSUICulture"
	Write-Verbose "$(Get-Date): PSCulture       : $PSCulture"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date): Word version    : $WordProduct"
		Write-Verbose "$(Get-Date): Word language   : $Script:WordLanguageValue"
	}
	Write-Verbose "$(Get-Date): PoSH version    : $($Host.Version)"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): Script start  : $Script:StartTime"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
}

Function SaveandCloseDocumentandShutdownWord
{
	#bug fix 1-Apr-2014
	#reset Grammar and Spelling options back to their original settings
	$Script:Word.Options.CheckGrammarAsYouType = $Script:CurrentGrammarOption
	$Script:Word.Options.CheckSpellingAsYouType = $Script:CurrentSpellingOption

	Write-Verbose "$(Get-Date): Save and Close document and Shutdown Word"
	If($Script:WordVersion -eq $wdWord2010)
	{
		#the $saveFormat below passes StrictMode 2
		#I found this at the following two links
		#http://blogs.technet.com/b/bshukla/archive/2011/09/27/3347395.aspx
		#http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat(v=office.14).aspx
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date): Saving DOCX file"
		}
		If($AddDateTime)
		{
			$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).docx"
			If($PDF)
			{
				$Script:FileName2 += "_$(Get-Date -f yyyy-MM-dd_HHmm).pdf"
			}
		}
		Write-Verbose "$(Get-Date): Running $Script:WordProduct and detected operating system $Script:RunningOS"
		$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
		$Script:Doc.SaveAs([REF]$Script:FileName1, [ref]$SaveFormat)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
			$Script:Doc.SaveAs([REF]$Script:FileName2, [ref]$saveFormat)
		}
	}
	ElseIf($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016)
	{
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date): Saving DOCX file"
		}
		If($AddDateTime)
		{
			$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).docx"
			If($PDF)
			{
				$Script:FileName2 += "_$(Get-Date -f yyyy-MM-dd_HHmm).pdf"
			}
		}
		Write-Verbose "$(Get-Date): Running $Script:WordProduct and detected operating system $Script:RunningOS"
		$Script:Doc.SaveAs2([REF]$Script:FileName1, [ref]$wdFormatDocumentDefault)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$Script:Doc.SaveAs([REF]$Script:FileName2, [ref]$wdFormatPDF)
		}
	}

	Write-Verbose "$(Get-Date): Closing Word"
	$Script:Doc.Close()
	$Script:Word.Quit()
	If($PDF)
	{
		[int]$cnt = 0
		While(Test-Path $Script:FileName1)
		{
			$cnt++
			If($cnt -gt 1)
			{
				Write-Verbose "$(Get-Date): Waiting another 10 seconds to allow Word to fully close (try # $cnt)"
				Start-Sleep -Seconds 10
				$Script:Word.Quit()
				If($cnt -gt 2)
				{
					#kill the winword process

					#find out our session (usually "1" except on TS/RDC or Citrix)
					$SessionID = (Get-Process -PID $PID).SessionId
					
					#Find out if winword is running in our session
					$wordprocess = ((Get-Process 'WinWord' -ea 0) | Where-Object {$_.SessionId -eq $SessionID}).Id
					If($wordprocess -gt 0)
					{
						Write-Verbose "$(Get-Date): Attempting to stop WinWord process # $wordprocess"
						Stop-Process $wordprocess -EA 0
					}
				}
			}
			Write-Verbose "$(Get-Date): Attempting to delete $Script:FileName1 since only $Script:FileName2 is needed (try # $cnt)"
			Remove-Item $Script:FileName1 -EA 0 4>$Null
		}
	}
	Write-Verbose "$(Get-Date): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
	If(Test-Path variable:global:word)
	{
		Remove-Variable -Name word -Scope Global 4>$Null
	}
	$SaveFormat = $Null
	
	#is the winword process still running? kill it

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId

	#Find out if winword is running in our session
	$wordprocess = $Null
	$wordprocess = ((Get-Process 'WinWord' -ea 0) | Where-Object {$_.SessionId -eq $SessionID}).Id
	If($null -ne $wordprocess -and $wordprocess -gt 0)
	{
		Write-Verbose "$(Get-Date): WinWord process is still running. Attempting to stop WinWord process # $wordprocess"
		Stop-Process $wordprocess -EA 0
	}
}

Function SaveandCloseTextDocument
{
	If($AddDateTime)
	{
		$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
	}

	Write-Output $Global:Output | Out-File $Script:Filename1 4>$Null
}

Function SaveandCloseHTMLDocument
{
	Out-File -FilePath $Script:FileName1 -Append -InputObject "<p></p></body></html>" 4>$Null
}

Function SetFileName1andFileName2
{
	Param([string]$OutputFileName)
	
	If($Folder -eq "")
	{
		$pwdpath = $pwd.Path
	}
	Else
	{
		$pwdpath = $Folder
	}

	If($pwdpath.EndsWith("\"))
	{
		#remove the trailing \
		$pwdpath = $pwdpath.SubString(0, ($pwdpath.Length - 1))
	}

	#set $filename1 and $filename2 with no file extension
	If($AddDateTime)
	{
		[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName)"
		If($PDF)
		{
			[string]$Script:FileName2 = "$($pwdpath)\$($OutputFileName)"
		}
	}

	If($MSWord -or $PDF)
	{
		CheckWordPreReq

		If(!$AddDateTime)
		{
			[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName).docx"
			If($PDF)
			{
				[string]$Script:FileName2 = "$($pwdpath)\$($OutputFileName).pdf"
			}
		}

		SetupWord
	}
	ElseIf($Text)
	{
		If(!$AddDateTime)
		{
			[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName).txt"
		}
		ShowScriptOptions
	}
	ElseIf($HTML)
	{
		SetupHTML
		ShowScriptOptions
	}
}

Function TestComputerName
{
	Param([string]$Cname)
	If(![String]::IsNullOrEmpty($CName)) 
	{
		#get computer name
		#first test to make sure the computer is reachable
		Write-Verbose "$(Get-Date): Testing to see if $CName is online and reachable"
		If(Test-Connection -ComputerName $CName -quiet)
		{
			Write-Verbose "$(Get-Date): Server $CName is online."
		}
		Else
		{
			Write-Verbose "$(Get-Date): Computer $CName is offline"
			$ErrorActionPreference = $SaveEAPreference
			Write-Error "`n`n`t`tComputer $CName is offline.`nScript cannot continue.`n`n"
			Exit
		}
	}

	#if computer name is localhost, get actual computer name
	If($CName -eq "localhost")
	{
		$CName = $env:ComputerName
		Write-Verbose "$(Get-Date): Computer name has been renamed from localhost to $CName"
	}

	#if computer name is an IP address, get host name from DNS
	#http://blogs.technet.com/b/gary/archive/2009/08/29/resolve-ip-addresses-to-hostname-using-powershell.aspx
	#help from Michael B. Smith
	$ip = $CName -as [System.Net.IpAddress]
	If($ip)
	{
		$Result = [System.Net.Dns]::gethostentry($ip).AddressList.IPAddressToString
		
		If($? -and $Null -ne $Result)
		{
			$CName = $Result.HostName
			Write-Verbose "$(Get-Date): Computer name has been renamed from $($ip) to $CName"
		}
		Else
		{
			Write-Warning "Unable to resolve $CName to a hostname"
		}
	}
	Else
	{
		#computer is online but for some reason $ComputerName cannot be converted to a System.Net.IpAddress
	}

	$Results = Get-DNSServer -ComputerName $CName -EA 0 3>$Null
	If($Null -ne $Results)
	{
		#the computer is a dns server
		Write-Verbose "$(Get-Date): Computer $CName is a DNS Server"
		Write-Verbose "$(Get-Date): "
		$Script:DNSServerData = $Results
		Return $CName
	}
	ElseIf($Null -eq $Results)
	{
		#the computer is not a dns server
		Write-Verbose "$(Get-Date): Computer $CName is not a DNS Server"
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tComputer $CName is not a DNS Server.`n`n`t`tRerun the script using -ComputerName with a valid DNS server name.`n`n`t`tScript cannot continue.`n`n"
		Exit
	}

	Return $CName
}

Function ProcessDocumentOutput
{
	If($MSWORD -or $PDF)
	{
		SaveandCloseDocumentandShutdownWord
	}
	ElseIf($Text)
	{
		SaveandCloseTextDocument
	}
	ElseIf($HTML)
	{
		SaveandCloseHTMLDocument
	}

	$GotFile = $False

	If($PDF)
	{
		If(Test-Path "$Script:FileName2")
		{
			Write-Verbose "$(Get-Date): $Script:FileName2 is ready for use"
			$GotFile = $True
		}
		Else
		{
			Write-Warning "$(Get-Date): Unable to save the output file, $Script:FileName2"
			Write-Error "Unable to save the output file, $Script:FileName2"
		}
	}
	Else
	{
		If(Test-Path $Script:FileName1)
		{
			Write-Verbose "$(Get-Date): $Script:FileName1 is ready for use"
			$GotFile = $True
		}
		Else
		{
			Write-Warning "$(Get-Date): Unable to save the output file, $Script:FileName1"
			Write-Error "Unable to save the output file, $Script:FileName1"
		}
	}
	
	#email output file if requested
	If($GotFile -and ![System.String]::IsNullOrEmpty( $SmtpServer ))
	{
		If($PDF)
		{
			$emailAttachment = $Script:FileName2
		}
		Else
		{
			$emailAttachment = $Script:FileName1
		}
		SendEmail $emailAttachment
	}
}

Function AbortScript
{
	If($MSWord -or $PDF)
	{
		$Script:Word.quit()
		Write-Verbose "$(Get-Date): System Cleanup"
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
		If(Test-Path variable:global:word)
		{
			Remove-Variable -Name word -Scope Global
		}
	}
	Write-Verbose "$(Get-Date): Script has been aborted"
	$ErrorActionPreference = $SaveEAPreference
	Exit
}
#endregion

#region script setup function
Function ProcessScriptStart
{
	$script:startTime = Get-Date

	$ComputerName = TestComputerName $ComputerName
	$Script:RptDomain = (Get-WmiObject -computername $ComputerName win32_computersystem).Domain
	[string]$Script:Title = "DNS Inventory Report for $Script:RptDomain"
}

Function ProcessScriptEnd
{
	Write-Verbose "$(Get-Date): Script has completed"
	Write-Verbose "$(Get-Date): "

	#http://poshtips.com/measuring-elapsed-time-in-powershell/
	Write-Verbose "$(Get-Date): Script started: $Script:StartTime"
	Write-Verbose "$(Get-Date): Script ended: $(Get-Date)"
	$runtime = $(Get-Date) - $Script:StartTime
	$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds", `
		$runtime.Days, `
		$runtime.Hours, `
		$runtime.Minutes, `
		$runtime.Seconds,
		$runtime.Milliseconds)
	Write-Verbose "$(Get-Date): Elapsed time: $Str"

	If($Dev)
	{
		If($SmtpServer -eq "")
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
		}
		Else
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
		}
	}

	If($ScriptInfo)
	{
		$SIFile = "$($pwd.Path)\DNSInventoryScriptInfo_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
		Out-File -FilePath $SIFile -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Add DateTime       : $AddDateTime" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Company Name       : $Script:CoName" 4>$Null		
		}
		Out-File -FilePath $SIFile -Append -InputObject "ComputerName       : $ComputerName" 4>$Null		
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Company Address    : $CompanyAddress" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Email      : $CompanyEmail" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Fax        : $CompanyFax" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Phone      : $CompanyPhone" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Cover Page         : $CoverPage" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Details            : $Details" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Dev                : $Dev" 4>$Null
		If($Dev)
		{
			Out-File -FilePath $SIFile -Append -InputObject "DevErrorFile       : $Script:DevErrorFile" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Filename1          : $Script:FileName1" 4>$Null
		If($PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Filename2          : $Script:FileName2" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Folder             : $Folder" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "From               : $From" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Log                : $($Log)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As HTML       : $HTML" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As PDF        : $PDF" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As TEXT       : $TEXT" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As WORD       : $MSWORD" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script Info        : $ScriptInfo" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Port          : $SmtpPort" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Server        : $SmtpServer" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "To                 : $To" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Use SSL            : $UseSSL" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "User Name          : $UserName" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "OS Detected        : $RunningOS" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSUICulture        : $PSUICulture" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSCulture          : $PSCulture" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Word version       : $Script:WordProduct" 4>$Null
			Out-File -FilePath $SIFile -Append -InputObject "Word language      : $Script:WordLanguageValue" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "PoSH version       : $($Host.Version)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script start       : $Script:StartTime" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Elapsed time       : $Str" 4>$Null
	}
	
	#V1.09 added
	#stop transcript logging
	If($Log -eq $True) 
	{
		If($Script:StartLog -eq $true) 
		{
			try 
			{
				Stop-Transcript | Out-Null
				Write-Verbose "$(Get-Date): $Script:LogPath is ready for use"
			} 
			catch 
			{
				Write-Verbose "$(Get-Date): Transcript/log stop failed"
			}
		}
	}
	$runtime = $Null
	$Str = $Null
	$ErrorActionPreference = $SaveEAPreference
}
#endregion

#region email function
Function SendEmail
{
	Param([string]$Attachments)
	Write-Verbose "$(Get-Date): Prepare to email"
	
	$emailAttachment = $Attachments
	$emailSubject = $Script:Title
	$emailBody = @"
Hello, <br />
<br />
$Script:Title is attached.
"@ 

	If($Dev)
	{
		Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
	}

	$error.Clear()

	If($UseSSL)
	{
		Write-Verbose "$(Get-Date): Trying to send email using current user's credentials with SSL"
		Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
		-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
		-UseSSL *>$Null
	}
	Else
	{
		Write-Verbose  "$(Get-Date): Trying to send email using current user's credentials without SSL"
		Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
		-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To *>$Null
	}

	$e = $error[0]

	If($e.Exception.ToString().Contains("5.7.57"))
	{
		#The server response was: 5.7.57 SMTP; Client was not authenticated to send anonymous mail during MAIL FROM
		Write-Verbose "$(Get-Date): Current user's credentials failed. Ask for usable credentials."

		If($Dev)
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
		}

		$error.Clear()

		$emailCredentials = Get-Credential -Message "Enter the email account and password to send email"

		If($UseSSL)
		{
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-UseSSL -credential $emailCredentials *>$Null 
		}
		Else
		{
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-credential $emailCredentials *>$Null 
		}

		$e = $error[0]

		If($? -and $Null -eq $e)
		{
			Write-Verbose "$(Get-Date): Email successfully sent using new credentials"
		}
		Else
		{
			Write-Verbose "$(Get-Date): Email was not sent:"
			Write-Warning "$(Get-Date): Exception: $e.Exception" 
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): Email was not sent:"
		Write-Warning "$(Get-Date): Exception: $e.Exception" 
	}
}
#endregion

#region ProcessDNSServer
Function ProcessDNSServer
{
	Write-Verbose "$(Get-Date): Processing DNS Server"
	Write-Verbose "$(Get-Date): `tRetrieving DNS Server Information using Server $ComputerName"
	
	$DNSServerSettings = $Script:DNSServerData.ServerSetting
	$DNSForwarders = $Script:DNSServerData.ServerForwarder
	$DNSServerRecursion = $Script:DNSServerData.ServerRecursion
	$DNSServerCache = $Script:DNSServerData.ServerCache
	$DNSServerScavenging = $Script:DNSServerData.ServerScavenging
	$DNSRootHints = $Script:DNSServerData.ServerRootHint
	$DNSServerDiagnostics = $Script:DNSServerData.ServerDiagnostics
	
	OutputDNSServer $DNSServerSettings $DNSForwarders $DNSServerRecursion $DNSServerCache $DNSServerScavenging $DNSRootHints $DNSServerDiagnostics
}

Function OutputDNSServer
{
	Param([object] $ServerSettings, [object] $DNSForwarders, [object] $ServerRecursion, [object] $ServerCache, [object] $ServerScavenging, [object] $RootHints, [object] $ServerDiagnostics)

	$RootHints = $RootHints | Sort-Object $RootHints.NameServer.RecordData.NameServer
	
	Write-Verbose "$(Get-Date): `t`tOutput DNS Server Settings"
	$txt = "DNS Server Properties"
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 $txt
	}
	ElseIf($Text)
	{
		Line 0 $txt
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 $txt
	}
	
	#Interfaces tab
	Write-Verbose "$(Get-Date): `t`t`tInterfaces"

	#coutesy of MBS
	#if the value does not exist, then All IP Addresses is selected
	$AllIPs = Get-RegistryValue "HKLM:\SYSTEM\CurrentControlSet\Services\DNS\Parameters" "ListenAddresses" $ComputerName
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Interfaces"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		If($Null -eq $AllIPs)
		{
			$ScriptInformation += @{ Data = "Listen on"; Value = "All IP addresses"; }
		}
		Else
		{
			$ips = ""
			ForEach($IP in $ServerSettings.ListeningIPAddress)
			{
				$ips += "$IP`r"
			}
			$ScriptInformation += @{ Data = "Listen on the following IP addresses"; Value = $ips; }
		}
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0 "Interfaces"
		Line 0 "Listen on: "
		If($Null -eq $AllIPs)
		{
			Line 1 "All IP addresses"
		}
		Else
		{
			Line 1 "Only the following IP addresses: " 
			ForEach($IP in $ServerSettings.ListeningIPAddress)
			{
				Line 2 $IP
			}
		}
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 2 0 "Interfaces"
		#WriteHTMLLine 0 0 "Listen on: "
		$rowdata = @()
		If($Null -eq $AllIPs)
		{
			#WriteHTMLLine 0 1 "All IP addresses"
			$columnHeaders = @("Listen on",($htmlsilver -bor $htmlbold),"All IP addresses",$htmlwhite)
		}
		Else
		{
			$ips = ""
			$First = $True
			ForEach($ipa in $ServerSettings.ListeningIPAddress)
			{
				If($First)
				{
					$columnHeaders = @("Listen on the following IP addresses",($htmlsilver -bor $htmlbold),$ipa,$htmlwhite)
				}
				Else
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$ipa,$htmlwhite))
				}
				$First = $False
			}
		}

		$msg = ""
		$columnWidths = @("200","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}

	#Forwarders tab
	Write-Verbose "$(Get-Date): `t`t`tForwarders"
	If($DNSForwarders.UseRootHint)
	{
		$UseRootHints = "Yes"
	}
	Else
	{
		$UseRootHints = "No"
	}
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Forwarders"
		[System.Collections.Hashtable[]] $FwdWordTable = @();
		ForEach($IP in $DNSForwarders.IPAddress.IPAddressToString)
		{
			$Resolved = ResolveIPtoFQDN $IP
			$WordTableRowHash = @{ 
			IPAddress = $IP;
			ServerFQDN = $Resolved;
			}

			$FwdWordTable += $WordTableRowHash;
		}
		$Table = AddWordTable -Hashtable $FwdWordTable `
		-Columns ServerFQDN, IPAddress `
		-Headers "Server FQDN", "IP Address" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 150;
		
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""

		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Number of seconds before forward queries time out"; Value = $DNSForwarders.Timeout; }
		$ScriptInformation += @{ Data = "Use root hint if no forwarders are available"; Value = $UseRootHints; }
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0 "Forwarders"
		ForEach($IP in $DNSForwarders.IPAddress.IPAddressToString)
		{
			$Resolved = ResolveIPtoFQDN $IP
			Line 0 "IP Address`t: " $IP
			Line 0 "Server FQDN`t: " $Resolved
			Line 0 ""
		}
		Line 0 ""
		Line 0 "Number of seconds before forward queries time out: " $DNSForwarders.Timeout
		Line 0 "Use root hint if no forwarders are available: " $UseRootHints
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 2 0 "Forwarders"
		$rowdata = @()
		ForEach($IP in $DNSForwarders.IPAddress.IPAddressToString)
		{
			$Resolved = ResolveIPtoFQDN $IP
			$rowdata += @(,(
			$Resolved,$htmlwhite,
			$IP,$htmlwhite))
		}
		$columnHeaders = @(
		'Server FQDN',($htmlsilver -bor $htmlbold),
		'IP Address',($htmlsilver -bor $htmlbold))

		$msg = ""
		$columnWidths = @("200","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
		
		$rowdata = @()
		$columnHeaders = @("Number of seconds before forward queries time out",($htmlsilver -bor $htmlbold),$DNSForwarders.Timeout.ToString(),$htmlwhite)
		$rowdata += @(,('Use root hint if no forwarders are available',($htmlsilver -bor $htmlbold),$UseRootHints,$htmlwhite))

		$msg = ""
		$columnWidths = @("200","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}
	
	#Advanced tab
	Write-Verbose "$(Get-Date): `t`t`tAdvanced"
	
	$ServerVersion = "$($ServerSettings.MajorVersion).$($ServerSettings.MinorVersion).$($ServerSettings.BuildNumber) (0x{0:X})" -f $ServerSettings.BuildNumber

	If($ServerRecursion.Enable)
	{
		$Recursion = "Not Selected"
	}
	Else
	{
		$Recursion = "Selected"
	}
	
	If($ServerSettings.BindSecondaries)
	{
		$Bind = "Selected"
	}
	Else
	{
		$Bind = "Not Selected"
	}

	If($ServerSettings.StrictFileParsing)
	{
		$FailOnLoad = "Selected"
	}
	Else
	{
		$FailOnLoad = "Not Selected"
	}
	
	If($ServerSettings.RoundRobin)
	{
		$RoundRobin = "Selected"
	}
	Else
	{
		$RoundRobin = "Not Selected"
	}

	If($ServerSettings.LocalNetPriority)
	{
		$NetMask = "Selected"
	}
	Else
	{
		$NetMask = "Not Selected"
	}
	
	If($ServerRecursion.SecureResponse -and $ServerCache.EnablePollutionProtection)
	{
		$Pollution = "Selected"
	}
	Else
	{
		$Pollution = "Not Selected"
	}
	
	If($ServerSettings.EnableDnsSec )
	{
		$DNSSEC = "Selected"
	}
	Else
	{
		$DNSSEC = "Not Selected"
	}
	
	Switch ($ServerSettings.NameCheckFlag)
	{
		0 {$NameCheck = "Strict RFC (ANSI)"; break}
		1 {$NameCheck = "Non RFC (ANSI)"; break}
		2 {$NameCheck = "Multibyte (UTF8)"; break}
		3 {$NameCheck = "All names"; break}
		Default {$NameCheck = "Unknown: NameCheckFlag Value is $($ServerSettings.NameCheckFlag)"}
	}
	
	Switch ($ServerSettings.BootMethod)
	{
		3 {$LoadZone = "From Active Directory and registry"; break}
		2 {$LoadZone = "From registry"; break}
		Default {$LoadZone = "Unknown: BootMethod Value is $($ServerSettings.BootMethod)"; break}
	}
	
	If($ServerScavenging.ScavengingInterval.days -gt 0 -or  $ServerScavenging.ScavengingInterval.hours -gt 0)
	{
		$EnableScavenging = "Selected"
		If($ServerScavenging.ScavengingInterval.days -gt 0)
		{
			$ScavengingInterval = "$($ServerScavenging.ScavengingInterval.days) days"
		}
		ElseIf($ServerScavenging.ScavengingInterval.hours -gt 0)
		{
			$ScavengingInterval = "$($ServerScavenging.ScavengingInterval.hours) hours"
		}
	}
	Else
	{
		$EnableScavenging = "Not Selected"
	}
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Advanced"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Server version number"; Value = $ServerVersion; }
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
		$ScriptInformation = @()
		
		WriteWordLine 0 0 "Server options:"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Disable recursion (also disables forwarders)"; Value = $Recursion; }
		$ScriptInformation += @{ Data = "Enable BIND secondaries"; Value = $Bind; }
		$ScriptInformation += @{ Data = "Fail on load if bad zone data"; Value = $FailOnLoad; }
		$ScriptInformation += @{ Data = "Enable round robin"; Value = $RoundRobin; }
		$ScriptInformation += @{ Data = "Enable netmask ordering"; Value = $NetMask; }
		$ScriptInformation += @{ Data = "Secure cache against pollution"; Value = $Pollution; }
		$ScriptInformation += @{ Data = "Enable DNSSec validation for remote responses"; Value = $DNSSEC; }
		$ScriptInformation += @{ Data = "Name checking"; Value = $NameCheck; }
		$ScriptInformation += @{ Data = "Load zone data on startup"; Value = $LoadZone; }
		$ScriptInformation += @{ Data = "Enable automatic scavenging of stale records"; Value = $EnableScavenging; }
		If($EnableScavenging -eq "Selected")
		{
			$ScriptInformation += @{ Data = "Scavenging period"; Value = $ScavengingInterval; }
		}
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0 "Advanced"
		Line 0 "Server version number: " $ServerVersion
		Line 0 ""
		Line 0 "Server options:"
		Line 0 "Disable recursion (also disables forwarders)`t: " $Recursion
		Line 0 "Enable BIND secondaries`t`t`t`t: " $Bind
		Line 0 "Fail on load if bad zone data`t`t`t: " $FailOnLoad
		Line 0 "Enable round robin`t`t`t`t: " $RoundRobin
		Line 0 "Enable netmask ordering`t`t`t`t: " $NetMask
		Line 0 "Secure cache against pollution`t`t`t: " $Pollution
		Line 0 "Enable DNSSec validation for remote responses`t: " $DNSSEC
		Line 0 ""
		Line 0 "Name checking`t`t`t`t`t: " $NameCheck
		Line 0 "Load zone data on startup`t`t`t: " $LoadZone
		Line 0 ""
		Line 0 "Enable automatic scavenging of stale records`t: " $EnableScavenging
		If($EnableScavenging -eq "Selected")
		{
			Line 0 "Scavenging period`t`t`t`t: " $ScavengingInterval
		}
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 2 0 "Advanced"
		#WriteHTMLLine 0 0 "Server version number: " $ServerVersion
		#WriteHTMLLine 0 0 " "
		$rowdata = @()
		$columnHeaders = @("Server version number",($htmlsilver -bor $htmlbold),$ServerVersion,$htmlwhite)

		$msg = ""
		$columnWidths = @("200","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "

		WriteHTMLLine 0 0 "Server options:"
		$rowdata = @()
		$columnHeaders = @("Disable recursion (also disables forwarders)",($htmlsilver -bor $htmlbold),$Recursion,$htmlwhite)
		$rowdata += @(,('Enable BIND secondaries',($htmlsilver -bor $htmlbold),$Bind,$htmlwhite))
		$rowdata += @(,('Fail on load if bad zone data',($htmlsilver -bor $htmlbold),$FailOnLoad,$htmlwhite))
		$rowdata += @(,('Enable round robin',($htmlsilver -bor $htmlbold),$RoundRobin,$htmlwhite))
		$rowdata += @(,('Enable netmask ordering',($htmlsilver -bor $htmlbold),$NetMask,$htmlwhite))
		$rowdata += @(,('Secure cache against pollution',($htmlsilver -bor $htmlbold),$Pollution,$htmlwhite))
		$rowdata += @(,('Enable DNSSec validation for remote responses',($htmlsilver -bor $htmlbold),$DNSSEC,$htmlwhite))
		$rowdata += @(,('Name checking',($htmlsilver -bor $htmlbold),$NameCheck,$htmlwhite))
		$rowdata += @(,('Load zone data on startup',($htmlsilver -bor $htmlbold),$LoadZone,$htmlwhite))
		$rowdata += @(,('Enable automatic scavenging of stale records',($htmlsilver -bor $htmlbold),$EnableScavenging,$htmlwhite))
		If($EnableScavenging -eq "Selected")
		{
			$rowdata += @(,('Scavenging period',($htmlsilver -bor $htmlbold),$ScavengingInterval,$htmlwhite))
		}

		$msg = ""
		$columnWidths = @("200","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}
	
	#Root Hints tab
	Write-Verbose "$(Get-Date): `t`t`tRoot Hints"
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Root Hints"
		[System.Collections.Hashtable[]] $RootWordTable = @();
		ForEach($RootHint in $RootHints)
		{
			$ip = @()
			$nameServer = $roothint.NameServer.RecordData.NameServer
			$ipAddresses = $roothint.IPAddress ## poorly named property, since it’s an array
			ForEach( $ipAddress in $ipAddresses )
			{
				$data = $ipAddress.RecordData
				$address = Get-Member -Name IPv4Address -InputObject $data
				If( $Null -eq $address )
				{
					$address = Get-Member -Name IPv6Address -InputObject $data
					If( $Null -eq $address )
					{
						#Write-Error “Bad IPAddress”
						Continue
					}
					$ip += "$($data.IPv6Address.IPAddressToString)`r"
				}
				Else
				{
					$ip += "$($data.IPv4Address.IPAddressToString)`r"
				}
			}

			$ip = $ip | Sort-Object -unique
			
			$WordTableRowHash = @{ 
			ServerFQDN = $RootHint.NameServer.RecordData.NameServer;
			IPAddress = $ip;
			}

			$RootWordTable += $WordTableRowHash;
		}
		$Table = AddWordTable -Hashtable $RootWordTable `
		-Columns ServerFQDN, IPAddress `
		-Headers "Server Fully Qualified Domain Name (FQDN)", "IP Address" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 150;
		
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0 "Root Hints"
		ForEach($RootHint in $RootHints)
		{
			$ip = @()
			$nameServer = $roothint.NameServer.RecordData.NameServer
			$ipAddresses = $roothint.IPAddress ## poorly named property, since it’s an array
			ForEach( $ipAddress in $ipAddresses )
			{
				$data = $ipAddress.RecordData
				$address = Get-Member -Name IPv4Address -InputObject $data
				If( $Null -eq $address )
				{
					$address = Get-Member -Name IPv6Address -InputObject $data
					If( $Null -eq $address )
					{
						#Write-Error “Bad IPAddress”
						Continue
					}
					$ip += $data.IPv6Address.IPAddressToString
				}
				Else
				{
					$ip += $data.IPv4Address.IPAddressToString
				}
			}

			$ip = $ip | Sort-Object -unique
			
			Line 0 "Server Fully Qualified Domain Name (FQDN)`t: " $RootHint.NameServer.RecordData.NameServer
			Line 0 "IP Address`t`t`t`t`t: " $ip
			Line 0 ""
		}
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 2 0 "Root Hints"
		$rowdata = @()
		ForEach($RootHint in $RootHints)
		{
			$ip = $Null
			$PrvIP = $Null
			$cnt = 0
			$nameServer = $roothint.NameServer.RecordData.NameServer
			$ipAddresses = $roothint.IPAddress ## poorly named property, since it’s an array
			ForEach( $ipAddress in $ipAddresses )
			{
				$cnt++
				$data = $ipAddress.RecordData
				$address = Get-Member -Name IPv4Address -InputObject $data
				If( $Null -eq $address )
				{
					$address = Get-Member -Name IPv6Address -InputObject $data
					If( $Null -eq $address )
					{
						#Write-Error “Bad IPAddress”
						Continue
					}
					$ip = $data.IPv6Address.IPAddressToString
				}
				Else
				{
					$ip = $data.IPv4Address.IPAddressToString
				}
				
				If($PrvIP -ne $ip)
				{
					If($cnt -eq 1)
					{
						$rowdata += @(,(
						$RootHint.NameServer.RecordData.NameServer,$htmlwhite,
						$ip,$htmlwhite))
					}
					Else
					{
						$rowdata += @(,(
						"",$htmlwhite,
						$ip,$htmlwhite))
					}
				}
				$PrvIP = $ip
			}

		}
		$columnHeaders = @(
		'Server Fully Qualified Domain Name (FQDN)',($htmlsilver -bor $htmlbold),
		'IP Address',($htmlsilver -bor $htmlbold))

		$msg = ""
		$columnWidths = @("200","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}
	
	#Event Logging
	Write-Verbose "$(Get-Date): `t`t`tEvent Logging"
	
	Switch ($ServerDiagnostics.EventLogLevel)
	{
		0 {$LogLevel = "No events"; break}
		1 {$LogLevel = "Errors only"; break}
		2 {$LogLevel = "Errors and warnings"; break}
		4 {$LogLevel = "All events"; break}	#my value is 7, everyone else appears to be 4
		7 {$LogLevel = "All events"; break}	#leaving as separate stmts for now just in case
		Default {$LogLevel = "Unknown: EventLogLevel Value is $($ServerDiagnostics.EventLogLevel)"; break}
	}
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Event Logging"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Log the following events"; Value = $LogLevel; }
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0 "Event Logging"
		Line 0 "Log the following events: " $LogLevel
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 2 0 "Event Logging"
		$rowdata = @()
		$columnHeaders = @("Log the following events",($htmlsilver -bor $htmlbold),$LogLevel,$htmlwhite)

		$msg = ""
		$columnWidths = @("200","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}

}

Function ResolveIPtoFQDN
{
	Param([string]$cname)

	Write-Verbose "$(Get-Date): `t`t`t`tAttempting to resolve $cname"
	
	$ip = $CName -as [System.Net.IpAddress]
	
	If($ip)
	{
		$Result = [System.Net.Dns]::gethostentry($ip)
		
		If($? -and $Null -ne $Result)
		{
			$CName = $Result.HostName
		}
		Else
		{
			$CName = 'Unable to resolve'
		}
	}
	Return $CName
}
#endregion

#region ProcessForwardLookupZones
Function ProcessForwardLookupZones
{
	Write-Verbose "$(Get-Date): Processing Forward Lookup Zones"

	$txt = "Forward Lookup Zones"
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 $txt
	}
	ElseIf($Text)
	{
		Line 0 $txt
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 $txt
	}

	$First = $True
	$DNSZones = $Script:DNSServerData.ServerZone | Where-Object {$_.IsReverseLookupZone -eq $False -and ($_.ZoneType -eq "Primary" -and $_.ZoneName -ne "TrustAnchors" -or $_.ZoneType -eq "Stub" -or $_.ZoneType -eq "Secondary")}
	
	ForEach($DNSZone in $DNSZones)
	{
		If(!$First)
		{
			If($MSWord -or $PDF)
			{
				$Selection.InsertNewPage()
			}
		}
		OutputLookupZone "Forward" $DNSZone
		If($Details)
		{
			ProcessLookupZoneDetails "Forward" $DNSZone
		}
		$First = $False
	}
}
#endregion

#region process lookzone data
Function OutputLookupZone
{
	Param([string] $zType, [object] $DNSZone)

	Write-Verbose "$(Get-Date): `tProcessing $($DNSZone.ZoneName)"
	
	#General tab
	Write-Verbose "$(Get-Date): `t`tGeneral"
	
	#set all the variable to N/A first since some of the variables/properties do not exist for all zones and zone types
	
	$Status = "N/A"
	$ZoneType = "N/A"
	$Replication = "N/A"
	$DynamicUpdate = "N/A"
	$NorefreshInterval = "N/A"
	$RefreshInterval = "N/A"
	$EnableScavenging = "N/A"
	
	If($DNSZone.IsPaused -eq $False)
	{
		$Status = "Running"
	}
	Else
	{
		$Status = "Paused"
	}
	
	If($DNSZone.ZoneType -eq "Primary" -and $DNSZone.IsDsIntegrated -eq $True)
	{
		$ZoneType = "Active Directory-Integrated"
	}
	ElseIf($DNSZone.ZoneType -eq "Primary" -and $DNSZone.IsDsIntegrated -eq $False)
	{
		$ZoneType = "Primary"
	}
	ElseIf($DNSZone.ZoneType -eq "Secondary" -and $DNSZone.IsDsIntegrated -eq $False)
	{
		$ZoneType = "Secondary"
	}
	ElseIf($DNSZone.ZoneType -eq "Stub")
	{
		$ZoneType = "Stub"
	}
	
	Switch ($DNSZone.ReplicationScope)
	{
		"Forest" {$Replication = "All DNS servers in this forest"; break}
		"Domain" {$Replication = "All DNS servers in this domain"; break}
		"Legacy" {$Replication = "All domain controllers in this domain (for Windows 2000 compatibility"; break}
		"None" {$Replication = "Not an Active-Directory-Integrated zone"; break}
		Default {$Replication = "Unknown: $($DNSZone.ReplicationScope)"; break}
	}
	
	If( ( validObject $DNSZone DynamicUpdate ) )
	{
		Switch ($DNSZone.DynamicUpdate)
		{
			"Secure" {$DynamicUpdate = "Secure only"; break}
			"NonsecureAndSecure" {$DynamicUpdate = "Nonsecure and secure"; break}
			"None" {$DynamicUpdate = "None"; break}
			Default {$DynamicUpdate = "Unknown: $($DNSZone.DynamicUpdate)"; break}
		}
	}
	
	If($DNSZone.ZoneType -eq "Primary")
	{
		$ZoneAging = Get-DnsServerZoneAging -Name $DNSZone.ZoneName -ComputerName $ComputerName -EA 0
		
		If($Null -ne $ZoneAging)
		{
			If($ZoneAging.AgingEnabled)
			{
				$EnableScavenging = "Selected"
				If($ZoneAging.NoRefreshInterval.days -gt 0)
				{
					$NorefreshInterval = "$($ZoneAging.NoRefreshInterval.days) days"
				}
				ElseIf($ZoneAging.NoRefreshInterval.hours -gt 0)
				{
					$NorefreshInterval = "$($ZoneAging.NoRefreshInterval.hours) hours"
				}
				If($ZoneAging.RefreshInterval.days -gt 0)
				{
					$RefreshInterval = "$($ZoneAging.RefreshInterval.days) days"
				}
				ElseIf($ZoneAging.RefreshInterval.hours -gt 0)
				{
					$RefreshInterval = "$($ZoneAging.RefreshInterval.hours) hours"
				}
			}
			Else
			{
				$EnableScavenging = "Not Selected"
			}
		}
		Else
		{
			$EnableScavenging = "Unknown"
		}
		
		$ScavengeServers = @()
		
		If($ZoneAging.ScavengeServers -is [array])
		{
			ForEach($Item in $ZoneAging.ScavengeServers)
			{
				$ScavengeServers += $ZoneAging.ScavengeServers.IPAddressToString
			}
		}
		Else
		{
			$ScavengeServers += $ZoneAging.ScavengeServers.IPAddressToString
		}
		
		If($ScavengeServers.Count -eq 0)
		{
			$ScavengeServers += "Not Configured"
		}
	}
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "$($DNSZone.ZoneName) Properties"
		WriteWordLine 3 0 "General"

		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Status"; Value = $Status; }
		$ScriptInformation += @{ Data = "Type"; Value = $ZoneType; }
		$ScriptInformation += @{ Data = "Replication"; Value = $Replication; }
		If($Null -ne $DNSZone.ZoneFile)
		{
			$ScriptInformation += @{ Data = "Zone file name"; Value = $DNSZone.ZoneFile; }
		}
		ElseIf($Null -eq $DNSZone.ZoneFile -and $DNSZone.IsDsIntegrated)
		{
			$ScriptInformation += @{ Data = "Data is stored in Active Directory"; Value = "Yes"; }
		}
		$ScriptInformation += @{ Data = "Dynamic updates"; Value = $DynamicUpdate; }
		If($DNSZone.ZoneType -eq "Primary")
		{
			$ScriptInformation += @{ Data = "Scavenge stale resource records"; Value = $EnableScavenging; }
			If($EnableScavenging -eq "Selected")
			{
				$ScriptInformation += @{ Data = "No-refresh interval"; Value = $NorefreshInterval; }
				$ScriptInformation += @{ Data = "Refresh interval"; Value = $RefreshInterval; }
			}
			$ScriptInformation += @{ Data = "Scavenge servers"; Value = $ScavengeServers[0]; }
			
			$cnt = -1
			ForEach($ScavengeServer in $ScavengeServers)
			{
				$cnt++
				
				If($cnt -gt 0)
				{
					$ScriptInformation += @{ Data = ""; Value = $ScavengeServer; }
				}
			}
		}
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0 "$($DNSZone.ZoneName) Properties"
		Line 1 "General"
		Line 2 "Status`t`t`t`t: " $Status
		Line 2 "Type`t`t`t`t: " $ZoneType
		Line 2 "Replication`t`t`t: " $Replication
		If($Null -ne $DNSZone.ZoneFile)
		{
			Line 2 "Zone file name`t`t`t: " $DNSZone.ZoneFile
		}
		ElseIf($Null -eq $DNSZone.ZoneFile -and $DNSZone.IsDsIntegrated)
		{
			Line 2 "Data stored in Active Directory`t: " "Yes"
		}
		Line 2 "Dynamic updates`t`t`t: " $DynamicUpdate
		If($DNSZone.ZoneType -eq "Primary")
		{
			Line 2 "Scavenge stale resource records`t: " $EnableScavenging
			If($EnableScavenging -eq "Selected")
			{
				Line 2 "No-refresh interval`t`t: " $NorefreshInterval
				Line 2 "Refresh interval`t`t: " $RefreshInterval
			}
			Line 2 "Scavenge servers`t`t: " $ScavengeServers[0]
			
			$cnt = -1
			ForEach($ScavengeServer in $ScavengeServers)
			{
				$cnt++
				
				If($cnt -gt 0)
				{
					Line 6 "  " $ScavengeServer
				}
			}
		}
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 2 0 "$($DNSZone.ZoneName) Properties"
		WriteHTMLLine 3 0 "General"

		$rowdata = @()
		$columnHeaders = @("Status",($htmlsilver -bor $htmlbold),$Status,$htmlwhite)
		$rowdata += @(,('Type',($htmlsilver -bor $htmlbold),$ZoneType,$htmlwhite))
		$rowdata += @(,('Replication',($htmlsilver -bor $htmlbold),$Replication,$htmlwhite))
		If($Null -ne $DNSZone.ZoneFile)
		{
			$rowdata += @(,('Zone file name',($htmlsilver -bor $htmlbold),$DNSZone.ZoneFile,$htmlwhite))
		}
		ElseIf($Null -eq $DNSZone.ZoneFile -and $DNSZone.IsDsIntegrated)
		{
			$rowdata += @(,('Data is stored in Active Directory',($htmlsilver -bor $htmlbold),"Yes",$htmlwhite))
		}
		$rowdata += @(,('Dynamic updates',($htmlsilver -bor $htmlbold),$DynamicUpdate,$htmlwhite))
		If($DNSZone.ZoneType -eq "Primary")
		{
			$rowdata += @(,('Scavenge stale resource records',($htmlsilver -bor $htmlbold),$EnableScavenging,$htmlwhite))
			If($EnableScavenging -eq "Selected")
			{
				$rowdata += @(,('No-refresh interval',($htmlsilver -bor $htmlbold),$NorefreshInterval,$htmlwhite))
				$rowdata += @(,('Refresh interval',($htmlsilver -bor $htmlbold),$RefreshInterval,$htmlwhite))
			}
			$rowdata += @(,('Scavenge servers',($htmlsilver -bor $htmlbold),$ScavengeServers[0],$htmlwhite))
			
			$cnt = -1
			ForEach($ScavengeServer in $ScavengeServers)
			{
				$cnt++
				
				If($cnt -gt 0)
				{
					$rowdata += @(,(' ',($htmlsilver -bor $htmlbold),$ScavengeServer,$htmlwhite))
				}
			}
		}

		$msg = ""
		$columnWidths = @("200","200")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "400"
		WriteHTMLLine 0 0 " "
	}

	#Start of Authority (SOA) tab
	Write-Verbose "$(Get-Date): `t`tStart of Authority (SOA)"

	$Results = Get-DnsServerResourceRecord -zonename $DNSZone.ZoneName -rrtype soa -ComputerName $ComputerName -EA 0

	If($? -and $Null -ne $Results)
	{
		$SOA = $Results[0]
		
		If($SOA.RecordData.RefreshInterval.Days -gt 0)
		{
			$RefreshInterval = "$($SOA.RecordData.RefreshInterval.Days) days"
		}
		ElseIf($SOA.RecordData.RefreshInterval.Hours -gt 0)
		{
			$RefreshInterval = "$($SOA.RecordData.RefreshInterval.Hours) hours"
		}
		ElseIf($SOA.RecordData.RefreshInterval.Minutes -gt 0)
		{
			$RefreshInterval = "$($SOA.RecordData.RefreshInterval.Minutes) minutes"
		}
		ElseIf($SOA.RecordData.RefreshInterval.Seconds -gt 0)
		{
			$RefreshInterval = "$($SOA.RecordData.RefreshInterval.Seconds) seconds"
		}
		Else
		{
			$RefreshInterval = "Unknown"
		}
		
		If($SOA.RecordData.RetryDelay.Days -gt 0)
		{
			$RetryDelay = "$($SOA.RecordData.RetryDelay.Days) days"
		}
		ElseIf($SOA.RecordData.RetryDelay.Hours -gt 0)
		{
			$RetryDelay = "$($SOA.RecordData.RetryDelay.Hours) hours"
		}
		ElseIf($SOA.RecordData.RetryDelay.Minutes -gt 0)
		{
			$RetryDelay = "$($SOA.RecordData.RetryDelay.Minutes) minutes"
		}
		ElseIf($SOA.RecordData.RetryDelay.Seconds -gt 0)
		{
			$RetryDelay = "$($SOA.RecordData.RetryDelay.Seconds) seconds"
		}
		Else
		{
			$RetryDelay = "Unknown"
		}
		
		If($SOA.RecordData.ExpireLimit.Days -gt 0)
		{
			$ExpireLimit = "$($SOA.RecordData.ExpireLimit.Days) days"
		}
		ElseIf($SOA.RecordData.ExpireLimit.Hours -gt 0)
		{
			$ExpireLimit = "$($SOA.RecordData.ExpireLimit.Hours) hours"
		}
		ElseIf($SOA.RecordData.ExpireLimit.Minutes -gt 0)
		{
			$ExpireLimit = "$($SOA.RecordData.ExpireLimit.Minutes) minutes"
		}
		ElseIf($SOA.RecordData.ExpireLimit.Seconds -gt 0)
		{
			$ExpireLimit = "$($SOA.RecordData.ExpireLimit.Seconds) seconds"
		}
		Else
		{
			$ExpireLimit = "Unknown"
		}
		
		If($SOA.RecordData.MinimumTimeToLive.Days -gt 0)
		{
			$MinimumTTL = "$($SOA.RecordData.MinimumTimeToLive.Days) days"
		}
		ElseIf($SOA.RecordData.MinimumTimeToLive.Hours -gt 0)
		{
			$MinimumTTL = "$($SOA.RecordData.MinimumTimeToLive.Hours) hours"
		}
		ElseIf($SOA.RecordData.MinimumTimeToLive.Minutes -gt 0)
		{
			$MinimumTTL = "$($SOA.RecordData.MinimumTimeToLive.Minutes) minutes"
		}
		ElseIf($SOA.RecordData.MinimumTimeToLive.Seconds -gt 0)
		{
			$MinimumTTL = "$($SOA.RecordData.MinimumTimeToLive.Seconds) seconds"
		}
		Else
		{
			$MinimumTTL = "Unknown"
		}
		
		If($MSWord -or $PDF)
		{
			WriteWordLine 3 0 "Start of Authority (SOA)"
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{ Data = "Serial number"; Value = $SOA.RecordData.SerialNumber.ToString(); }
			$ScriptInformation += @{ Data = "Primary server"; Value = $SOA.RecordData.PrimaryServer; }
			$ScriptInformation += @{ Data = "Responsible person"; Value = $SOA.RecordData.ResponsiblePerson; }
			$ScriptInformation += @{ Data = "Refresh interval"; Value = $RefreshInterval; }
			$ScriptInformation += @{ Data = "Retry interval"; Value = $RetryDelay; }
			$ScriptInformation += @{ Data = "Expires after"; Value = $ExpireLimit; }
			$ScriptInformation += @{ Data = "Minimum (default) TTL"; Value = $MinimumTTL; }
			$ScriptInformation += @{ Data = "TTL for this record"; Value = $SOA.TimeToLive.ToString(); }
			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 200;
			$Table.Columns.Item(2).Width = 200;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		ElseIf($Text)
		{
			Line 1 "Start of Authority (SOA)"
			Line 2 "Serial number`t`t`t: " $SOA.RecordData.SerialNumber.ToString()
			Line 2 "Primary server`t`t`t: " $SOA.RecordData.PrimaryServer
			Line 2 "Responsible person`t`t: " $SOA.RecordData.ResponsiblePerson
			Line 2 "Refresh interval`t`t: " $RefreshInterval
			Line 2 "Retry interval`t`t`t: " $RetryDelay
			Line 2 "Expires after`t`t`t: " $ExpireLimit
			Line 2 "Minimum (default) TTL`t`t: " $MinimumTTL
			Line 2 "TTL for this record`t`t: " $SOA.TimeToLive.ToString()
			Line 0 ""
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 3 0 "Start of Authority (SOA)"
			$rowdata = @()
			$columnHeaders = @("Serial number",($htmlsilver -bor $htmlbold),$SOA.RecordData.SerialNumber.ToString(),$htmlwhite)
			$rowdata += @(,('Primary server',($htmlsilver -bor $htmlbold),$SOA.RecordData.PrimaryServer,$htmlwhite))
			$rowdata += @(,('Responsible person',($htmlsilver -bor $htmlbold),$SOA.RecordData.ResponsiblePerson,$htmlwhite))
			$rowdata += @(,('Refresh interval',($htmlsilver -bor $htmlbold),$RefreshInterval,$htmlwhite))
			$rowdata += @(,('Retry interval',($htmlsilver -bor $htmlbold),$RetryDelay,$htmlwhite))
			$rowdata += @(,('Expires after',($htmlsilver -bor $htmlbold),$ExpireLimit,$htmlwhite))
			$rowdata += @(,('Minimum (default) TTL',($htmlsilver -bor $htmlbold),$MinimumTTL,$htmlwhite))
			$rowdata += @(,('TTL for this record',($htmlsilver -bor $htmlbold),$SOA.TimeToLive.ToString(),$htmlwhite))

			$msg = ""
			$columnWidths = @("200","200")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "400"
			WriteHTMLLine 0 0 " "
		}
	}
	Else
	{
		$txt1 = "Start of Authority (SOA)"
		$txt2 = "Start of Authority data could not be retrieved"
		If($MSWord -or $PDF)
		{
			WriteWordLine 3 0 $txt1
			WriteWordLine 0 0 $txt2
			WriteWordLine 0 0 ""
		}
		ElseIf($Text)
		{
			Line 1 $txt1
			Line 0 $txt2
			Line 0 ""
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 3 0 $txt1
			WriteHTMLLine 0 0 $txt2
			WriteHTMLLine 0 0 " "
		}
	}
	
	#Name Servers tab
	Write-Verbose "$(Get-Date): `t`tName Servers"
	$NameServers = Get-DnsServerResourceRecord -zonename $DNSZone.ZoneName -rrtype ns -node -ComputerName $ComputerName -EA 0

	If($? -and $Null -ne $NameServers)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 3 0 "Name Servers"
			[System.Collections.Hashtable[]] $NSWordTable = @();
			ForEach($NS in $NameServers)
			{
				$ipAddress = ([System.Net.Dns]::gethostentry($NS.RecordData.NameServer)).AddressList.IPAddressToString
				
				If($ipAddress -is [array])
				{
					$cnt = -1
					
					ForEach($ip in $ipAddress)
					{
						$cnt++
						
						If($cnt -eq 0)
						{
							$WordTableRowHash = @{ 
							ServerFQDN = $NS.RecordData.NameServer;
							IPAddress = $ip;
							}
						}
						Else
						{
							$WordTableRowHash = @{ 
							ServerFQDN = $NS.RecordData.NameServer;
							IPAddress = $ip;
							}
						}
					}
				}
				Else
				{
					$WordTableRowHash = @{ 
					ServerFQDN = $NS.RecordData.NameServer;
					IPAddress = $ipAddress;
					}
				}

				$NSWordTable += $WordTableRowHash;
			}
			$Table = AddWordTable -Hashtable $NSWordTable `
			-Columns ServerFQDN, IPAddress `
			-Headers "Server Fully Qualified Domain Name (FQDN)", "IP Address" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 200;
			$Table.Columns.Item(2).Width = 200;
			
			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		ElseIf($Text)
		{
			Line 1 "Name Servers:"
			ForEach($NS in $NameServers)
			{
				$ipAddress = ([System.Net.Dns]::gethostentry($NS.RecordData.NameServer)).AddressList.IPAddressToString
				
				Line 2 "Server FQDN`t`t`t: " $NS.RecordData.NameServer
				If($ipAddress -is [array])
				{
					$cnt = -1
					
					ForEach($ip in $ipAddress)
					{
						$cnt++
						
						If($cnt -eq 0)
						{
							Line 2 "IP Address`t`t`t: " $ip
						}
						Else
						{
							Line 6 "  " $ip
						}
					}
				}
				Else
				{
					Line 2 "IP Address`t`t`t: " $ipAddress
				}
				Line 0 ""
			}
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 3 0 "Name Servers"
			$rowdata = @()
			ForEach($NS in $NameServers)
			{
				$ipAddress = ([System.Net.Dns]::gethostentry($NS.RecordData.NameServer)).AddressList.IPAddressToString
				
				If($ipAddress -is [array])
				{
					$cnt = -1
					
					ForEach($ip in $ipAddress)
					{
						$cnt++
						
						If($cnt -eq 0)
						{
							$rowdata += @(,(
							$NS.RecordData.NameServer,$htmlwhite,
							$ip,$htmlwhite))
						}
						Else
						{
							$rowdata += @(,(
							$NS.RecordData.NameServer,$htmlwhite,
							$ip,$htmlwhite))
						}
					}
				}
				Else
				{
					$rowdata += @(,(
					$NS.RecordData.NameServer,$htmlwhite,
					$ipAddress,$htmlwhite))
				}
			}
			$columnHeaders = @(
			'Server Fully Qualified Domain Name (FQDN)',($htmlsilver -bor $htmlbold),
			'IP Address',($htmlsilver -bor $htmlbold))

			$msg = ""
			$columnWidths = @("200","200")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "400"
			WriteHTMLLine 0 0 " "
		}
	}
	Else
	{
		$txt1 = "Name Servers"
		$txt2 = "Name Servers data could not be retrieved"
		If($MSWord -or $PDF)
		{
			WriteWordLine 3 0 $txt1
			WriteWordLine 0 0 $txt2
			WriteWordLine 0 0 ""
		}
		ElseIf($Text)
		{
			Line 1 $txt1
			Line 0 $txt2
			Line 0 ""
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 3 0 $txt1
			WriteHTMLLine 0 0 $txt2
			WriteHTMLLine 0 0 " "
		}
	}

	If($zType -eq "Forward")
	{
		#WINS tab
		Write-Verbose "$(Get-Date): `t`tWINS"
		If( ( validObject $DNSZone IsWinsEnabled ) )
		{
			If($DNSZone.IsWinsEnabled)
			{
				$WINSEnabled = "Selected"
				
				$WINS = Get-DnsServerResourceRecord -zonename $DNSZone.ZoneName -rrtype wins -ComputerName $ComputerName -EA 0
				
				If($? -and $Null -ne $WINS)
				{
					If($WINS.RecordData.Replicate)
					{
						$WINSReplicate = "Selected"
					}
					Else
					{
						$WINSReplicate = "Not selected"
					}

					$ip = @()
					ForEach($ipAddress in $WINS.RecordData.WinsServers)
					{
						$ip += "$ipAddress`r"
					}
					
					If($WINS.RecordData.CacheTimeout.Days -gt 0)
					{
						$CacheTimeout = "$($WINS.RecordData.CacheTimeout.Days) days"
					}
					ElseIf($WINS.RecordData.CacheTimeout.Hours -gt 0)
					{
						$CacheTimeout = "$($WINS.RecordData.CacheTimeout.Hours) hours"
					}
					ElseIf($WINS.RecordData.CacheTimeout.Minutes -gt 0)
					{
						$CacheTimeout = "$($WINS.RecordData.CacheTimeout.Minutes) minutes"
					}
					ElseIf($WINS.RecordData.CacheTimeout.Seconds -gt 0)
					{
						$CacheTimeout = "$($WINS.RecordData.CacheTimeout.Seconds) seconds"
					}
					Else
					{
						$CacheTimeout = "Unknown"
					}

					If($WINS.RecordData.LookupTimeout.Days -gt 0)
					{
						$LookupTimeout = "$($WINS.RecordData.LookupTimeout.Days) days"
					}
					ElseIf($WINS.RecordData.LookupTimeout.Hours -gt 0)
					{
						$LookupTimeout = "$($WINS.RecordData.LookupTimeout.Hours) hours"
					}
					ElseIf($WINS.RecordData.LookupTimeout.Minutes -gt 0)
					{
						$LookupTimeout = "$($WINS.RecordData.LookupTimeout.Minutes) minutes"
					}
					ElseIf($WINS.RecordData.LookupTimeout.Seconds -gt 0)
					{
						$LookupTimeout = "$($WINS.RecordData.LookupTimeout.Seconds) seconds"
					}
					Else
					{
						$LookupTimeout = "Unknown"
					}

					If($MSWord -or $PDF)
					{
						WriteWordLine 3 0 "WINS"
						[System.Collections.Hashtable[]] $ScriptInformation = @()
						$ScriptInformation += @{ Data = "Use WINS forward lookup"; Value = $WINSEnabled; }
						$ScriptInformation += @{ Data = "Do not replicate this record"; Value = $WINSReplicate; }
						$ScriptInformation += @{ Data = "IP address"; Value = $ip; }
						$ScriptInformation += @{ Data = "Cache time-out"; Value = $CacheTimeout; }
						$ScriptInformation += @{ Data = "Lookup time-out"; Value = $LookupTimeout; }
						$Table = AddWordTable -Hashtable $ScriptInformation `
						-Columns Data,Value `
						-List `
						-Format $wdTableGrid `
						-AutoFit $wdAutoFitFixed;

						SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

						$Table.Columns.Item(1).Width = 200;
						$Table.Columns.Item(2).Width = 200;

						$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

						FindWordDocumentEnd
						$Table = $Null
						WriteWordLine 0 0 ""
					}
					ElseIf($Text)
					{
						Line 1 "WINS"
						Line 2 "Use WINS forward lookup`t`t: " $WINSEnabled
						Line 2 "Do not replicate this record`t: " $WINSReplicate
						Line 2 "IP address`t`t`t: " $ip
						Line 2 "Cache time-out`t`t`t: " $CacheTimeout
						Line 2 "Lookup time-out`t`t`t: " $LookupTimeout
						Line 0 ""
					}
					ElseIf($HTML)
					{
						WriteHTMLLine 3 0 "WINS"
						$rowdata = @()
						$columnHeaders = @("Use WINS forward lookup",($htmlsilver -bor $htmlbold),$WINSEnabled,$htmlwhite)
						$rowdata += @(,('Do not replicate this record',($htmlsilver -bor $htmlbold),$WINSReplicate,$htmlwhite))
						$First = $True
						ForEach($ipa in $ip)
						{
							If($First)
							{
								$rowdata += @(,('IP address',($htmlsilver -bor $htmlbold),$ipa,$htmlwhite))
							}
							Else
							{
								$rowdata += @(,('',($htmlsilver -bor $htmlbold),$ipa,$htmlwhite))
							}
							$First = $False
						}
						$rowdata += @(,('Cache time-out',($htmlsilver -bor $htmlbold),$CacheTimeout,$htmlwhite))
						$rowdata += @(,('Lookup time-out',($htmlsilver -bor $htmlbold),$LookupTimeout,$htmlwhite))

						$msg = ""
						$columnWidths = @("200","200")
						FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "400"
						WriteHTMLLine 0 0 " "
					}
				}
				Else
				{
					$txt1 = "WINS"
					$txt2 = "Use WINS forward lookup: $WINSEnabled"
					$txt3 = "Unable to retrieve WINS details"
					If($MSWord -or $PDF)
					{
						WriteWordLine 3 0 $txt1
						WriteWordLine 0 0 $txt2
						WriteWordLine 0 0 $txt3
						WriteWordLine 0 0 ""
					}
					ElseIf($Text)
					{
						Line 1 $txt1
						Line 2 $txt2
						Line 0 $txt3
						Line 0 ""
					}
					ElseIf($HTML)
					{
						WriteHTMLLine 3 0 $txt1
						WriteHTMLLine 0 0 $txt2
						WriteHTMLLine 0 0 $txt3
						WriteHTMLLine 0 0 " "
					}
				}
			}
			Else
			{
				$WINSEnabled = "Not selected"
				If($MSWord -or $PDF)
				{
					WriteWordLine 3 0 "WINS"
					[System.Collections.Hashtable[]] $ScriptInformation = @()
					$ScriptInformation += @{ Data = "Use WINS forward lookup"; Value = $WINSEnabled; }
					$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data,Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed;

					SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Columns.Item(1).Width = 200;
					$Table.Columns.Item(2).Width = 200;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}
				ElseIf($Text)
				{
					Line 1 "WINS"
					Line 2 "Use WINS forward lookup`t`t: " $WINSEnabled
					Line 0 ""
				}
				ElseIf($HTML)
				{
					WriteHTMLLine 3 0 "WINS"
					$rowdata = @()
					$columnHeaders = @("Use WINS forward lookup",($htmlsilver -bor $htmlbold),$WINSEnabled,$htmlwhite)

					$msg = ""
					$columnWidths = @("200","200")
					FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "400"
					WriteHTMLLine 0 0 " "
				}
			}
		}
	}
	ElseIf($zType -eq "Reverse")
	{
		#WINS-R tab
		Write-Verbose "$(Get-Date): `t`tWINS-R"

		If( ( validObject $DNSZone IsWinsEnabled ) )
		{
			If($DNSZone.IsWinsEnabled)
			{
				$WINSEnabled = "Selected"
				
				$WINS = Get-DnsServerResourceRecord -zonename $DNSZone.ZoneName -rrtype winsr -ComputerName $ComputerName -EA 0
				
				If($? -and $Null -ne $WINS)
				{
					If($WINS.RecordData.Replicate)
					{
						$WINSReplicate = "Selected"
					}
					Else
					{
						$WINSReplicate = "Not selected"
					}

					If($WINS.RecordData.CacheTimeout.Days -gt 0)
					{
						$CacheTimeout = "$($WINS.RecordData.CacheTimeout.Days) days"
					}
					ElseIf($WINS.RecordData.CacheTimeout.Hours -gt 0)
					{
						$CacheTimeout = "$($WINS.RecordData.CacheTimeout.Hours) hours"
					}
					ElseIf($WINS.RecordData.CacheTimeout.Minutes -gt 0)
					{
						$CacheTimeout = "$($WINS.RecordData.CacheTimeout.Minutes) minutes"
					}
					ElseIf($WINS.RecordData.CacheTimeout.Seconds -gt 0)
					{
						$CacheTimeout = "$($WINS.RecordData.CacheTimeout.Seconds) seconds"
					}
					Else
					{
						$CacheTimeout = "Unknown"
					}

					If($WINS.RecordData.LookupTimeout.Days -gt 0)
					{
						$LookupTimeout = "$($WINS.RecordData.LookupTimeout.Days) days"
					}
					ElseIf($WINS.RecordData.LookupTimeout.Hours -gt 0)
					{
						$LookupTimeout = "$($WINS.RecordData.LookupTimeout.Hours) hours"
					}
					ElseIf($WINS.RecordData.LookupTimeout.Minutes -gt 0)
					{
						$LookupTimeout = "$($WINS.RecordData.LookupTimeout.Minutes) minutes"
					}
					ElseIf($WINS.RecordData.LookupTimeout.Seconds -gt 0)
					{
						$LookupTimeout = "$($WINS.RecordData.LookupTimeout.Seconds) seconds"
					}
					Else
					{
						$LookupTimeout = "Unknown"
					}

					If($MSWord -or $PDF)
					{
						WriteWordLine 3 0 "WINS-R"
						[System.Collections.Hashtable[]] $ScriptInformation = @()
						$ScriptInformation += @{ Data = "Use WINS-R lookup"; Value = $WINSEnabled; }
						#$ScriptInformation += @{ Data = "Do not replicate this record"; Value = "Can't Find"; }
						$ScriptInformation += @{ Data = "Domain name to append to returned name"; Value = $WINS.RecordData.ResultDomain; }
						$ScriptInformation += @{ Data = "Cache time-out"; Value = $CacheTimeout; }
						$ScriptInformation += @{ Data = "Lookup time-out"; Value = $LookupTimeout; }
						$ScriptInformation += @{ Data = "Submit DNS domain as NetBIOS scope"; Value = $WINSReplicate; }
						$Table = AddWordTable -Hashtable $ScriptInformation `
						-Columns Data,Value `
						-List `
						-Format $wdTableGrid `
						-AutoFit $wdAutoFitFixed;

						SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

						$Table.Columns.Item(1).Width = 200;
						$Table.Columns.Item(2).Width = 200;

						$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

						FindWordDocumentEnd
						$Table = $Null
						WriteWordLine 0 0 ""
					}
					ElseIf($Text)
					{
						Line 1 "WINS-R"
						Line 2 "Use WINS forward lookup`t`t: " $WINSEnabled
						#Line 2 "Do not replicate this record`t: " "Can't Find"
						Line 2 "Domain name to append`t`t: " $WINS.RecordData.ResultDomain
						Line 2 "Cache time-out`t`t`t: " $CacheTimeout
						Line 2 "Lookup time-out`t`t`t: " $LookupTimeout
						Line 2 "Submit DNS domain as NetBIOS scope: " $WINSReplicate
						Line 0 ""
					}
					ElseIf($HTML)
					{
						WriteHTMLLine 3 0 "WINS-R"
						$rowdata = @()
						$columnHeaders = @("Use WINS forward lookup",($htmlsilver -bor $htmlbold),$WINSEnabled,$htmlwhite)
						#$rowdata += @(,('Do not replicate this record',($htmlsilver -bor $htmlbold),"Can't Find",$htmlwhite))
						$rowdata += @(,('Domain name to append to returned name',($htmlsilver -bor $htmlbold),$WINS.RecordData.ResultDomain,$htmlwhite))
						$rowdata += @(,('Cache time-out',($htmlsilver -bor $htmlbold),$CacheTimeout,$htmlwhite))
						$rowdata += @(,('Lookup time-out',($htmlsilver -bor $htmlbold),$LookupTimeout,$htmlwhite))
						$rowdata += @(,('Submit DNS domain as NetBIOS scope',($htmlsilver -bor $htmlbold),$WINSReplicate,$htmlwhite))

						$msg = ""
						$columnWidths = @("200","200")
						FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "400"
						WriteHTMLLine 0 0 " "
					}
				}
				Else
				{
					$txt1 = "WINS"
					$txt2 = "Use WINS forward lookup: $WINSEnabled"
					$txt3 = "Unable to retrieve WINS details"
					If($MSWord -or $PDF)
					{
						WriteWordLine 3 0 $txt1
						WriteWordLine 0 0 $txt2
						WriteWordLine 0 0 $txt3
						WriteWordLine 0 0 ""
					}
					ElseIf($Text)
					{
						Line 1 $txt1
						Line 2 $txt2
						Line 0 $txt3
						Line 0 ""
					}
					ElseIf($HTML)
					{
						WriteHTMLLine 3 0 $txt1
						WriteHTMLLine 0 0 $txt2
						WriteHTMLLine 0 0 $txt3
						WriteHTMLLine 0 0 " "
					}
				}
			}
			Else
			{
				$WINSEnabled = "Not selected"
				If($MSWord -or $PDF)
				{
					WriteWordLine 3 0 "WINS-R"
					[System.Collections.Hashtable[]] $ScriptInformation = @()
					$ScriptInformation += @{ Data = "Use WINS-R lookup"; Value = $WINSEnabled; }
					$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data,Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed;

					SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Columns.Item(1).Width = 200;
					$Table.Columns.Item(2).Width = 200;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}
				ElseIf($Text)
				{
					Line 1 "WINS-R"
					Line 2 "Use WINS-R lookup`t`t: " $WINSEnabled
					Line 0 ""
				}
				ElseIf($HTML)
				{
					WriteHTMLLine 3 0 "WINS-R"
					$rowdata = @()
					$columnHeaders = @("Use WINS-R lookup",($htmlsilver -bor $htmlbold),$WINSEnabled,$htmlwhite)

					$msg = ""
					$columnWidths = @("200","200")
					FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "400"
					WriteHTMLLine 0 0 " "
				}
			}
		}
	}
	
	#Zone Transfers tab
	Write-Verbose "$(Get-Date): `t`tZone Transfers"
	
	If( ( validObject $DNSZone SecureSecondaries ) )
	{
		If($DNSZone.SecureSecondaries -ne "NoTransfer")
		{
			If($DNSZone.SecureSecondaries -eq "TransferAnyServer")
			{
				$ZoneTransfer = "To any server"
			}
			ElseIf($DNSZone.SecureSecondaries -eq "TransferToZoneNameServer")
			{
				$ZoneTransfer = "Only to servers listed on the Name Servers tab"
			}
			ElseIf($DNSZone.SecureSecondaries -eq "TransferToSecureServers")
			{
				$ZoneTransfer = "Only to the following servers"
			}
			Else
			{
				$ZoneTransfer = "Unknown"
			}

			If($ZoneTransfer -eq "Only to the following servers")
			{
				$ipSecondaryServers = ""
				ForEach($ipAddress in $DNSZone.SecondaryServers)
				{
					$ipSecondaryServers += "$ipAddress`r"
				}
			}

			If($DNSZone.Notify -eq "NotifyServers")
			{
				$ipNotifyServers = ""
				ForEach($ipAddress in $DNSZone.NotifyServers)
				{
					$ipNotifyServers += "$ipAddress`r"
				}
			}
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 3 0 "Zone Transfers"
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{ Data = "Allow zone transfers"; Value = $ZoneTransfer; }
				If($ZoneTransfer -eq "Only to the following servers")
				{
					$ScriptInformation += @{ Data = ""; Value = $ipSecondaryServers; }
				}
				If($DNSZone.Notify -eq "NoNotify")
				{
					$ScriptInformation += @{ Data = "Automatically notify"; Value = "Not selected"; }
				}
				ElseIf($DNSZone.Notify -eq "Notify")
				{
					$ScriptInformation += @{ Data = "Automatically notify"; Value = "Servers listed on the Name Servers tab"; }
				}
				ElseIf($DNSZone.Notify -eq "NotifyServers")
				{
					$ScriptInformation += @{ Data = "Automatically notify the following servers"; Value = $ipNotifyServers; }
				}
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 200;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			ElseIf($Text)
			{
				Line 1 "Zone Transfers"
				Line 2 "Allow zone transfers`t`t: " $ZoneTransfer
				If($ZoneTransfer -eq "Only to the following servers")
				{
					ForEach($x in $ipSecondaryServers)
					{
						Line 6 "  " $x
					}
				}
				If($DNSZone.Notify -eq "NoNotify")
				{
					Line 2 "Automatically notify`t`t: Not selected"
				}
				ElseIf($DNSZone.Notify -eq "Notify")
				{
					Line 2 "Automatically notify`t`t: Servers listed on the Name Servers tab"
				}
				ElseIf($DNSZone.Notify -eq "NotifyServers")
				{
					Line 2 "Automatically notify`t`t: The following servers " 
					ForEach($x in $ipNotifyServers)
					{
						Line 6 "  " $x
					}
				}
				Line 0 ""
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 3 0 "Zone Transfers"
				$rowdata = @()
				$columnHeaders = @("Allow zone transfers",($htmlsilver -bor $htmlbold),$ZoneTransfer,$htmlwhite)
				If($ZoneTransfer -eq "Only to the following servers")
				{
					ForEach($ipa in $ipSecondaryServers)
					{
						$rowdata += @(,('',($htmlsilver -bor $htmlbold),$ipa,$htmlwhite))
					}
				}
				If($DNSZone.Notify -eq "NoNotify")
				{
					$rowdata += @(,('Automatically notify',($htmlsilver -bor $htmlbold),"Not selected",$htmlwhite))
				}
				ElseIf($DNSZone.Notify -eq "Notify")
				{
					$rowdata += @(,('Automatically notify',($htmlsilver -bor $htmlbold),"Servers listed on the Name Servers tab",$htmlwhite))
				}
				ElseIf($DNSZone.Notify -eq "NotifyServers")
				{
					$First = $True
					ForEach($ipa in $ipNotifyServers)
					{
						If($First)
						{
							$rowdata += @(,('Automatically notify the following servers',($htmlsilver -bor $htmlbold),$ipa,$htmlwhite))
						}
						Else
						{
							$rowdata += @(,('',($htmlsilver -bor $htmlbold),$ipa,$htmlwhite))
						}
						$First = $False
					}
				}
				$msg = ""
				$columnWidths = @("200","200")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "400"
				WriteHTMLLine 0 0 " "
			}
		}
		Else
		{
			$ZoneTransfer = "Not selected"
			If($MSWord -or $PDF)
			{
				WriteWordLine 3 0 "Zone Transfers"
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{ Data = "Allow zone transfers"; Value = $ZoneTransfer; }
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 200;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			ElseIf($Text)
			{
				Line 1 "Zone Transfers"
				Line 2 "Allow zone transfers`t`t: " $ZoneTransfer
				Line 0 ""
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 3 0 "Zone Transfers"
				$rowdata = @()
				$columnHeaders = @("Allow zone transfers",($htmlsilver -bor $htmlbold),$ZoneTransfer,$htmlwhite)

				$msg = ""
				$columnWidths = @("200","200")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "400"
				WriteHTMLLine 0 0 " "
			}
		}
	}
}
#endregion

#region lookup zone details
Function ProcessLookupZoneDetails
{
	Param([string] $zType, [object] $DNSZone)

	Write-Verbose "$(Get-Date): `t`tProcessing details for zone $($DNSZone.ZoneName)"
	
#	$ZoneDetails = Get-DNSServerResourceRecord -ZoneName $DNSZone.ZoneName -ComputerName $ComputerName -EA 0 | `
#	Where-Object {!($_.hostname -like "_*" -or $_.hostname -eq "DomainDNSZones" -or $_.hostname -eq "ForestDNSZones")}	
	
	$ZoneDetails = Get-DNSServerResourceRecord -ZoneName $DNSZone.ZoneName -ComputerName $ComputerName -EA 0

	If($? -and $Null -ne $ZoneDetails)
	{
		OutputLookupZoneDetails $ztype $ZoneDetails $DNSZone.ZoneName
	}
	ElseIf($? -and $Null -eq $ZoneDetails)
	{
		$txt = "There are no Resource Records for zone $($DNSZone.ZoneName)"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 $txt
			WriteWordLine 0 0 ""
		}
		ElseIf($Text)
		{
			Line 0 $txt
			Line 0 ""
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 $txt
			WriteHTMLLine 0 0 " "
		}
	}
	Else
	{
		$txt = "Resource Records for zone $($DNSZone.ZoneName) could not be retrieved"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 $txt
			WriteWordLine 0 0 ""
		}
		ElseIf($Text)
		{
			Line 0 $txt
			Line 0 ""
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 $txt
			WriteHTMLLine 0 0 " "
		}
	}
	
}

Function OutputLookupZoneDetails
{
	Param([string] $zType, [object] $ZoneDetails, [string] $ZoneName)
	
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 3 0 "Resource Records"
	}
	ElseIf($Text)
	{
		Line 1 "Resource Records"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 3 0 "Resource Records"
	}

	If($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $WordTable = @();
	}
	ElseIf($HTML)
	{
		$rowdata = @()
	}
	
	$ipprefix = ""
	If($zType -eq "Reverse")
	{
		$tmpArray = $ZoneName.Split(".")
		If($tmpArray[2] -eq "in-addr" -or $tmpArray[2] -eq "arpa")
		{
			$tmpArray[2] = ""
		}
		If($tmpArray[1] -eq "in-addr" -or $tmpArray[1] -eq "arpa")
		{
			$tmpArray[1] = ""
		}

   		$ipprefix = "$($tmparray[2]).$($tmparray[1]).$($tmparray[0])."
	}
	
	#https://technet.microsoft.com/en-us/library/cc958958.aspx
	
	<#
		-- A (GUI)
		-- AAAA (GUI)
		-- Afsdb (GUI)
		-- Atma (GUI)
		-- CName (GUI)
		-- DhcId (GUI)
		-- DName (GUI)
		-- DnsKey (GUI)
		-- DS (GUI)
		-- Gpos (???)
		-- HInfo (GUI)
		-- Isdn (GUI)
		-- Key (GUI)
		-- Loc (???)
		-- Mb (GUI)
		-- Md (???)
		-- Mf (???)
		-- Mg (GUI)
		-- MInfo (GUI)
		-- Mr (GUI)
		-- Mx (GUI)
		-- Naptr (GUI)
		-- NasP (???)
		-- NasPtr (???)
		-- Ns (GUI)
		-- NSec (Created by DNSSEC)
		-- NSec3 (Created by DNSSEC)
		-- NSec3Param (Created by DNSSEC)
		-- NsNxt (???)
		-- Ptr (GUI)
		-- Rp (GUI)
		-- RRSig (Created by DNSSEC)
		-- Rt (GUI)
		-- Soa (GUI)
		-- Srv (GUI)
		-- Txt (GUI)
		-- Wins (Cmdlet)
		-- WinsR (Cmdlet)
		-- Wks (GUI)
		-- X25 (GUI)
	#>	
	
	ForEach($Detail in $ZoneDetails)
	{
		$tmpType = ""
		Switch ($Detail.RecordType)
		{
			"A"				{$tmpType = "HOST (A)"; break}
			"AAAA"			{$tmpType = "IPv6 HOST (AAAA)"; break}
			"AFSDB"			{$tmpType = "AFS Database (AFSDB)"; break}
			"ATMA"			{$tmpType = "ATM Address (ATMA)"; break}
			"CNAME"			{$tmpType = "Alias (CNAME)"; break}
			"DHCID"			{$tmpType = "DHCID"; break}
			"DNAME"			{$tmpType = "Domain Alias (DNAME)"; break}
			"DNSKEY"		{$tmpType = "DNS KEY (DNSKEY)"; break}
			"DS"			{$tmpType = "Delegation Signer (DS)"; break}
			"HINFO"			{$tmpType = "Host Information (HINFO)"; break}
			"ISDN"			{$tmpType = "ISDN"; break}
			"KEY"			{$tmpType = "Public Key (KEY)"; break}
			"MB"			{$tmpType = "Mailbox (MB)"; break}
			"MG"			{$tmpType = "Mail Group (MG)"; break}
			"MINFO"			{$tmpType = "Mailbox Information (MINFO)"; break}
			"MR"			{$tmpType = "Renamed Mailbox (MR)"; break}
			"MX"			{$tmpType = "Mail Exchanger (MX)"; break}
			"NAPTR"			{$tmpType = "Naming Authority Pointer (NAPTR)"; break}
			"NS"			{$tmpType = "Name Server (NS)"; break}
			"NSEC"			{$tmpType = "Next Secure (NSEC)"; break}
			"NSEC3"			{$tmpType = "Next Secure 3 (NSEC3)"; break}
			"NSEC3PARAM"	{$tmpType = "Next Secure 3 Parameters (NSEC3PARAM)"; break}
			"NXT"			{$tmpType = "Next Domain (NXT)"; break}
			"PTR"			{$tmpType = "Pointer (PTR)"; break}
			"RP"			{$tmpType = "Responsible Person (RP)"; break}
			"RRSIG"			{$tmpType = "RR Signature (RRSIG)"; break}
			"RT"			{$tmpType = "Route Through (RT)"; break}
			"SIG"			{$tmpType = "Signature (SIG)"; break}
			"SOA"			{$tmpType = "Start of Authority (SOA)"; break}
			"SRV"			{$tmpType = "Service Location (SRV)"; break}
			"TXT"			{$tmpType = "Text (TXT)"; break}
			"WINS"			{$tmpType = "WINS Lookup"; break}
			"WINSR"			{$tmpType = "WINS Reverse Lookup (WINS-R_"; break}
			"WKS"			{$tmpType = "Well Known Services (WKS)"; break}
			"X25"			{$tmpType = "X.25"; break}
			Default 		{$tmpType = "Unable to determine Record Type: $($Detail.RecordType)"; break}
		}
			
		If($zType -eq "Reverse")	#V1.09 fixed from = to -eq
		{
			If($Detail.HostName -eq "@")
			{
				$xHostName = "(same as parent folder)"
			}
			Else
			{
                If($Detail.RecordData.PtrDomainName -eq "localhost.")
                {
    				$xHostName = "127.0.0.1"
                }
                Else
                {
    				$xHostName = "$($ipprefix)$($Detail.HostName)"
                }
			}
		}
		Else
		{
			$xHostName = $Detail.HostName #V1.09 change from "" 
		}

		#The follow resource record types are obsolete and do not return any RecordData value
		# KEY, MB, MG, MINFO, MR, NXT, SIG
		# NAPTR is not obsolete but returns no data in RecordData
		
		$DetailData = ""
		If($Detail.HostName -eq "@" -and $Detail.RecordType -eq "NS")
		{
			$DetailData = $Detail.RecordData.NameServer
		}
		ElseIf($Detail.HostName -eq "@" -and $Detail.RecordType -eq "SOA")
		{
			$DetailData = "[$($Detail.RecordData.SerialNumber)], $($Detail.RecordData.PrimaryServer), $($Detail.RecordData.ResponsiblePerson)"
		}
		ElseIf($Detail.HostName -eq "@" -and $Detail.RecordType -eq "A")
		{
			$DetailData = $Detail.RecordData.IPv4Address
		}
		ElseIf($Detail.RecordType -eq "NS")
		{
			$DetailData = $Detail.RecordData.NameServer
		}
		ElseIf($Detail.RecordType -eq "A")
		{
			$DetailData = $Detail.RecordData.IPv4Address
		}
		ElseIf($Detail.RecordType -eq "AAAA")
		{
			$DetailData = $Detail.RecordData.IPv6Address
		}
		ElseIf($Detail.RecordType -eq "AFSDB")
		{
			$tmp = ""
			If($Detail.RecordData.SubType -eq 1)
			{
				$tmp = "AFS"
			}
			ElseIf($Detail.RecordData.SubType -eq 2)
			{
				$tmp = "DCE"
			}
			Else
			{
				$tmp = $Detail.RecordData.SubType
			}
			$DetailData = "[$($tmp)] $($Detail.RecordData.ServerName)"
		}
		ElseIf($Detail.RecordType -eq "ATMA")
		{
			$tmp = ""
			If($Detail.RecordData.AddressType -eq "E164")
			{
				$tmp = "E164"
			}
			ElseIf($Detail.RecordData.AddressType -eq "AESA")
			{
				$tmp = "NSAP"
			}
			$DetailData = "($($tmp)) $($Detail.RecordData.Address)"
		}
		ElseIf($Detail.RecordType -eq "CNAME")
		{
			$DetailData = $Detail.RecordData.HostNameAlias
		}
		ElseIf($Detail.RecordType -eq "DHCID")
		{
			$DetailData = $Detail.RecordData.DHCID
		}
		ElseIf($Detail.RecordType -eq "DNAME")
		{
			$DetailData = $Detail.RecordData.DomainNameAlias
		}
		ElseIf($Detail.RecordType -eq "DNSKEY")
		{
			$Crypto = ""
			Switch ($Detail.RecordData.CryptoAlgorithm) 
			{
				"ECDsaP256Sha256"	{$Crypto = "ECDSAP256/SHA-256"; break}
				"ECDsaP384Sha384"	{$Crypto = "ECDSAP384/SHA-384"; break}
				"RsaSha1"			{$Crypto = "RSA/SHA-1"; break}
				"RsaSha1NSec3"		{$Crypto = "RSA/SHA-1 (NSEC)"; break}
				"RsaSha256"			{$Crypto = "RSA/SHA-256"; break}
				"RsaSha512"			{$Crypto = "RSA/SHA-512"; break}
				Default 			{$Crypto = "Unknown CryptoAlgorithm: $($Detail.RecordData.CryptoAlgorithm)"; break}
			}
			
			$DetailData = "[$($Detail.RecordData.KeyFlags)][DNSSEC][$($Crypto)][$($Detail.RecordData.KeyTag)]"
		}
		ElseIf($Detail.RecordType -eq "DS")
		{
			$Crypto = ""
			Switch ($Detail.RecordData.CryptoAlgorithm) 
			{
				"ECDsaP256Sha256"	{$Crypto = "ECDSAP256/SHA-256"; break}
				"ECDsaP384Sha384"	{$Crypto = "ECDSAP384/SHA-384"; break}
				"RsaSha1"			{$Crypto = "RSA/SHA-1"; break}
				"RsaSha1NSec3"		{$Crypto = "RSA/SHA-1 (NSEC)"; break}
				"RsaSha256"			{$Crypto = "RSA/SHA-256"; break}
				"RsaSha512"			{$Crypto = "RSA/SHA-512"; break}
				Default 			{$Crypto = "Unknown CryptoAlgorithm: $($Detail.RecordData.CryptoAlgorithm)"; break}
			}
			
			$DigestType = ""
			Switch ($Detail.RecordData.DigestType)
			{
				"Sha1"		{$DigestType = "SHA-1"; break}
				"Sha256"	{$DigestType = "SHA-256"; break}
				"Sha384"	{$DigestType = "SHA-384"; break}
				Default		{$DigestType = "Unknown DigestType: $($Detail.RecordData.DigestType)"; break}
			}
			$DetailData = "[$($Detail.RecordData.KeyTag)][$($DigestType)][$($Crypto)][$($Detail.RecordData.Digest)]"
		}
		ElseIf($Detail.RecordType -eq "HINFO")
		{
			$DetailData = "$($Detail.RecordData.CPU), $($Detail.RecordData.OperatingSystem)"
		}
		ElseIf($Detail.RecordType -eq "ISDN")
		{
			$DetailData = "$($Detail.RecordData.IsdnNumber), $($Detail.RecordData.IsdnSubAddress)"
		}
		ElseIf($Detail.RecordType -eq "MB")
		{
			$DetailData = $Detail.RecordData
		}
		ElseIf($Detail.RecordType -eq "KEY")
		{
			$DetailData = $Detail.RecordData
		}
		ElseIf($Detail.RecordType -eq "MG")
		{
			$DetailData = $Detail.RecordData
		}
		ElseIf($Detail.RecordType -eq "MINFO")
		{
			$DetailData = $Detail.RecordData
		}
		ElseIf($Detail.RecordType -eq "MR")
		{
			$DetailData = $Detail.RecordData
		}
		ElseIf($Detail.RecordType -eq "MX")
		{
			$DetailData = "[$($Detail.RecordData.Preference)] $($Detail.RecordData.MailExchange)"
		}
		ElseIf($Detail.RecordType -eq "NAPTR")
		{
			$DetailData = $Detail.RecordData
		}
		ElseIf($Detail.RecordType -eq "NSEC")
		{
			$CoveredRecordTypes = ""
			
			If($Null -ne $Detail.RecordData.CoveredRecordTypes)
			{
				ForEach($Item in $Detail.RecordData.CoveredRecordTypes)
				{
					$CoveredRecordTypes += "$($Item) "
				}
			}
			
			$DetailData = "[$($Detail.RecordData.Name)][$($CoveredRecordTypes)]"
		}
		ElseIf($Detail.RecordType -eq "NSEC3")
		{
			$Crypto = ""
			Switch ($Detail.RecordData.HashAlgorithm) 
			{
				"ECDsaP256Sha256"	{$Crypto = "ECDSAP256/SHA-256"; break}
				"ECDsaP384Sha384"	{$Crypto = "ECDSAP384/SHA-384"; break}
				"RsaSha1"			{$Crypto = "RSA/SHA-1"; break}
				"RsaSha1NSec3"		{$Crypto = "RSA/SHA-1 (NSEC)"; break}
				"RsaSha256"			{$Crypto = "RSA/SHA-256"; break}
				"RsaSha512"			{$Crypto = "RSA/SHA-512"; break}
				Default 			{$Crypto = "Unknown CryptoAlgorithm: $($Detail.RecordData.HashAlgorithm)"; break}
			}

			$OptOut = "NO Opt-Out"
			If($Detail.RecordData.OptOut -eq $True)
			{
				$OptOut = "YES Opt-Out"
			}

			$CoveredRecordTypes = ""
			
			If($Null -ne $Detail.RecordData.CoveredRecordTypes)
			{
				ForEach($Item in $Detail.RecordData.CoveredRecordTypes)
				{
					$CoveredRecordTypes += "$($Item) "
				}
			}
			
			$DetailData = "[$($Crypto)][$($OptOut)][$($Detail.RecordData.Iterations)][$($Detail.RecordData.Salt)][$($Detail.RecordData.NextHashedOwnerName)][$($CoveredRecordTypes)]"
		}
		ElseIf($Detail.RecordType -eq "NSEC3PARAM")
		{
			$Crypto = ""
			Switch ($Detail.RecordData.HashAlgorithm) 
			{
				"ECDsaP256Sha256"	{$Crypto = "ECDSAP256/SHA-256"; break}
				"ECDsaP384Sha384"	{$Crypto = "ECDSAP384/SHA-384"; break}
				"RsaSha1"			{$Crypto = "RSA/SHA-1"; break}
				"RsaSha1NSec3"		{$Crypto = "RSA/SHA-1 (NSEC)"; break}
				"RsaSha256"			{$Crypto = "RSA/SHA-256"; break}
				"RsaSha512"			{$Crypto = "RSA/SHA-512"; break}
				Default 			{$Crypto = "Unknown CryptoAlgorithm: $($Detail.RecordData.HashAlgorithm)"; break}
			}
			
			$Timestamp = ""
			
			If($Null -eq $Detail.Timestamp )
			{
				$Timestamp = "0"
			}
			Else
			{
				$Timestamp = $Detail.Timestamp
			}
			
			$DetailData = "[$($Crypto)][$($Timestamp)][$($Detail.RecordData.Iterations)][$($Detail.RecordData.Salt)]"
		}
		ElseIf($Detail.RecordType -eq "NXT")
		{
			$DetailData = $Detail.RecordData
		}
		ElseIf($Detail.RecordType -eq "PTR")
		{
			$DetailData = $Detail.RecordData.PtrDomainName
		}
		ElseIf($Detail.RecordType -eq "RP")
		{
			$DetailData = "$($Detail.RecordData.ResponsiblePerson), $($Detail.RecordData.Description)"
		}
		ElseIf($Detail.RecordType -eq "RRSIG")
		{
			$Crypto = ""
			Switch ($Detail.RecordData.CryptoAlgorithm) 
			{
				"ECDsaP256Sha256"	{$Crypto = "ECDSAP256/SHA-256"; break}
				"ECDsaP384Sha384"	{$Crypto = "ECDSAP384/SHA-384"; break}
				"RsaSha1"			{$Crypto = "RSA/SHA-1"; break}
				"RsaSha1NSec3"		{$Crypto = "RSA/SHA-1 (NSEC)"; break}
				"RsaSha256"			{$Crypto = "RSA/SHA-256"; break}
				"RsaSha512"			{$Crypto = "RSA/SHA-512"; break}
				Default 			{$Crypto = "Unknown CryptoAlgorithm: $($Detail.RecordData.CryptoAlgorithm)"; break}
			}
			
			$InceptionDate = $Detail.RecordData.SignatureInception.ToUniversalTime().ToShortDateString()
			$InceptionTime = $Detail.RecordData.SignatureInception.ToUniversalTime().ToLongTimeString()
			$ExpirationDate = $Detail.RecordData.SignatureExpiration.ToUniversalTime().ToShortDateString()
			$ExpirationTime = $Detail.RecordData.SignatureExpiration.ToUniversalTime().ToLongTimeString()
			
			$DetailData = "[$($Detail.RecordData.TypeCovered)][Inception(UTC): $($InceptionDate) $($InceptionTime)][Expiration(UTC): $($ExpirationDate) $($ExpirationTime)][$($Detail.RecordData.NameSigner)][$($Crypto)][$($Detail.RecordData.LabelCount)][$($Detail.RecordData.KeyTag)]"
		}
		ElseIf($Detail.RecordType -eq "RT")
		{
			$DetailData = "[$($Detail.RecordData.Preference)] $($Detail.RecordData.IntermediateHost)"
		}
		ElseIf($Detail.RecordType -eq "SIG")
		{
			$DetailData = $Detail.RecordData
		}
		ElseIf($Detail.RecordType -eq "SRV")
		{
			$DetailData = "[$($Detail.RecordData.Priority)][$($Detail.RecordData.Weight))][$($Detail.RecordData.Port))][$($Detail.RecordData.DomainName)]"
		}
		ElseIf($Detail.RecordType -eq "TXT")
		{
			$DetailData = "$($Detail.RecordData.DescriptiveText)"
		}
		ElseIf($Detail.RecordType -eq "WINS")
		{
			$xServer = ""
			ForEach($xData in $Detail.RecordData.WinsServers)
			{
				$xServer += "$($xData) "
			}
			$DetailData = "[$($xServer)]"
		}
		ElseIf($Detail.RecordType -eq "WINSR")
		{
			$DetailData = $Detail.RecordData.ResultDomain
		}
		ElseIf($Detail.RecordType -eq "WKS")
		{
			$xService = ""
			ForEach($xData in $Detail.RecordData.Service)
			{
				$xService += "$($xData) "
			}
			$DetailData = "[$($Detail.RecordData.InternetProtocol)] $xService"
		}
		ElseIf($Detail.RecordType -eq "X25")
		{
			$DetailData = $Detail.RecordData.PSDNAddress
		}
		Else
		{
			$DetailData = "Unknown: RR=$($Detail.RecordType), RecordData=$($Detail.RecordData)"
		}

		If($Null -eq $Detail.TimeStamp)
		{
			$TimeStamp = "Static"
		}
		Else
		{
			$TimeStamp = $Detail.TimeStamp
		}
		
		If($MSWord -or $PDF)
		{
			$WordTableRowHash = @{
			DetailHostName = $xHostName; 
			DetailType = $tmpType; 
			DetailData = $DetailData; 
			DetailTimeStamp = $TimeStamp; 
			}
			$WordTable += $WordTableRowHash;
		}
		ElseIf($Text)
		{
			Line 2 "Name`t`t: " $xHostName
			Line 2 "Type`t`t: " $tmpType
			Line 2 "Data`t`t: " $DetailData
			Line 2 "Timestamp`t: " $TimeStamp
			Line 0 ""
		}
		ElseIf($HTML)
		{
			$rowdata += @(,(
			$xHostName,$htmlwhite,
			$tmpType,$htmlwhite,
			$DetailData,$htmlwhite,
			$TimeStamp,$htmlwhite))
		}
	}
	
	If($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $WordTable `
		-Columns  DetailHostName, DetailType, DetailData, DetailTimeStamp `
		-Headers  "Name", "Type", "Data", "Timestamp" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table -Size 9
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 105;
		$Table.Columns.Item(2).Width = 105;
		$Table.Columns.Item(3).Width = 155;
		$Table.Columns.Item(4).Width = 110;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($HTML)
	{
		$columnHeaders = @(
		'Name',($htmlsilver -bor $htmlbold),
		'Type',($htmlsilver -bor $htmlbold),
		'Data',($htmlsilver -bor $htmlbold),
		'Timestamp',($htmlsilver -bor $htmlbold)
		)

		$columnWidths = @("150","150","100","150")
		$msg = ""
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "550"
		WriteHTMLLine 0 0 " "
	}
	
	
}
#endregion

#region ProcessReverseLookupZones
Function ProcessReverseLookupZones
{
	Write-Verbose "$(Get-Date): Processing Reverse Lookup Zones"
	
	$txt = "Reverse Lookup Zones"
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 $txt
	}
	ElseIf($Text)
	{
		Line 0 $txt
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 $txt
	}

	$First = $True
	$DNSZones = $Script:DNSServerData.ServerZone | Where-Object {$_.IsReverseLookupZone -eq $True}
	
	ForEach($DNSZone in $DNSZones)
	{
		If(!$First)
		{
			If($MSWord -or $PDF)
			{
				$Selection.InsertNewPage()
			}
		}
		OutputLookupZone "Reverse" $DNSZone
		If($Details)
		{
			ProcessLookupZoneDetails "Reverse" $DNSZone
		}
		$First = $False
	}
}
#endregion

#region ProcessTrustPoints
Function ProcessTrustPoints
{
	Write-Verbose "$(Get-Date): Processing Trust Points"
	
	$txt = "Trust Points"
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 $txt
	}
	ElseIf($Text)
	{
		Line 0 $txt
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 $txt
	}

	$TrustPoints = Get-DNSServerTrustPoint -ComputerName $ComputerName -EA 0
	
	If($? -and $Null -ne $TrustPoints)
	{
		ForEach($Trust in $TrustPoints)
		{
		
			$Anchors = Get-DnsServerTrustAnchor -name $Trust.TrustPointName -ComputerName $ComputerName -EA 0
			
			If($? -and $Null -ne $Anchors)
			{
				$First = $True
				ForEach($Anchor in $Anchors)
				{
					If(!$First)
					{
						If($MSWord -or $PDF)
						{
							$Selection.InsertNewPage()
						}
					}
					OutputTrustPoint $Trust $Anchor
				}
			}
			$First = $False
		}
	}
	ElseIf($? -and $Null -ne $TrustPoints)
	{
		$txt1 = "Trust Zones"
		$txt2 = "There is no Trust Zones data"
		If($MSWord -or $PDF)
		{
			WriteWordLine 3 0 $txt1
			WriteWordLine 0 0 $txt2
			WriteWordLine 0 0 ""
		}
		ElseIf($Text)
		{
			Line 1 $txt1
			Line 0 $txt2
			Line 0 ""
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 3 0 $txt1
			WriteHTMLLine 0 0 $txt2
			WriteHTMLLine 0 0 " "
		}
	}
	Else
	{
		$txt1 = "Trust Zones"
		$txt2 = "Trust Zones data could not be retrieved"
		If($MSWord -or $PDF)
		{
			WriteWordLine 3 0 $txt1
			WriteWordLine 0 0 $txt2
			WriteWordLine 0 0 ""
		}
		ElseIf($Text)
		{
			Line 1 $txt1
			Line 0 $txt2
			Line 0 ""
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 3 0 $txt1
			WriteHTMLLine 0 0 $txt2
			WriteHTMLLine 0 0 " "
		}
	}
}

Function OutputTrustPoint
{
	Param([object] $Trust, [object] $Anchor)

	Write-Verbose "$(Get-Date): `tProcessing $Trust.TrustPointName"
	
	If($Anchor.TrustAnchorData.ZoneKey)
	{
		$ZoneKey = "Selected"
		Switch ($Anchor.TrustAnchorData.KeyProtocol)
		{
			"DnsSec" {$KeyProtocol = "DNSSEC"}
			Default {$KeyProtocol = "Unknown: Zone Key Protocol = $($Anchor.TrustAnchorData.KeyProtocol)"}
		}
	}
	Else
	{
		$ZoneKey = "Not Selected"
		$KeyProtocol = "N/A"
	}

	If($Anchor.TrustAnchorData.SecureEntryPoint)
	{
		$SEP = "Selected"
		Switch ($Anchor.TrustAnchorData.CryptoAlgorithm)
		{	
			"RsaSha1"		{$SEPAlgorithm = "RSA/SHA-1"; break}
			"RsaSha1NSec3"	{$SEPAlgorithm = "RSA/SHA-1 (NSEC3)"; break}
			"RsaSha256"		{$SEPAlgorithm = "RSA/SHA-256"; break}
			"RsaSha512"		{$SEPAlgorithm = "RSA/SHA-512"; break}
			Default 		{$SEPAlgorithm = "Unknown: Algorithm = $($Anchor.TrustAnchorData.CryptoAlgorithm)"; break}
		}
	}
	Else
	{
		$SEP = "Not Selected"
		$SEPAlgorithm = "N/A"
	}
	
	If($MSWord -or $PDF)
	{
		If($Trust.TrustPointName -eq ".")
		{
			WriteWordLine 3 0 "$($Trust.TrustPointName)(root) Properties"
		}
		Else
		{
			WriteWordLine 3 0 "$($Trust.TrustPointName) Properties"
		}
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Name"; Value = "(Same as parent folder)"; }
		$ScriptInformation += @{ Data = "Status"; Value = $Trust.TrustPointState; }
		#$ScriptInformation += @{ Data = "Type"; Value = "Can't find"; }
		$ScriptInformation += @{ Data = "Valid From"; Value = $Trust.LastActiveRefreshTime; }
		$ScriptInformation += @{ Data = "Valid To"; Value = $Trust.NextActiveRefreshTime; }
		$ScriptInformation += @{ Data = "Fully qualified domain name (FQDN)"; Value = $Trust.TrustPointName; }
		$ScriptInformation += @{ Data = "Key Tag"; Value = $Anchor.TrustAnchorData.KeyTag; }
		$ScriptInformation += @{ Data = "Zone Key"; Value = $ZoneKey; }
		$ScriptInformation += @{ Data = "Protocol"; Value = $KeyProtocol; }
		$ScriptInformation += @{ Data = "Secure Entry Point"; Value = $SEP; }
		$ScriptInformation += @{ Data = "Algorithm"; Value = $SEPAlgorithm; }
		#$ScriptInformation += @{ Data = "Delete this record when it becomes stale"; Value = "Can't find"; }
		#$ScriptInformation += @{ Data = "Record time stamp"; Value = "Can't find"; }
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		If($Trust.TrustPointName -eq ".")
		{
			Line 0 "$($Trust.TrustPointName)(root) Properties"
		}
		Else
		{
			Line 0 "$($Trust.TrustPointName) Properties"
		}
		Line 1 "Name`t`t`t`t`t`t: (Same as parent folder)"
		Line 1 "Status`t`t`t`t`t`t: " $Trust.TrustPointState
		#Line 1 "Type`t`t`t`t`t`t: " "Can't find"
		Line 1 "Valid From`t`t`t`t`t: " $Trust.LastActiveRefreshTime
		Line 1 "Valid To`t`t`t`t`t: " $Trust.NextActiveRefreshTime
		Line 1 "Fully qualified domain name (FQDN)`t`t: " $Trust.TrustPointName
		Line 1 "Key Tag`t`t`t`t`t`t: " $Anchor.TrustAnchorData.KeyTag
		Line 1 "Zone Key`t`t`t`t`t: " $ZoneKey
		Line 1 "Protocol`t`t`t`t`t: " $KeyProtocol
		Line 1 "Secure Entry Point`t`t`t`t: " $SEP
		Line 1 "Algorithm`t`t`t`t`t: " $SEPAlgorithm
		#Line 1 "Delete this record when it becomes stale`t: " "Can't find"
		#Line 1 "Record time stamp`t`t`t`t: " "Can't find"
		Line 0 ""
	}
	ElseIf($HTML)
	{
		If($Trust.TrustPointName -eq ".")
		{
			WriteHTMLLine 3 0 "$($Trust.TrustPointName)(root) Properties"
		}
		Else
		{
			WriteHTMLLine 3 0 "$($Trust.TrustPointName) Properties"
		}
		$rowdata = @()
		$columnHeaders = @("Name",($htmlsilver -bor $htmlbold),"(Same as parent folder)",$htmlwhite)
		$rowdata += @(,('Status',($htmlsilver -bor $htmlbold),$Trust.TrustPointState,$htmlwhite))
		#$rowdata += @(,('Type',($htmlsilver -bor $htmlbold),"Can't find",$htmlwhite))
		$rowdata += @(,('Valid From',($htmlsilver -bor $htmlbold),$Trust.LastActiveRefreshTime,$htmlwhite))
		$rowdata += @(,('Valid To',($htmlsilver -bor $htmlbold),$Trust.NextActiveRefreshTime,$htmlwhite))
		$rowdata += @(,('Fully qualified domain name (FQDN)',($htmlsilver -bor $htmlbold),$Trust.TrustPointName,$htmlwhite))
		$rowdata += @(,('Key Tag',($htmlsilver -bor $htmlbold),$Anchor.TrustAnchorData.KeyTag,$htmlwhite))
		$rowdata += @(,('Zone Key',($htmlsilver -bor $htmlbold),$ZoneKey,$htmlwhite))
		$rowdata += @(,('Protocol',($htmlsilver -bor $htmlbold),$KeyProtocol,$htmlwhite))
		$rowdata += @(,('Secure Entry Point',($htmlsilver -bor $htmlbold),$SEP,$htmlwhite))
		$rowdata += @(,('Algorithm',($htmlsilver -bor $htmlbold),$SEPAlgorithm,$htmlwhite))
		#$rowdata += @(,('Delete this record when it becomes stale',($htmlsilver -bor $htmlbold),"Can't find",$htmlwhite))
		#$rowdata += @(,('Record time stamp',($htmlsilver -bor $htmlbold),"Can't find",$htmlwhite))

		$msg = ""
		$columnWidths = @("200","200")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "400"
		WriteHTMLLine 0 0 " "
	}
}
#endregion

#region ProcessConditionalForwarders
Function ProcessConditionalForwarders
{
	Write-Verbose "$(Get-Date): Processing Conditional Forwarders"
	
	$txt = "Conditional Forwarders"
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 $txt
	}
	ElseIf($Text)
	{
		Line 0 $txt
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 $txt
	}

	$First = $True
	$DNSZones = $Script:DNSServerData.ServerZone | Where-Object {$_.ZoneType -eq "Forwarder"}
	
	If($? -and $Null -ne $DNSZones)
	{
		ForEach($DNSZone in $DNSZones)
		{
			If(!$First)
			{
				If($MSWord -or $PDF)
				{
					$Selection.InsertNewPage()
				}
			}
			OutputConditionalForwarder $DNSZone
			$First = $False
		}
	}
	ElseIf($? -and $Null -ne $DNSZones)
	{
		$txt1 = "Conditional Forwarders"
		$txt2 = "There is no Conditional Forwarders data"
		If($MSWord -or $PDF)
		{
			WriteWordLine 3 0 $txt1
			WriteWordLine 0 0 $txt2
			WriteWordLine 0 0 ""
		}
		ElseIf($Text)
		{
			Line 1 $txt1
			Line 0 $txt2
			Line 0 ""
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 3 0 $txt1
			WriteHTMLLine 0 0 $txt2
			WriteHTMLLine 0 0 " "
		}
	}
	Else
	{
		$txt1 = "Conditional Forwarders"
		$txt2 = "Conditional Forwarders data could not be retrieved"
		If($MSWord -or $PDF)
		{
			WriteWordLine 3 0 $txt1
			WriteWordLine 0 0 $txt2
			WriteWordLine 0 0 ""
		}
		ElseIf($Text)
		{
			Line 1 $txt1
			Line 0 $txt2
			Line 0 ""
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 3 0 $txt1
			WriteHTMLLine 0 0 $txt2
			WriteHTMLLine 0 0 " "
		}
	}
}

Function OutputConditionalForwarder
{
	Param([object] $DNSZone)

	Write-Verbose "$(Get-Date): `tProcessing $($DNSZone.ZoneName)"
	
	#General tab
	Write-Verbose "$(Get-Date): `t`tGeneral"
	Switch ($DNSZone.ReplicationScope)
	{
		"Forest" {$Replication = "All DNS servers in this forest"; break}
		"Domain" {$Replication = "All DNS servers in this domain"; break}
		"Legacy" {$Replication = "All domain controllers in this domain (for Windows 2000 compatibility"; break}
		"None" {$Replication = "Not an Active-Directory-Integrated zone"; break}
		Default {$Replication = "Unknown: $($DNSZone.ReplicationScope)"; break}
	}
	
	$IPAddresses = $DNSZone.MasterServers
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "$($DNSZone.ZoneName) Properties"
		WriteWordLine 3 0 "General"

		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Type"; Value = "Conditional Forwarder"; }
		$ScriptInformation += @{ Data = "Replication"; Value = $Replication; }
		$ScriptInformation += @{ Data = "Number of seconds before forward queries time out"; Value = $DNSZone.ForwarderTimeout; }
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""

		WriteWordLine 3 0 "Master Servers"
		[System.Collections.Hashtable[]] $NSWordTable = @();
		ForEach($ip in $IPAddresses)
		{
			$Resolved = ResolveIPtoFQDN $IP
			
			$WordTableRowHash = @{ 
			IPAddress = $IP;
			ServerFQDN = $Resolved;
			}

			$NSWordTable += $WordTableRowHash;
		}
		$Table = AddWordTable -Hashtable $NSWordTable `
		-Columns ServerFQDN, IPAddress `
		-Headers "Server FQDN", "IP Address" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 200;
		
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0 "$($DNSZone.ZoneName) Properties"
		Line 1 "General"
		Line 2 "Type`t`t`t`t`t`t: " "Conditional Forwarder"
		Line 2 "Replication`t`t`t`t`t: " $Replication
		Line 2 "# of seconds before forward queries time out`t: " $DNSZone.ForwarderTimeout
		Line 0 ""

		Line 1 "Master Servers:"
		ForEach($ip in $IPAddresses)
		{
			$Resolved = ResolveIPtoFQDN $IP.IPAddressToString
			
   			Line 2 "Server FQDN`t`t`t`t`t: " $Resolved
			Line 2 "IP Address`t`t`t`t`t: " $ip.IPAddressToString
			Line 0 ""
		}
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 2 0 "$($DNSZone.ZoneName) Properties"
		WriteHTMLLine 3 0 "General"
		$rowdata = @()
		$columnheaders = @('Type',($htmlsilver -bor $htmlbold),"Conditional Forwarder",$htmlwhite)
		$rowdata += @(,('Replication',($htmlsilver -bor $htmlbold),$Replication,$htmlwhite))
		$rowdata += @(,('Number of seconds before forward queries time out',($htmlsilver -bor $htmlbold),$DNSZone.ForwarderTimeout,$htmlwhite))

		$msg = ""
		$columnWidths = @("200","200")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "400"
		WriteHTMLLine 0 0 " "

		WriteHTMLLine 3 0 "Master Servers"
		$rowdata = @()
		ForEach($ip in $IPAddresses)
		{
			$Resolved = ResolveIPtoFQDN $IP
			$rowdata += @(,(
			$Resolved,$htmlwhite,
			$IP,$htmlwhite))
		}
		$columnHeaders = @(
		'Server FQDN',($htmlsilver -bor $htmlbold),
		'IP Address',($htmlsilver -bor $htmlbold))

		$msg = ""
		$columnWidths = @("200","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}
}
#endregion

#region script core
#Script begins

ProcessScriptStart

SetFileName1andFileName2 "$($Script:RptDomain)_DNS"

ProcessDNSServer

ProcessForwardLookupZones

ProcessReverseLookupZones

ProcessTrustPoints

ProcessConditionalForwarders
#endregion

#region finish script
Write-Verbose "$(Get-Date): Finishing up document"
#end of document processing

$AbstractTitle = "Microsoft DNS Inventory Report"
$SubjectTitle = "DNS Inventory Report"

UpdateDocumentProperties $AbstractTitle $SubjectTitle

ProcessDocumentOutput

ProcessScriptEnd

#endregion
