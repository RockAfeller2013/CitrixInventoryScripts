#Requires -Version 3.0
#This File is in Unicode format.  Do not edit in an ASCII editor.

<#
.SYNOPSIS
	Creates a complete inventory of a Citrix XenApp 5 for Windows Server 2008 farm using Microsoft Word 2010 or 2013.
.DESCRIPTION
	Creates a complete inventory of a Citrix XenApp 5 for Windows Server 2008 farm using Microsoft Word and PowerShell.
	Works for XenApp 5 Server 2008 32-bit and 64-bit
	Creates either a Word document or PDF named after the XenApp 5 for Windows Server 2008 farm.
	Document includes a Cover Page, Table of Contents and Footer.
	Version 4.xx includes support for the following language versions of Microsoft Word:
		Catalan
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
		
.PARAMETER CompanyName
	Company Name to use for the Cover Page.  
	Default value is contained in HKCU:\Software\Microsoft\Office\Common\UserInfo\CompanyName or
	HKCU:\Software\Microsoft\Office\Common\UserInfo\Company, whichever is populated on the 
	computer running the script.
	This parameter has an alias of CN.
.PARAMETER CoverPage
	What Microsoft Word Cover Page to use.
	Only Word 2010 and 2013 are supported.
	(default cover pages in Word en-US)
	Valid input is:
		Alphabet (Word 2010. Works)
		Annual (Word 2010. Doesn't work well for this report)
		Austere (Word 2010. Works)
		Austin (Word 2010/2013. Doesn't work in 2013, mostly works in 2010 but Subtitle/Subject & Author fields need to me moved after title box is moved up)
		Banded (Word 2013. Works)
		Conservative (Word 2010. Works)
		Contrast (Word 2010. Works)
		Cubicles (Word 2010. Works)
		Exposure (Word 2010. Works if you like looking sideways)
		Facet (Word 2013. Works)
		Filigree (Word 2013. Works)
		Grid (Word 2010/2013.Works in 2010)
		Integral (Word 2013. Works)
		Ion (Dark) (Word 2013. Top date doesn't fit, box needs to be manually resized or font changed to 8 point)
		Ion (Light) (Word 2013. Top date doesn't fit, box needs to be manually resized or font changed to 8 point)
		Mod (Word 2010. Works)
		Motion (Word 2010/2013. Works if top date is manually changed to 36 point)
		Newsprint (Word 2010. Works but date is not populated)
		Perspective (Word 2010. Works)
		Pinstripes (Word 2010. Works)
		Puzzle (Word 2010. Top date doesn't fit, box needs to be manually resized or font changed to 14 point)
		Retrospect (Word 2013. Works)
		Semaphore (Word 2013. Works)
		Sideline (Word 2010/2013. Doesn't work in 2013, works in 2010)
		Slice (Dark) (Word 2013. Doesn't work)
		Slice (Light) (Word 2013. Doesn't work)
		Stacks (Word 2010. Works)
		Tiles (Word 2010. Date doesn't fit unless changed to 26 point)
		Transcend (Word 2010. Works)
		ViewMaster (Word 2013. Works)
		Whisp (Word 2013. Works)
	Default value is Sideline.
	This parameter has an alias of CP.
.PARAMETER UserName
	User name to use for the Cover Page and Footer.
	Default value is contained in $env:username
	This parameter has an alias of UN.
.PARAMETER PDF
	SaveAs PDF file instead of DOCX file.
	This parameter is disabled by default.
	For Word 2007, the Microsoft add-in for saving as a PDF muct be installed.
	For Word 2007, please see http://www.microsoft.com/en-us/download/details.aspx?id=9943
	The PDF file is roughly 5X to 10X larger than the DOCX file.
.PARAMETER MSWord
	SaveAs DOCX file
	This parameter is set True if no other output format is selected.
.PARAMETER Hardware
	Use WMI to gather hardware information on: Computer System, Disks, Processor and Network Interface Cards
	This parameter is disabled by default.
.PARAMETER Software
	Gather software installed by querying the registry.  
	Use SoftwareExclusions.txt to exclude software from the report.
	SoftwareExclusions.txt must exist, and be readable, in the same folder as this script.
	SoftwareExclusions.txt can be an empty file to have no installed applications excluded.
	See Get-Help About-Wildcards for help on formatting the lines to exclude applications.
	This parameter is disabled by default.
.PARAMETER StartDate
	Start date, in MM/DD/YYYY HH:MM format, for the Configuration Logging report.
	Default is today's date minus seven days.
	If the StartDate is entered as 01/01/2014, the date becomes 01/01/2014 00:00:00.
.PARAMETER EndDate
	End date, in MM/DD/YYYY HH:MM format, for the Configuration Logging report.
	Default is today's date.
	If the EndDate is entered as 01/01/2014, the date becomes 01/01/2014 00:00:00.
.PARAMETER Summary
	Only give summary information, no details.
	This parameter is disabled by default.
	This parameter cannot be used with either the Hardware, Software, StartDate or EndDate parameters..EXAMPLE
	PS C:\PSScript > .\XA52008_Inventory_V42.ps1
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Conservative for the Cover Page format.
	Administrator for the User Name.
.PARAMETER AddDateTime
	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2014 at 6PM is 2014-06-01_1800.
	Output filename will be ReportName_2014-06-01_1800.docx (or .pdf).
	This parameter is disabled by default.
.PARAMETER Section
	Processes a specific section of the report.
	Valid options are:
		Admins (Administrators)
		Apps (Applications)
		ConfigLog (Configuration Logging)
		Farm
		LoadEvals (Load Evaluators)
		Policies
		Printers (Print Drivers and Print Driver Mappings)
		Servers
		Zones
		All
	This parameter defaults to All sections.
.EXAMPLE
	PS C:\PSScript > .\XA52008_Inventory_V42.ps1
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Conservative for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XA52008_Inventory_V42.ps1 -PDF
	
	Will use all Default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Conservative for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XA52008_Inventory_V42.ps1 -Summary
	
	Creates a Summary report with no detail.
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XA52008_Inventory_V42.ps1 -PDF -Summary
	
	Creates a Summary report with no detail.
	Will use all Default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XA52008_Inventory_V42.ps1 -Hardware
	
	Will use all Default values and add additional information for each server about its hardware.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XA52008_Inventory_V42.ps1 -StartDate "01/01/2014" -EndDate "01/02/2014"
	
	Will use all Default values and add additional information for each server about its installed applications.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	Will return all Configuration Logging entries from "01/01/2014 00:00:00" through "01/02/2014 "00:00:00".
.EXAMPLE
	PS C:\PSScript > .\XA52008_Inventory_V42.ps1 -StartDate "01/01/2014" -EndDate "01/01/2014"
	
	Will use all Default values and add additional information for each server about its installed applications.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	Will return all Configuration Logging entries from "01/01/2014 00:00:00" through "01/01/2014 "00:00:00".  In other words, nothing is returned.
.EXAMPLE
	PS C:\PSScript > .\XA52008_Inventory_V42.ps1 -StartDate "01/01/2014 21:00:00" -EndDate "01/01/2014 22:00:00"
	
	Will use all Default values and add additional information for each server about its installed applications.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	Will return all Configuration Logging entries from 9PM to 10PM on 01/01/2014.
.EXAMPLE
	PS C:\PSScript .\XA52008_Inventory_V42.ps1 -CompanyName "Carl Webster Consulting" -CoverPage "Mod" -UserName "Carl Webster"

	Will use:
		Carl Webster Consulting for the Company Name.
		Mod for the Cover Page format.
		Carl Webster for the User Name.
.EXAMPLE
	PS C:\PSScript .\XA52008_Inventory_V42.ps1 -CN "Carl Webster Consulting" -CP "Mod" -UN "Carl Webster"

	Will use:
		Carl Webster Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Carl Webster for the User Name (alias UN).
.EXAMPLE
	PS C:\PSScript > .\XA52008_Inventory_V42.ps1 -Section Policies
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	Processes only the Policies section of the report.
.EXAMPLE
	PS C:\PSScript > .\XA52008_Inventory_V42.ps1 -AddDateTime
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2014 at 6PM is 2014-06-01_1800.
	Output filename will be XA5FarmName_2014-06-01_1800.docx
.EXAMPLE
	PS C:\PSScript > .\XA52008_Inventory_V42.ps1 -PDF -AddDateTime
	
	Will use all Default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	Will display verbose messages as the script is running.

	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2014 at 6PM is 2014-06-01_1800.
	Output filename will be XA5FarmName_2014-06-01_1800.pdf
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  This script creates a Word or PDF document.
.NOTES
	NAME: XA52008_Inventory_V42.ps1
	VERSION: 4.2
	AUTHOR: Carl Webster (with a lot of help from Michael B. Smith and Jeff Wouters)
	LASTEDIT: August 4, 2014
#>


#thanks to @jeffwouters and Michael B. Smith for helping me with these parameters
[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "Word") ]

Param(
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[Switch]$MSWord=$False,

	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[Switch]$PDF=$False,

	[parameter(Mandatory=$False)] 
	[Switch]$Hardware=$False, 

	[parameter(Mandatory=$False)] 
	[Switch]$Software=$False,

	[parameter(Mandatory=$False)] 
	[Datetime]$StartDate = ((Get-Date -displayhint date).AddDays(-7)),

	[parameter(Mandatory=$False)] 
	[Datetime]$EndDate = (Get-Date -displayhint date),
	
	[parameter(Mandatory=$False)] 
	[Switch]$Summary=$False,	
	
	[parameter(Mandatory=$False)] 
	[Switch]$AddDateTime=$False,
	
	[parameter(Mandatory=$False)] 
	[string]$Section="All",
	
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[Alias("CN")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyName="",
    
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[Alias("CP")]
	[ValidateNotNullOrEmpty()]
	[string]$CoverPage="Sideline", 

	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[Alias("UN")]
	[ValidateNotNullOrEmpty()]
	[string]$UserName=$env:username

	)


#Original Script created 8/17/2010 by Michael Bogobowicz, Citrix Systems.
#To contact, please message @mikebogo on Twitter
#The original script was designed to be run on a XenApp 6 server

#Modifications by Carl Webster, CTP and independent consultant
#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com
#modified from original script for XenApp 5
#originally released to the Citrix community on October 3, 2011
#Version 4.2
#	Fix the SWExclusions function to work if SoftwareExclusions.txt file contains only one item
#	Cleanup the script's parameters section
#	Code cleanup and standardization with the master template script
#	Requires PowerShell V3 or later
#	Removed support for Word 2007
#	Word 2007 references in help text removed
#	Cover page parameter now states only Word 2010 and 2013 are supported
#	Most Word 2007 references in script removed:
#		Function ValidateCoverPage
#		Function SetupWord
#		Function SaveandCloseDocumentandShutdownWord
#	Function CheckWord2007SaveAsPDFInstalled removed
#	If Word 2007 is detected, an error message is now given and the script is aborted
#	Cleanup Word table code for the first row and background color
#	Cleanup retrieving services and service startup type with Iain Brighton's optimization
#	Add Iain Brighton's Word table functions
#	Move Citrix Services table to new table functions
#	Move hardware info to new table functions
#	Move Citrix and Microsoft hotfix tables to new table functions
#	Move Appendix A and B tables to new table function
#	Add more write statements and error handling to the Configuration Logging report section
#	Added beginning and ending dates for retrieving Configuration Logging data
#	Add Section parameter
#	Valid Section options are:
#		Admins (Administrators)
#		Apps (Applications)
#		ConfigLog (Configuration Logging)
#		Farm
#		LoadEvals (Load Evaluators)
#		Policies
#		Printers (Print Drivers and Print Driver Mappings)
#		Servers
#		Zones
#		All

Set-StrictMode -Version 2

#force  on
$PSDefaultParameterValues = @{"*:Verbose"=$True}
$SaveEAPreference = $ErrorActionPreference
$ErrorActionPreference = 'SilentlyContinue'

If($MSWord -eq $Null)
{
	$MSWord = $False
}
If($PDF -eq $Null)
{
	$PDF = $False
}
If($Hardware -eq $Null)
{
	$Hardware = $False
}
If($Software -eq $Null)
{
	$Software = $False
}
If($StartDate -eq $Null)
{
	$StartDate = ((Get-Date -displayhint date).AddDays(-7))
}
If($EndDate -eq $Null)
{
	$EndDate = (Get-Date -displayhint date)
}
If($Summary -eq $Null)
{
	$Summary = $False
}
If($AddDateTime -eq $Null)
{
	$AddDateTime = $False
}
If($Section -eq $Null)
{
	$Section = "All"
}

If(!(Test-Path Variable:MSWord))
{
	$MSWord = $False
}
If(!(Test-Path Variable:PDF))
{
	$PDF = $False
}
If(!(Test-Path Variable:Hardware))
{
	$Hardware = $False
}
If(!(Test-Path Variable:Software))
{
	$Software = $False
}
If(!(Test-Path Variable:StartDate))
{
	$StartDate = ((Get-Date -displayhint date).AddDays(-7))
}
If(!(Test-Path Variable:EndDate))
{
	$EndDate = ((Get-Date -displayhint date))
}
If(!(Test-Path Variable:Summary))
{
	$Summary = $False
}
If(!(Test-Path Variable:AddDateTime))
{
	$AddDateTime = $False
}
If(!(Test-Path Variable:Section))
{
	$Section = "All"
}

If($MSWord -eq $Null)
{
	If($PDF)
	{
		$MSWord = $False
	}
	Else
	{
		$MSWord = $True
	}
}

If($MSWord -eq $False -and $PDF -eq $False)
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
Else
{
	$ErrorActionPreference = $SaveEAPreference
	Write-Verbose "$(Get-Date): Unable to determine output parameter"
	If($MSWord -eq $Null)
	{
		Write-Verbose "$(Get-Date): MSWord is Null"
	}
	ElseIf($PDF -eq $Null)
	{
		Write-Verbose "$(Get-Date): PDF is Null"
	}
	Else
	{
		Write-Verbose "$(Get-Date): MSWord is $($MSWord)"
		Write-Verbose "$(Get-Date): PDF is $($PDF)"
	}
	Write-Error "Unable to determine output parameter.  Script cannot continue"
	Exit
}

$ValidSection = $False
Switch ($Section)
{
	"Admins" {$ValidSection = $True}
	"Apps" {$ValidSection = $True}
	"ConfigLog" {$ValidSection = $True}
	"Farm" {$ValidSection = $True}
	"LoadEvals" {$ValidSection = $True}
	"Policies" {$ValidSection = $True}
	"Printers" {$ValidSection = $True}
	"Servers" {$ValidSection = $True}
	"Zones" {$ValidSection = $True}
	"All" {$ValidSection = $True}
}

If($ValidSection -eq $False)
{
	$ErrorActionPreference = $SaveEAPreference
	Write-Error -Message "`n`tThe Section parameter specified, $($Section), is an invalid Section option.`n`tValid options are:
	
	`t`tAdmins
	`t`tApps
	`t`tConfigLog
	`t`tFarm
	`t`tLoadEvals
	`t`tPolicies
	`t`tPrinters
	`t`tServers
	`t`tZones
	`t`tAll
	
	`tScript cannot continue."
	Exit
}
	
If($MSWord -or $PDF)
{
	#try and fix the issue with the $CompanyName variable
	$CoName = $CompanyName
	Write-Verbose "$(Get-Date): CoName is $($CoName)"
	
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
	[int]$wdColorRed = 255
	[int]$wdColorBlack = 0
	[int]$wdWord2007 = 12
	[int]$wdWord2010 = 14
	[int]$wdWord2013 = 15
	[int]$wdFormatDocumentDefault = 16
	[int]$wdSaveFormatPDF = 17
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

	[string]$RunningOS = (Get-WmiObject -class Win32_OperatingSystem -EA 0).Caption
}

Function GetComputerWMIInfo
{
	Param([string]$RemoteComputerName)
	
	# original work by Kees Baggerman, 
	# Senior Technical Consultant @ Inter Access
	# k.baggerman@myvirtualvision.com
	# @kbaggerman on Twitter
	# http://blog.myvirtualvision.com
	# modified 1-May-2014 to work in trusted AD Forests and using different domain admin credentials	

	#Get Computer info
	Write-Verbose "$(Get-Date): `t`tProcessing WMI Computer information"
	Write-Verbose "$(Get-Date): `t`t`tHardware information"
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Computer Information"
		WriteWordLine 0 1 "General Computer"
	}
	
	[bool]$GotComputerItems = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_computersystem
	}
	
	Catch
	{
		$Results = $Null
	}
	
	If($? -and $Results -ne $Null)
	{
		$ComputerItems = $Results | Select Manufacturer, Model, Domain, @{N="TotalPhysicalRam"; E={[math]::round(($_.TotalPhysicalMemory / 1GB),0)}}
		$Results = $Null

		ForEach($Item in $ComputerItems)
		{
			OutputComputerItem $Item
		}
	}
	ElseIf(!$?)
	{
		Write-Verbose "$(Get-Date): Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results returned for Computer information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results returned for Computer information" "" $Null 0 $False $True
		}
	}
	
	#Get Disk info
	Write-Verbose "$(Get-Date): `t`t`tDrive information"

	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 "Drive(s)"
	}

	[bool]$GotDrives = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName Win32_LogicalDisk
	}
	
	Catch
	{
		$Results = $Null
	}

	If($? -and $Results -ne $Null)
	{
		$drives = $Results | Select caption, @{N="drivesize"; E={[math]::round(($_.size / 1GB),0)}}, 
		filesystem, @{N="drivefreespace"; E={[math]::round(($_.freespace / 1GB),0)}}, 
		volumename, drivetype, volumedirty, volumeserialnumber
		$Results = $Null
		ForEach($drive in $drives)
		{
			If($drive.caption -ne "A:" -and $drive.caption -ne "B:")
			{
				OutputDriveItem $drive
			}
		}
	}
	ElseIf(!$?)
	{
		Write-Verbose "$(Get-Date): Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results returned for Drive information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results returned for Drive information" "" $Null 0 $False $True
		}
	}
	

	#Get CPU's and stepping
	Write-Verbose "$(Get-Date): `t`t`tProcessor information"

	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 "Processor(s)"
	}

	[bool]$GotProcessors = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_Processor
	}
	
	Catch
	{
		$Results = $Null
	}

	If($? -and $Results -ne $Null)
	{
		$Processors = $Results | Select availability, name, description, maxclockspeed, 
		l2cachesize, l3cachesize, numberofcores, numberoflogicalprocessors
		$Results = $Null
		ForEach($processor in $processors)
		{
			OutputProcessorItem $processor
		}
	}
	ElseIf(!$?)
	{
		Write-Verbose "$(Get-Date): Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Get-WmiObject win32_Processor failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results returned for Processor information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results returned for Processor information" "" $Null 0 $False $True
		}
	}

	#Get Nics
	Write-Verbose "$(Get-Date): `t`t`tNIC information"

	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 "Network Interface(s)"
	}

	[bool]$GotNics = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_networkadapterconfiguration
	}
	
	Catch
	{
		$Results
	}

	If($? -and $Results -ne $Null)
	{
		$Nics = $Results | Where {$_.ipaddress -ne $Null}
		$Results = $Null

		If($Nics -eq $Null ) 
		{ 
			$GotNics = $False 
		} 
		Else 
		{ 
			$GotNics = !($Nics.__PROPERTY_COUNT -eq 0) 
		} 
	
		If($GotNics)
		{
			ForEach($nic in $nics)
			{
				Try
				{
					$ThisNic = Get-WmiObject -computername $RemoteComputerName win32_networkadapter | Where {$_.index -eq $nic.index}
				}
				
				Catch 
				{
					$ThisNic = $Null
				}
				
				If($? -and $ThisNic -ne $Null)
				{
					OutputNicItem $Nic $ThisNic
				}
				ElseIf(!$?)
				{
					Write-Warning "$(Get-Date): Error retrieving NIC information"
					Write-Verbose "$(Get-Date): Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
					Write-Warning "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
					If($MSWORD -or $PDF)
					{
						WriteWordLine 0 2 "Error retrieving NIC information" "" $Null 0 $False $True
						WriteWordLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" "" $Null 0 $False $True
						WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
						WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
						WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
					}
				}
				Else
				{
					Write-Verbose "$(Get-Date): No results returned for NIC information"
					If($MSWORD -or $PDF)
					{
						WriteWordLine 0 2 "No results returned for NIC information" "" $Null 0 $False $True
					}
				}
			}
		}	
	}
	ElseIf(!$?)
	{
		Write-Warning "$(Get-Date): Error retrieving NIC configuration information"
		Write-Verbose "$(Get-Date): Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Error retrieving NIC configuration information" "" $Null 0 $False $True
			WriteWordLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results returned for NIC configuration information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results returned for NIC configuration information" "" $Null 0 $False $True
		}
	}
	
	If($MSWORD -or $PDF)
	{
		WriteWordLine 0 0 ""
	}

	$Results = $Null
	$ComputerItems = $Null
	$Drives = $Null
	$Processors = $Null
	$Nics = $Null
}

Function OutputComputerItem
{
	Param([object]$Item)
	If($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $ItemInformation = @()
		$ItemInformation += @{ Data = "Manufacturer"; Value = $Item.manufacturer; }
		$ItemInformation += @{ Data = "Model"; Value = $Item.model; }
		$ItemInformation += @{ Data = "Domain"; Value = $Item.domain; }
		$ItemInformation += @{ Data = "Total Ram"; Value = "$($Item.totalphysicalram) GB"; }
		$Table = AddWordTable -Hashtable $ItemInformation -Columns Data,Value -List -AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 125;
		$Table.Columns.Item(2).Width = 100;

		$Table.Rows.SetLeftIndent($Indent2TabStops,$wdAdjustNone)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 2 ""
		
	}
}

Function OutputDriveItem
{
	Param([object]$Drive)
	If($MSWORD -or $PDF)
	{
		[System.Collections.Hashtable[]] $DriveInformation = @()
		$DriveInformation += @{ Data = "Caption"; Value = $Drive.caption; }
		$DriveInformation += @{ Data = "Size"; Value = "$($drive.drivesize) GB"; }
		If(![String]::IsNullOrEmpty($drive.filesystem))
		{
			$DriveInformation += @{ Data = "File System"; Value = $Drive.filesystem; }
		}
		$DriveInformation += @{ Data = "Free Space"; Value = "$($drive.drivefreespace) GB"; }
		If(![String]::IsNullOrEmpty($drive.volumename))
		{
			$DriveInformation += @{ Data = "Volume Name"; Value = $Drive.volumename; }
		}
		If(![String]::IsNullOrEmpty($drive.volumedirty))
		{
			If($drive.volumedirty)
			{
				$tmp = "Yes"
			}
			Else
			{
				$tmp = "No"
			}
			$DriveInformation += @{ Data = "Volume is Dirty"; Value = $tmp; }
		}
		If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
		{
			$DriveInformation += @{ Data = "Volume Serial Number"; Value = $Drive.volumeserialnumber; }
		}
		Switch ($drive.drivetype)
		{
			0	{$tmp = "Unknown"}
			1	{$tmp = "No Root Directory"}
			2	{$tmp = "Removable Disk"}
			3	{$tmp = "Local Disk"}
			4	{$tmp = "Network Drive"}
			5	{$tmp = "Compact Disc"}
			6	{$tmp = "RAM Disk"}
			Default {$tmp = "Unknown"}
		}
		$DriveInformation += @{ Data = "Drive Type"; Value = $tmp; }
		$Table = AddWordTable -Hashtable $DriveInformation -Columns Data,Value -List -AutoFit $wdAutoFitContent;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 125;
		$Table.Columns.Item(2).Width = 100;

		$Table.Rows.SetLeftIndent($Indent2TabStops,$wdAdjustNone)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 2 ""
	}
}

Function OutputProcessorItem
{
	Param([object]$Processor)
	If($MSWORD -or $PDF)
	{
		[System.Collections.Hashtable[]] $ProcessorInformation = @()
		$ProcessorInformation += @{ Data = "Name"; Value = $Processor.name; }
		$ProcessorInformation += @{ Data = "Description"; Value = $Processor.description; }
		$ProcessorInformation += @{ Data = "Max Clock Speed"; Value = "$($processor.maxclockspeed) MHz"; }
		If($processor.l2cachesize -gt 0)
		{
			$ProcessorInformation += @{ Data = "L2 Cache Size"; Value = "$($processor.l2cachesize) KB"; }
		}
		If($processor.l3cachesize -gt 0)
		{
			$ProcessorInformation += @{ Data = "L3 Cache Size"; Value = "$($processor.l3cachesize) KB"; }
		}
		If($processor.numberofcores -gt 0)
		{
			$ProcessorInformation += @{ Data = "Number of Cores"; Value = $Processor.numberofcores; }
		}
		If($processor.numberoflogicalprocessors -gt 0)
		{
			$ProcessorInformation += @{ Data = "Number of Logical Processors"; Value = $Processor.numberoflogicalprocessors; }
		}
		Switch ($processor.availability)
		{
			1	{$tmp = "Other"}
			2	{$tmp = "Unknown"}
			3	{$tmp = "Running or Full Power"}
			4	{$tmp = "Warning"}
			5	{$tmp = "In Test"}
			6	{$tmp = "Not Applicable"}
			7	{$tmp = "Power Off"}
			8	{$tmp = "Off Line"}
			9	{$tmp = "Off Duty"}
			10	{$tmp = "Degraded"}
			11	{$tmp = "Not Installed"}
			12	{$tmp = "Install Error"}
			13	{$tmp = "Power Save - Unknown"}
			14	{$tmp = "Power Save - Low Power Mode"}
			15	{$tmp = "Power Save - Standby"}
			16	{$tmp = "Power Cycle"}
			17	{$tmp = "Power Save - Warning"}
			Default	{$tmp = "Unknown"}
		}
		$ProcessorInformation += @{ Data = "Availability"; Value = $tmp; }
		$Table = AddWordTable -Hashtable $ProcessorInformation -Columns Data,Value -List -AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent2TabStops,$wdAdjustNone)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 2 ""
	}
}

Function OutputNicItem
{
	Param([object]$Nic, [object]$ThisNic)
	If($MSWORD -or $PDF)
	{
		[System.Collections.Hashtable[]] $NicInformation = @()
		If($ThisNic.Name -eq $nic.description)
		{
			$NicInformation += @{ Data = "Name"; Value = $ThisNic.Name; }
		}
		Else
		{
			$NicInformation += @{ Data = "Name"; Value = $ThisNic.Name; }
			$NicInformation += @{ Data = "Description"; Value = $Nic.description; }
		}
		$NicInformation += @{ Data = "Connection ID"; Value = $ThisNic.NetConnectionID; }
		$NicInformation += @{ Data = "Manufacturer"; Value = $Nic.manufacturer; }
		Switch ($ThisNic.availability)
		{
			1	{$tmp = "Other"}
			2	{$tmp = "Unknown"}
			3	{$tmp = "Running or Full Power"}
			4	{$tmp = "Warning"}
			5	{$tmp = "In Test"}
			6	{$tmp = "Not Applicable"}
			7	{$tmp = "Power Off"}
			8	{$tmp = "Off Line"}
			9	{$tmp = "Off Duty"}
			10	{$tmp = "Degraded"}
			11	{$tmp = "Not Installed"}
			12	{$tmp = "Install Error"}
			13	{$tmp = "Power Save - Unknown"}
			14	{$tmp = "Power Save - Low Power Mode"}
			15	{$tmp = "Power Save - Standby"}
			16	{$tmp = "Power Cycle"}
			17	{$tmp = "Power Save - Warning"}
			Default	{$tmp = "Unknown"}
		}
		$NicInformation += @{ Data = "Availability"; Value = $tmp; }
		$NicInformation += @{ Data = "Physical Address"; Value = $Nic.macaddress; }
		$NicInformation += @{ Data = "IP Address"; Value = $Nic.ipaddress; }
		$NicInformation += @{ Data = "Default Gateway"; Value = $Nic.Defaultipgateway; }
		$NicInformation += @{ Data = "Subnet Mask"; Value = $Nic.ipsubnet; }
		If($nic.dhcpenabled)
		{
			$DHCPLeaseObtainedDate = $nic.ConvertToDateTime($nic.dhcpleaseobtained)
			$DHCPLeaseExpiresDate = $nic.ConvertToDateTime($nic.dhcpleaseexpires)
			$NicInformation += @{ Data = "DHCP Enabled"; Value = $Nic.dhcpenabled; }
			$NicInformation += @{ Data = "DHCP Lease Obtained"; Value = $dhcpleaseobtaineddate; }
			$NicInformation += @{ Data = "DHCP Lease Expires"; Value = $dhcpleaseexpiresdate; }
			$NicInformation += @{ Data = "DHCP Server"; Value = $Nic.dhcpserver; }
		}
		If(![String]::IsNullOrEmpty($nic.dnsdomain))
		{
			$NicInformation += @{ Data = "DNS Domain"; Value = $Nic.dnsdomain; }
		}
		If($nic.dnsdomainsuffixsearchorder -ne $Null -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
		{
			[int]$x = 1
			WriteWordLine 0 2 "DNS Search Suffixes`t:" -nonewline
			$nicdnsdomainsuffixsearchorder = $nic.dnsdomainsuffixsearchorder
			$tmp = @()
			ForEach($DNSDomain in $nicdnsdomainsuffixsearchorder)
			{
				$tmp += "$($DNSDomain)`r"
			}
			$NicInformation += @{ Data = "DNS Search Suffixes"; Value = $tmp; }
		}
		If($nic.dnsenabledforwinsresolution)
		{
			$tmp = "Yes"
		}
		Else
		{
			$tmp = "No"
		}
		$NicInformation += @{ Data = "DNS WINS Enabled"; Value = $tmp; }
		If($nic.dnsserversearchorder -ne $Null -and $nic.dnsserversearchorder.length -gt 0)
		{
			$nicdnsserversearchorder = $nic.dnsserversearchorder
			$tmp = @()
			ForEach($DNSServer in $nicdnsserversearchorder)
			{
				$tmp += "$($DNSServer)`r"
			}
			$NicInformation += @{ Data = "DNS Servers"; Value = $tmp; }
		}
		Switch ($nic.TcpipNetbiosOptions)
		{
			0	{$tmp = "Use NetBIOS setting from DHCP Server"}
			1	{$tmp = "Enable NetBIOS"}
			2	{$tmp = "Disable NetBIOS"}
			Default	{$tmp = "Unknown"}
		}
		$NicInformation += @{ Data = "NetBIOS Setting"; Value = $tmp; }
		If($nic.winsenablelmhostslookup)
		{
			$tmp = "Yes"
		}
		Else
		{
			$tmp = "No"
		}
		$NicInformation += @{ Data = "WINS: Enabled LMHosts"; Value = $tmp; }
		If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
		{
			$NicInformation += @{ Data = "Host Lookup File"; Value = $Nic.winshostlookupfile; }
		}
		If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
		{
			$NicInformation += @{ Data = "Primary Server"; Value = $Nic.winsprimaryserver; }
		}
		If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
		{
			$NicInformation += @{ Data = "Secondary Server"; Value = $Nic.winssecondaryserver; }
		}
		If(![String]::IsNullOrEmpty($nic.winsscopeid))
		{
			$NicInformation += @{ Data = "Scope ID"; Value = $Nic.winsscopeid; }
		}
		$Table = AddWordTable -Hashtable $NicInformation -Columns Data,Value -List -AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent2TabStops,$wdAdjustNone)

		FindWordDocumentEnd
		$Table = $Null
	}
}

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

	[string]$toc = $(
		Switch ($CultureCode)
		{
			'ca-'	{ 'Taula automática 2' }

			'da-'	{ 'Automatisk tabel 2' }

			'de-'	{ 'Automatische Tabelle 2' }

			'en-'	{ 'Automatic Table 2' }

			'es-'	{ 'Tabla automática 2' }

			'fi-'	{ 'Automaattinen taulukko 2' }

			'fr-'	{ 'Sommaire Automatique 2' }

			'nb-'	{ 'Automatisk tabell 2' }

			'nl-'	{ 'Automatische inhoudsopgave 2' }

			'pt-'	{ 'Sumário Automático 2' }

			'sv-'	{ 'Automatisk innehållsförteckning2' }
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

	Switch ($WordValue)
	{
		{$CatalanArray -contains $_} {$CultureCode = "ca-"}
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
				If($xWordVersion -eq $wdWord2013)
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
				If($xWordVersion -eq $wdWord2013)
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
				If($xWordVersion -eq $wdWord2013)
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
				If($xWordVersion -eq $wdWord2013)
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
				If($xWordVersion -eq $wdWord2013)
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
				If($xWordVersion -eq $wdWord2013)
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
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("ViewMaster", "Secteur (foncé)", "Sémaphore",
					"Rétrospective", "Ion (foncé)", "Ion (clair)", "Intégrale",
					"Filigrane", "Facette", "Secteur (clair)", "À bandes", "Austin",
					"Guide", "Whisp", "Lignes latérales", "Quadrillage")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Mosaïques", "Ligne latérale", "Annuel", "Perspective",
					"Contraste", "Emplacements de bureau", "Moderne", "Blocs empilés",
					"Rayures fines", "Austère", "Transcendant", "Classique", "Quadrillage",
					"Exposition", "Alphabet", "Mots croisés", "Papier journal", "Austin", "Guide")
				}
			}

		'nb-'	{
				If($xWordVersion -eq $wdWord2013)
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
				If($xWordVersion -eq $wdWord2013)
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
				If($xWordVersion -eq $wdWord2013)
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
				If($xWordVersion -eq $wdWord2013)
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

		Default	{
					If($xWordVersion -eq $wdWord2013)
					{
						$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid", "Integral",
						"Ion (Dark)", "Ion (Light)", "Motion", "Retrospect", "Semaphore",
						"Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster", "Whisp")
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

Function ConvertTo-ScriptBlock 
{
	#by Jeff Wouters, PowerShell MVP
	Param([string]$string)
	$ScriptBlock = $executioncontext.invokecommand.NewScriptBlock($string)
	Return $ScriptBlock
}

Function SWExclusions 
{
	# original work by Shaun Ritchie
	# performance improvements by Jeff Wouters, PowerShell MVP
	# modified by Webster
	# modified 3-jan-2014 to add displayversion
	# bug found 30-jul-2014 by Sam Jacobs
	# this function did not work if the SoftwareExlusions.txt file contained only one line
	$var = ""
	$Tmp = '$InstalledApps | Where {'
	$Exclusions = Get-Content "$($pwd.path)\SoftwareExclusions.txt" -EA 0
	If($? -and $Exclusions -ne $Null)
	{
		If($Exclusions -is [array])	
		{
			ForEach($Exclusion in $Exclusions) 
			{
				$Tmp += "(`$`_.DisplayName -notlike ""$($Exclusion)"") -and "
			}
			$var += $Tmp.Substring(0,($Tmp.Length - 6))
			}
		Else
		{
			# added 30-jul-2014 to handle if the file contained only one line
			$tmp += "(`$`_.DisplayName -notlike ""$($Exclusions)"")"
			$var = $tmp
		}
		$var += "} | Select-Object DisplayName, DisplayVersion | Sort DisplayName -unique"
	}
	return $var
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
	[bool]$wordrunning = ((Get-Process 'WinWord' -ea 0)|?{$_.SessionId -eq $SessionID}) -ne $Null
	If($wordrunning)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Host "`n`n`tPlease close all instances of Microsoft Word before running this report.`n`n"
		Exit
	}
}

Function ValidateCompanyName
{
	$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	If($xResult)
	{
		Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	}
	Else
	{
		$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		If($xResult)
		{
			Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		}
		Else
		{
			Return ""
		}
	}
}

#http://stackoverflow.com/questions/5648931/test-if-registry-value-exists
# This Function just gets $True or $False
Function Test-RegistryValue($path, $name)
{
    $key = Get-Item -LiteralPath $path -EA 0
    $key -and $Null -ne $key.GetValue($name, $Null)
}

# Gets the specified registry value or $Null if it is missing
Function Get-RegistryValue($path, $name)
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

Function Check-NeededPSSnapins
{
	Param([parameter(Mandatory = $True)][alias("Snapin")][string[]]$Snapins)

	#Function specifics
	$MissingSnapins = @()
	[bool]$FoundMissingSnapin = $False
	$LoadedSnapins = @()
	$RegisteredSnapins = @()

	#Creates arrays of strings, rather than objects, we're passing strings so this will be more robust.
	$loadedSnapins += get-pssnapin | % {$_.name}
	$registeredSnapins += get-pssnapin -Registered | % {$_.name}

	ForEach($Snapin in $Snapins)
	{
		#check if the snapin is loaded
		If(!($LoadedSnapins -like $snapin))
		{
			#Check if the snapin is missing
			If(!($RegisteredSnapins -like $Snapin))
			{
				#set the flag if it's not already
				If(!($FoundMissingSnapin))
				{
					$FoundMissingSnapin = $True
				}
				#add the entry to the list
				$MissingSnapins += $Snapin
			}
			Else
			{
				#Snapin is registered, but not loaded, loading it now:
				Write-Host "Loading Windows PowerShell snap-in: $snapin"
				Add-PSSnapin -Name $snapin -EA 0
			}
		}
	}

	If($FoundMissingSnapin)
	{
		Write-Warning "Missing Windows PowerShell snap-ins Detected:"
		$missingSnapins | % {Write-Warning "($_)"}
		return $False
	}
	Else
	{
		Return $True
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
		0 {$Selection.Style = $myHash.Word_NoSpacing}
		1 {$Selection.Style = $myHash.Word_Heading1}
		2 {$Selection.Style = $myHash.Word_Heading2}
		3 {$Selection.Style = $myHash.Word_Heading3}
		4 {$Selection.Style = $myHash.Word_Heading4}
		Default {$Selection.Style = $myHash.Word_NoSpacing}
	}
	
	#build # of tabs
	While($tabs -gt 0)
	{ 
		$output += "`t"; $tabs--; 
	}
 
	If(![String]::IsNullOrEmpty($fontName)) 
	{
		$Selection.Font.name = $fontName
	} 

	If($fontSize -ne 0) 
	{
		$Selection.Font.size = $fontSize
	} 
 
	If($italics -eq $True) 
	{
		$Selection.Font.Italic = $True
	} 
 
	If($boldface -eq $True) 
	{
		$Selection.Font.Bold = $True
	} 

	#output the rest of the parameters.
	$output += $name + $value
	$Selection.TypeText($output)
 
	#test for new WriteWordLine 0.
	If($nonewline)
	{
		# Do nothing.
	} 
	Else 
	{
		$Selection.TypeParagraph()
	}
}

Function _SetDocumentProperty 
{
	#jeff hicks
	Param([object]$Properties,[string]$Name,[string]$Value)
	#get the property object
	$prop=$properties | ForEach { 
		$propname=$_.GetType().InvokeMember("Name","GetProperty",$Null,$_,$Null)
		If($propname -eq $Name) 
		{
			Return $_
		}
	} #ForEach

	#set the value
	$Prop.GetType().InvokeMember("Value","SetProperty",$Null,$prop,$Value)
}

Function FindWordDocumentEnd
{
	#return focus to main document    
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument
	#move to the end of the current document
	$Script:Selection.EndKey($wdStory,$wdMove) | Out-Null
}

Function Process2008Policies
{
	#Bandwidth
	$xArray = ($Setting.TurnWallpaperOffState, $Setting.TurnWindowContentsOffState, $Setting.TurnWindowContentsOffState)
	If($xArray -contains "Enabled" -or $xArray -contains "Disabled")
	{
		Write-Verbose "$(Get-Date): `t`t`tBandwidth\Visual Effects\"
		WriteWordLine 0 2 "Bandwidth\Visual Effects\"
		If($Setting.TurnWallpaperOffState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Turn Off Desktop Wallpaper: " $Setting.TurnWallpaperOffState
		}
		If($Setting.TurnMenuAnimationsOffState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Turn Off Menu and Windows Animations: " $Setting.TurnMenuAnimationsOffState
		}
		If($Setting.TurnWindowContentsOffState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Turn Off Window Contents While Dragging: " $Setting.TurnWindowContentsOffState
		}
	}
	
	If($Setting.ImageAccelerationState -ne "NotConfigured")
	{
		Write-Verbose "$(Get-Date): `t`t`tBandwidth\SpeedScreen\"
		WriteWordLine 0 2 "Bandwidth\SpeedScreen\"
		WriteWordLine 0 3 "Image acceleration using lossy compression: " $Setting.ImageAccelerationState
		If($Setting.ImageAccelerationState -eq "Enabled")
		{
			WriteWordLine 0 3 "Compression level: " -nonewline
			
			Switch ($Setting.ImageAccelerationCompressionLevel)
			{
				"HighCompression"   {WriteWordLine 0 0 "High compression; lower image quality"}
				"MediumCompression" {WriteWordLine 0 0 "Medium compression; good image quality"}
				"LowCompression"    {WriteWordLine 0 0 "Low compression; best image quality"}
				"NoCompression"     {WriteWordLine 0 0 "Do not use lossy compression"}
				Default {WriteWordLine 0 0 "Compression level could not be determined: $($Setting.ImageAccelerationCompressionLevel)"}
			}
			If($Setting.ImageAccelerationCompressionIsRestricted)
			{
				WriteWordLine 0 3 "Restrict compression to connections under this "
				WriteWordLine 0 4 "bandwidth\Threshold (Kb/sec): " $Setting.ImageAccelerationCompressionLimit	
			}
			WriteWordLine 0 3 "SpeedScreen Progressive Display compression level: "
			Switch ($Setting.ImageAccelerationProgressiveLevel)
			{
				"UltrahighCompression" {WriteWordLine 0 4 "Ultra high compression; ultra low quality"}
				"VeryHighCompression"  {WriteWordLine 0 4 "Very high compression; very low quality"}
				"HighCompression"      {WriteWordLine 0 4 "High compression; low quality"}
				"MediumCompression"    {WriteWordLine 0 4 "Medium compression; medium quality"}
				"LowCompression"       {WriteWordLine 0 4 "Low compression; high quality"}
				"Disabled"             {WriteWordLine 0 4 "Disabled; no progressive display"}
				Default {WriteWordLine 0 0 "SpeedScreen Progressive Display compression level could not be determined: $($Setting.ImageAccelerationProgressiveLevel)"}
			}
			If($Setting.ImageAccelerationProgressiveIsRestricted)
			{
				WriteWordLine 0 3 "Restrict compression to connections under this "
				WriteWordLine 0 4 "bandwidth\Threshold (Kb/sec): " $Setting.ImageAccelerationProgressiveLimit	
			}
			WriteWordLine 0 3 "Use Heavyweight compression (extra CPU, retains quality): " -nonewline
			If($Setting.ImageAccelerationIsHeavyweightUsed)
			{
				WriteWordLine 0 0 "Yes"
			}
			Else
			{
				WriteWordLine 0 0 "No"
			}
		}
	}
	
	$xArray = (	$Setting.SessionAudioState,	$Setting.SessionClipboardState,		$Setting.SessionComportsState, 
			$Setting.SessionDrivesState,	$Setting.SessionLptPortsState,		$Setting.SessionOemChannelsState, 
			$Setting.SessionOverallState,	$Setting.SessionPrinterBandwidthState,	$Setting.SessionTwainRedirectionState)
	If($xArray -contains "Enabled" -or $xArray -contains "Disabled")
	{
		Write-Verbose "$(Get-Date): `t`t`tBandwidth\Session Limits\"
		WriteWordLine 0 2 "Bandwidth\Session Limits\"
		If($Setting.SessionAudioState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Audio: " $Setting.SessionAudioState
			If($Setting.SessionAudioState -eq "Enabled")
			{
				WriteWordLine 0 4 "Limit (Kb/sec): " $Setting.SessionAudioLimit
			}
		}
		If($Setting.SessionClipboardState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Clipboard: " $Setting.SessionClipboardState
			If($Setting.SessionClipboardState -eq "Enabled")
			{
				WriteWordLine 0 4 "Limit (Kb/sec): " $Setting.SessionClipboardLimit
			}
		}
		If($Setting.SessionComportsState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "COM Ports: " $Setting.SessionComportsState
			If($Setting.SessionComportsState -eq "Enabled")
			{
				WriteWordLine 0 4 "Limit (Kb/sec): " $Setting.SessionComportsLimit
			}
		}
		If($Setting.SessionDrivesState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Drives: " $Setting.SessionDrivesState
			If($Setting.SessionDrivesState -eq "Enabled")
			{
				WriteWordLine 0 4 "Limit (Kb/sec): " $Setting.SessionDrivesLimit
			}
		}
		If($Setting.SessionLptPortsState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "LPT Ports: " $Setting.SessionLptPortsState
			If($Setting.SessionLptPortsState -eq "Enabled")
			{
				WriteWordLine 0 4 "Limit (Kb/sec): " $Setting.SessionLptPortsLimit
			}
		}
		If($Setting.SessionOemChannelsState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "OEM Virtual Channels: " $Setting.SessionOemChannelsState
			If($Setting.SessionOemChannelsState -eq "Enabled")
			{
				WriteWordLine 0 4 "Limit (Kb/sec): " $Setting.SessionOemChannelsLimit
			}
		}
		If($Setting.SessionOverallState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Overall Session: " $Setting.SessionOverallState
			If($Setting.SessionOverallState -eq "Enabled")
			{
				WriteWordLine 0 4 "Limit (Kb/sec): " $Setting.SessionOverallLimit
			}
		}
		If($Setting.SessionPrinterBandwidthState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Printer: " $Setting.SessionPrinterBandwidthState
			If($Setting.SessionPrinterBandwidthState -eq "Enabled")
			{
				WriteWordLine 0 4 "Limit (Kb/sec): " $Setting.SessionPrinterBandwidthLimit
			}
		}
		If($Setting.SessionTwainRedirectionState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "TWAIN Redirection: " $Setting.SessionTwainRedirectionState
			If($Setting.SessionTwainRedirectionState -eq "Enabled")
			{
				WriteWordLine 0 4 "Limit (Kb/sec): " $Setting.SessionTwainRedirectionLimit
			}
		}
	}

	$xArray = (	$Setting.SessionAudioPercentState,	$Setting.SessionClipboardPercentState,	$Setting.SessionComportsPercentState, 
			$Setting.SessionDrivesPercentState,	$Setting.SessionLptPortsPercentState,	$Setting.SessionOemChannelsPercentState, 
			$Setting.SessionOverallState,		$Setting.SessionPrinterPercentState,	$Setting.SessionTwainRedirectionPercentState)
	If($xArray -contains "Enabled" -or $xArray -contains "Disabled")
	{
		Write-Verbose "$(Get-Date): `t`t`tBandwidth\Session Limits (%)\"
		WriteWordLine 0 2 'Bandwidth\Session Limits (%)\'
		If($Setting.SessionAudioPercentState -ne "NotConfigured")
		{
			WriteWordLine 0 3 'Audio: ' $Setting.SessionAudioPercentState
			If($Setting.SessionAudioPercentState -eq "Enabled")
			{
				WriteWordLine 0 4 'Limit (%): ' $Setting.SessionAudioPercentLimit
			}
		}
		If($Setting.SessionClipboardPercentState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Clipboard: " $Setting.SessionClipboardPercentState
			If($Setting.SessionClipboardPercentState -eq "Enabled")
			{
				WriteWordLine 0 4 'Limit (%): ' $Setting.SessionClipboardPercentLimit
			}
		}
		If($Setting.SessionComportsPercentState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "COM Ports: " $Setting.SessionComportsPercentState
			If($Setting.SessionComportsPercentState -eq "Enabled")
			{
				WriteWordLine 0 4 'Limit (%): ' $Setting.SessionComportsPercentLimit
			}
		}
		If($Setting.SessionDrivesPercentState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Drives: " $Setting.SessionDrivesPercentState
			If($Setting.SessionDrivesPercentState -eq "Enabled")
			{
				WriteWordLine 0 4 'Limit (%): ' $Setting.SessionDrivesPercentLimit
			}
		}
		If($Setting.SessionLptPortsPercentState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "LPT Ports: " $Setting.SessionLptPortsPercentState
			If($Setting.SessionLptPortsPercentState -eq "Enabled")
			{
				WriteWordLine 0 4 'Limit (%): ' $Setting.SessionLptPortsPercentLimit
			}
		}
		If($Setting.SessionOemChannelsPercentState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "OEM Virtual Channels: " $Setting.SessionOemChannelsPercentState
			If($Setting.SessionOemChannelsPercentState -eq "Enabled")
			{
				WriteWordLine 0 4 'Limit (%): ' $Setting.SessionOemChannelsPercentLimit
			}
		}
		If($Setting.SessionPrinterPercentState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Printer: " $Setting.SessionPrinterPercentState
			If($Setting.SessionPrinterPercentState -eq "Enabled")
			{
				WriteWordLine 0 4 'Limit (%): ' $Setting.SessionPrinterPercentLimit
			}
		}
		If($Setting.SessionTwainRedirectionPercentState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "TWAIN Redirection: " $Setting.SessionTwainRedirectionPercentState
			If($Setting.SessionTwainRedirectionPercentState -eq "Enabled")
			{
				WriteWordLine 0 4 'Limit (%): ' $Setting.SessionTwainRedirectionPercentLimit
			}
		}
	}
	
	$xArray = (	$Setting.ClientMicrophonesState,	$Setting.ClientSoundQualityState,		$Setting.TurnClientAudioMappingOffState,
			$Setting.ClientDrivesState,		$Setting.ClientDriveMappingState,		$Setting.ClientAsynchronousWritesState,
			$Setting.TwainRedirectionState,	$Setting.TurnClipboardMappingOffState,	$Setting.TurnOemVirtualChannelsOffState,
			$Setting.TurnComPortsOffState,	$Setting.TurnLptPortsOffState,		$Setting.TurnVirtualComPortMappingOffState)
	If($xArray -contains "Enabled" -or $xArray -contains "Disabled")
	{
		Write-Verbose "$(Get-Date): `t`t`tClient Devices\Resources"
		WriteWordLine 0 2 "Client Devices\Resources"
		If($Setting.ClientMicrophonesState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Audio\Microphones: " $Setting.ClientMicrophonesState
			If($Setting.ClientMicrophonesState -eq "Enabled")
			{
				If($Setting.ClientMicrophonesAreUsed)
				{
					WriteWordLine 0 4 "Use client microphones for audio input"
				}
				Else
				{
					WriteWordLine 0 4 "Do not use client microphones for audio input"
				}
			}
		}
		If($Setting.ClientSoundQualityState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Audio\Sound quality: " $Setting.ClientSoundQualityState
			If($Setting.ClientSoundQualityState)
			{
				WriteWordLine 0 4 "Maximum allowable client audio quality: " 
				Switch ($Setting.ClientSoundQualityLevel)
				{
					"Medium" {WriteWordLine 0 5 "Optimized for Speech"}
					"Low"    {WriteWordLine 0 5 "Low Bandwidth"}
					"High"   {WriteWordLine 0 5 "High Definition"}
					Default {WriteWordLine 0 0 "Maximum allowable client audio quality could not be determined: $($Setting.ClientSoundQualityLevel)"}
				}
			}
		}
		If($Setting.TurnClientAudioMappingOffState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Audio\Turn off speakers: " $Setting.TurnClientAudioMappingOffState
			If($Setting.TurnClientAudioMappingOffState -eq "Enabled")
			{
				WriteWordLine 0 4 "Turn off audio mapping to client speakers"
			}
		}
		
		Write-Verbose "$(Get-Date): `t`t`tClient Devices\Resources\Drives"
		WriteWordLine 0 2 "Client Devices\Resources\Drives"
		If($Setting.ClientDrivesState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Connection: " $Setting.ClientDrivesState
			If($Setting.ClientDrivesState -eq "Enabled")
			{
				If($Setting.ClientDrivesAreConnected)
				{
					WriteWordLine 0 4 "Connect Client Drives at Logon"
				}
				Else
				{
					WriteWordLine 0 4 "Do Not Connect Client Drives at Logon"
				}
			}
		}
		If($Setting.ClientDriveMappingState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Mappings: " $Setting.ClientDriveMappingState
			If($Setting.ClientDriveMappingState -eq "Enabled")
			{
				If($Setting.TurnFloppyDriveMappingOff)
				{
					WriteWordLine 0 4 "Turn off Floppy disk drives"	
				}
				If($Setting.TurnHardDriveMappingOff)
				{
					WriteWordLine 0 4 "Turn off Hard drives"	
				}
				If($Setting.TurnCDRomDriveMappingOff)
				{
					WriteWordLine 0 4 "Turn off CD-ROM drives"	
				}
				If($Setting.TurnRemoteDriveMappingOff)
				{
					WriteWordLine 0 4 "Turn off Remote drives"	
				}
			}
		}
		If($Setting.ClientAsynchronousWritesState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Optimize\Asynchronous writes: " $Setting.ClientAsynchronousWritesState
			If($Setting.ClientAsynchronousWritesState -eq "Enabled")
			{
				WriteWordLine 0 4 "Turn on asynchronous disk writes to client disks"
			}

			WriteWordLine 0 3 "Special folder redirection: " $Setting.TurnSpecialFolderRedirectionOffState
			If($Setting.TurnSpecialFolderRedirectionOffState -eq "Enabled")
			{
				WriteWordLine 0 4 "Do not allow special folder redirection"
			}
		}

		$xArray = ($Setting.TwainRedirectionState, $Setting.TurnClipboardMappingOffState, $Setting.TurnOemVirtualChannelsOffState)
		If($xArray -contains "Enabled" -or $xArray -contains "Disabled")
		{
			Write-Verbose "$(Get-Date): `t`t`tClient Devices\Resources\Other"
			WriteWordLine 0 2 "Client Devices\Resources\Other"
			If($Setting.TwainRedirectionState -ne "NotConfigured")
			{
				WriteWordLine 0 3 "Configure TWAIN redirection: " $Setting.TwainRedirectionState
				If($Setting.TwainRedirectionState -eq "Enabled")
				{
					If($Setting.TwainRedirectionAllowed)
					{
						WriteWordLine 0 4 "Allow TWAIN redirection"
						If($Setting.TwainRedirectionImageCompression -eq "NoCompression")
						{
							WriteWordLine 0 4 "Do not use lossy compression for high color images"
						}
						Else
						{
							WriteWordLine 0 4 "Use lossy compression for high color images: "
							
							Switch ($Setting.TwainRedirectionImageCompression)
							{
								"HighCompression"   {WriteWordLine 0 5 "High compression; lower image quality"}
								"MediumCompression" {WriteWordLine 0 5 "Medium compression; good image quality"}
								"LowCompression"    {WriteWordLine 0 5 "Low compression; best image quality"}
								Default {WriteWordLine 0 0 "Lossy compression for high color images could not be determined: $($Setting.TwainRedirectionImageCompression)"}
							}
						}
					}
					Else
					{
						WriteWordLine 0 4 "Do not allow TWAIN redirection"
					}
				}
			}
			If($Setting.TurnClipboardMappingOffState -ne "NotConfigured")
			{
				WriteWordLine 0 3 "Turn off clipboard mapping: " $Setting.TurnClipboardMappingOffState
			}
			If($Setting.TurnOemVirtualChannelsOffState -ne "NotConfigured")
			{
				WriteWordLine 0 3 "Turn off OEM virtual channels: " $Setting.TurnOemVirtualChannelsOffState
			}
		}

		$xArray = ($Setting.TurnComPortsOffState, $Setting.TurnLptPortsOffState)
		If($xArray -contains "Enabled" -or $xArray -contains "Disabled")
		{
			Write-Verbose "$(Get-Date): `t`t`tClient Devices\Resources\Ports"
			WriteWordLine 0 2 "Client Devices\Resources\Ports"
			If($Setting.TurnComPortsOffState -ne "NotConfigured")
			{
				WriteWordLine 0 3 "Turn off COM ports: " $Setting.TurnComPortsOffState
			}
			If($Setting.TurnLptPortsOffState -ne "NotConfigured")
			{
				WriteWordLine 0 3 "Turn off LPT ports: " $Setting.TurnLptPortsOffState
			}
		}
		
		If($Setting.TurnVirtualComPortMappingOffState -ne "NotConfigured")
		{
			Write-Verbose "$(Get-Date): `t`t`tClient Devices\Resources\PDA Devices"
			WriteWordLine 0 2 "Client Devices\Resources\PDA Devices"
			WriteWordLine 0 3 "Turn on automatic virtual COM port mapping: " $Setting.TurnVirtualComPortMappingOffState
		}
	}
	
	If($Setting.TurnAutoClientUpdateOffState -ne "NotConfigured")
	{
		Write-Verbose "$(Get-Date): `t`t`t`Client Devices\Maintenance"
		WriteWordLine 0 2 "Client Devices\Maintenance"
		WriteWordLine 0 3 "Turn off auto client update: " $Setting.TurnAutoClientUpdateOffState
	}
	
	$xArray = (	$Setting.ClientPrinterAutoCreationState,	$Setting.LegacyClientPrintersState,
			$Setting.PrinterPropertiesRetentionState,	$Setting.PrinterJobRoutingState,
			$Setting.TurnClientPrinterMappingOffState)
	If($xArray -contains "Enabled" -or $xArray -contains "Disabled")
	{
		Write-Verbose "$(Get-Date): `t`t`tPrinting\Client Printers"
		WriteWordLine 0 2 "Printing\Client Printers"
		If($Setting.ClientPrinterAutoCreationState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Auto-creation: " $Setting.ClientPrinterAutoCreationState
			If($Setting.ClientPrinterAutoCreationState -eq "Enabled")
			{
				WriteWordLine 0 4 "When connecting:"
				Switch ($Setting.ClientPrinterAutoCreationOption)
				{
					"LocalPrintersOnly"  {WriteWordLine 0 5 "Auto-create local (non-network) client printers only"}
					"AllPrinters"        {WriteWordLine 0 5 "Auto-create all client printers"}
					"DefaultPrinterOnly" {WriteWordLine 0 5 "Auto-create the client's Default printer only"}
					"DoNotAutoCreate"    {WriteWordLine 0 5 "Do not auto-create client printers"}
					Default {WriteWordLine 0 0 "Client Printers\Auto-creation could not be determined: $($Setting.ClientPrinterAutoCreationOption)"}
				}
			}
		}

		If($Setting.LegacyClientPrintersState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Legacy client printers: " $Setting.LegacyClientPrintersState
			If($Setting.LegacyClientPrintersState -eq "Enabled")
			{
				If($Setting.LegacyClientPrintersDynamic)
				{
					WriteWordLine 0 4 "Create dynamic session-private client printers"
				}
				Else
				{
					WriteWordLine 0 4 "Create old-style client printers"
				}
			}
		}
		If($Setting.PrinterPropertiesRetentionState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Printer properties retention: " $Setting.PrinterPropertiesRetentionState
			If($Setting.PrinterPropertiesRetentionState -eq "Enabled")
			{
				WriteWordLine 0 4 "Printer properties " -nonewline
				
				Switch ($Setting.PrinterPropertiesRetentionOption)
				{
					"FallbackToProfile"     {WriteWordLine 0 0 "Held in profile only if not saved on client"}
					"RetainedInUserProfile" {WriteWordLine 0 0 "Retained in user profile only"}
					"SavedOnClientDevice"   {WriteWordLine 0 0 "Saved on the client device only"}
					Default {WriteWordLine 0 0 "Printer properties retention could not be determined: $($Setting.PrinterPropertiesRetentionOption)"}
				}
			}
		}
		If($Setting.PrinterJobRoutingState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Print job routing: " $Setting.PrinterJobRoutingState
			If($Setting.PrinterJobRoutingState -eq "Enabled")
			{
				WriteWordLine 0 4 "For client printers on a network printer server: "
				If($Setting.PrinterJobRoutingDirect)
				{
					WriteWordLine 0 5 "Connect directly to network print server if possible"
				}
				Else
				{
					WriteWordLine 0 5 "Always connect indirectly as a client printer"
				}
			}
		}
		If($Setting.TurnClientPrinterMappingOffState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Turn off client printer mapping: " $Setting.TurnClientPrinterMappingOffState
		}
	}
	
	$xArray = ($Setting.DriverAutoInstallState, $Setting.UniversalDriverState)
	If($xArray -contains "Enabled" -or $xArray -contains "Disabled")
	{
		Write-Verbose "$(Get-Date): `t`t`tPrinting\Drivers"
		WriteWordLine 0 2 "Printing\Drivers"
		If($Setting.DriverAutoInstallState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Native printer driver auto-install: " $Setting.DriverAutoInstallState
			If($Setting.DriverAutoInstallState -eq "Enabled")
			{
				If($Setting.DriverAutoInstallAsNeeded)
				{
					WriteWordLine 0 4 "Install Windows native drivers as needed"
				}
				Else
				{
					WriteWordLine 0 4 "Do not automatically install drivers"
				}
			}
		}
		If($Setting.UniversalDriverState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Universal driver: " $Setting.UniversalDriverState
			If($Setting.UniversalDriverState -eq "Enabled")
			{
				WriteWordLine 0 4 "When auto-creating client printers: "
				
				Switch ($Setting.UniversalDriverOption)
				{
					"FallbackOnly"  {WriteWordLine 0 4 "Use universal driver only if requested driver is unavailable"}
					"SpecificOnly"  {WriteWordLine 0 4 "Use only printer model specific drivers"}
					"ExclusiveOnly" {WriteWordLine 0 4 "Use universal driver only"}
					Default {WriteWordLine 0 0 "When auto-creating client printers could not be determined: $($Setting.UniversalDriverOption)"}
				}
			}
		}
	}
	
	If($Setting.SessionPrintersState -ne "NotConfigured")
	{
		Write-Verbose "$(Get-Date): `t`t`tPrinting\Session printers"
		WriteWordLine 0 2 "Printing\Session printers"
		WriteWordLine 0 3 "Session printers: " $Setting.SessionPrintersState
		If($Setting.SessionPrintersState -eq "Enabled")
		{
			If($Setting.SessionPrinterList)
			{
				WriteWordLine 0 4 "Network printers to connect at logon:"
				ForEach($Printer in $Setting.SessionPrinterList)
				{
					WriteWordLine 0 5 $Printer
					$index = $Printer.SubString(2).IndexOf('\')
					If($index -ge 0)
					{
						$srv = $Printer.SubString(0, $index + 2)
						$ptr  = $Printer.SubString($index + 3)
					}
					$SessionPrinterSettings = Get-XASessionPrinter -PolicyName $Policy.PolicyName -PrinterName $Ptr -EA 0
					If(![String]::IsNullOrEmpty($SessionPrinterSettings))
					{
						If($SessionPrinterSettings.ApplyCustomSettings)
						{
							WriteWordLine 0 5 "Shared Name`t: " $SessionPrinterSettings.PrinterName
							WriteWordLine 0 5 "Server`t`t: " $SessionPrinterSettings.ServerName
							WriteWordLine 0 5 "Printer Model`t: " $SessionPrinterSettings.DriverName
							If(![String]::IsNullOrEmpty($SessionPrinterSettings.Location))
							{
								WriteWordLine 0 5 "Location`t: " $SessionPrinterSettings.Location
							}
							WriteWordLine 0 5 "Paper Size`t: " -nonewline
							Switch ($SessionPrinterSettings.PaperSize)
							{
								"A4"          {WriteWordLine 0 0 "A4"}
								"A4Small"     {WriteWordLine 0 0 "A4 Small"}
								"Envelope10"  {WriteWordLine 0 0 "Envelope #10"}
								"EnvelopeB5"  {WriteWordLine 0 0 "Envelope B5"}
								"EnvelopeC5"  {WriteWordLine 0 0 "Envelope C5"}
								"EnvelopeDL"  {WriteWordLine 0 0 "Envelope DL"}
								"Monarch"     {WriteWordLine 0 0 "Envelope Monarch"}
								"Executive"   {WriteWordLine 0 0 "Executive"}
								"Legal"       {WriteWordLine 0 0 "Legal"}
								"Letter"      {WriteWordLine 0 0 "Letter"}
								"LetterSmall" {WriteWordLine 0 0 "Letter Small"}
								"Note" {WriteWordLine 0 0 "Note"}
								Default 
								{
									WriteWordLine 0 0 "Custom..."
									WriteWordLine 0 5 "Width`t`t: $($SessionPrinterSettings.Width) (Millimeters)" 
									WriteWordLine 0 5 "Height`t`t: $($SessionPrinterSettings.Height) (Millimeters)" 
								}
							}
							WriteWordLine 0 5 "Copy Count`t: " $SessionPrinterSettings.CopyCount
							If($SessionPrinterSettings.CopyCount -gt 1)
							{
								WriteWordLine 0 5 "Collated`t: " -nonewline
								If($SessionPrinterSettings.Collated)
								{
									WriteWordLine 0 0 "Yes"
								}
								Else
								{
									WriteWordLine 0 0 "No"
								}
							}
							WriteWordLine 0 5 "Print Quality`t: " -nonewline
							Switch ($SessionPrinterSettings.PrintQuality)
							{
								"Dpi600" {WriteWordLine 0 0 "600 dpi"}
								"Dpi300" {WriteWordLine 0 0 "300 dpi"}
								"Dpi150" {WriteWordLine 0 0 "150 dpi"}
								"Dpi75"  {WriteWordLine 0 0 "75 dpi"}
								Default {WriteWordLine 0 0 "Print Quality could not be determined: $($SessionPrinterSettings.PrintQuality)"}
							}
							WriteWordLine 0 5 "Orientation`t: " $SessionPrinterSettings.PaperOrientation
							WriteWordLine 0 5 "Apply customized settings at every logon: " -nonewline
							If($SessionPrinterSettings.ApplySettingsOnLogOn)
							{
								WriteWordLine 0 0 "Yes"
							}
							Else
							{
								WriteWordLine 0 0 "No"
							}
						}
					}
				}
			}
			WriteWordLine 0 3 "Client's Default printer: "
			If($Setting.SessionPrinterDefaultOption -eq "SetToPrinterIndex")
			{
				WriteWordLine 0 4 $Setting.SessionPrinterList[$Setting.SessionPrinterDefaultIndex]
			}
			Else
			{
				Switch ($Setting.SessionPrinterDefaultOption)
				{
					"SetToClientMainPrinter" {WriteWordLine 0 4 "Set Default printer to the client's main printer"}
					"DoNotAdjust"            {WriteWordLine 0 4 "Do not adjust the user's Default printer"}
					Default {WriteWordLine 0 0 "Client's Default printer could not be determined: $($Setting.SessionPrinterDefaultOption)"}
				}
				
			}
		}
	}

	#User Workspace
	$xArray = ($Setting.ConcurrentSessionsState, $Setting.ZonePreferenceAndFailoverState)
	If($xArray -contains "Enabled" -or $xArray -contains "Disabled")
	{
		Write-Verbose "$(Get-Date): `t`t`tUser Workspace\Connections"
		WriteWordLine 0 2 "User Workspace\Connections"
		If($Setting.ConcurrentSessionsState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Limit total concurrent sessions: " $Setting.ConcurrentSessionsState
			If($Setting.ConcurrentSessionsState -eq "Enabled")
			{
				WriteWordLine 0 4 "Limit: " $Setting.ConcurrentSessionsLimit
			}
		}
		If($Setting.ZonePreferenceAndFailoverState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Zone preference and failover: " $Setting.ZonePreferenceAndFailoverState
			If($Setting.ZonePreferenceAndFailoverState -eq "Enabled")
			{
				WriteWordLine 0 4 "Zone preference settings:"
				ForEach($Pref in $Setting.ZonePreferences)
				{
					WriteWordLine 0 5 $Pref
				}
			}
		}
	}
	
	If($Setting.ContentRedirectionState -ne "NotConfigured")
	{
		Write-Verbose "$(Get-Date): `t`t`tUser Workspace\Content Redirection"
		WriteWordLine 0 2 "User Workspace\Content Redirection"
		WriteWordLine 0 3 "Server to client: " $Setting.ContentRedirectionState
		If($Setting.ContentRedirectionState -eq "Enabled")
		{
			If($Setting.ContentRedirectionIsUsed)
			{
				WriteWordLine 0 4 "Use Content Redirection from server to client"
			}
			Else
			{
				WriteWordLine 0 4 "Do not use Content Redirection from server to client"
			}
		}
	}

	$xArray = ($Setting.ShadowingState, $Setting.ShadowingPermissionsState)
	If($xArray -contains "Enabled" -or $xArray -contains "Disabled")
	{
		Write-Verbose "$(Get-Date): `t`t`tUser Workspace\Shadowing"
		WriteWordLine 0 2 "User Workspace\Shadowing"
		If($Setting.ShadowingState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Configuration: " $Setting.ShadowingState
			If($Setting.ShadowingState -eq "Enabled")
			{
				If($Setting.ShadowingAllowed)
				{
					WriteWordLine 0 4 "Allow Shadowing"
					WriteWordLine 0 4 "Prohibit Being Shadowed Without Notification: " $Setting.ShadowingProhibitedWithoutNotification
					WriteWordLine 0 4 "Prohibit Remote Input When Being Shadowed: " $Setting.ShadowingRemoteInputProhibited
				}
				Else
				{
					WriteWordLine 0 3 "Do Not Allow Shadowing"
				}
			}
		}
		If($Setting.ShadowingPermissionsState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Permissions: " $Setting.ShadowingPermissionsState
			If($Setting.ShadowingPermissionsState -eq "Enabled")
			{
				If($Setting.ShadowingAccountsAllowed)
				{
					WriteWordLine 0 4 "Accounts allowed to shadow:"
					ForEach($Allowed in $Setting.ShadowingAccountsAllowed)
					{
						WriteWordLine 0 5 $Allowed
					}
				}
				If($Setting.ShadowingAccountsDenied)
				{
					WriteWordLine 0 4 "Accounts denied from shadowing:"
					ForEach($Denied in $Setting.ShadowingAccountsDenied)
					{
						WriteWordLine 0 5 $Denied
					}
				}
			}
		}
	}

	$xArray = ($Setting.TurnClientLocalTimeEstimationOffState, $Setting.TurnClientLocalTimeEstimationOffState)
	If($xArray -contains "Enabled" -or $xArray -contains "Disabled")
	{
		Write-Verbose "$(Get-Date): `t`t`tUser Workspace\Time Zones"
		WriteWordLine 0 2 "User Workspace\Time Zones"
		If($Setting.TurnClientLocalTimeEstimationOffState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Do not estimate local time for legacy clients: " $Setting.TurnClientLocalTimeEstimationOffState
		}
		If($Setting.TurnClientLocalTimeEstimationOffState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Do not use Client's local time: " $Setting.TurnClientLocalTimeOffState
		}
	}
	
	$xArray = ($Setting.CentralCredentialStoreState, $Setting.TurnPasswordManagerOffState)
	If($xArray -contains "Enabled" -or $xArray -contains "Disabled")
	{
		Write-Verbose "$(Get-Date): `t`t`tUser Workspace\Citrix Password Manager"
		WriteWordLine 0 2 "User Workspace\Citrix Password Manager"
		If($Setting.CentralCredentialStoreState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Central Credential Store: " $Setting.CentralCredentialStoreState
			If($Setting.CentralCredentialStoreState -eq "Enabled")
			{
				If($Setting.CentralCredentialStorePath)
				{
					WriteWordLine 0 4 "UNC path of Central Credential Store: " $Setting.CentralCredentialStorePath
				}
				Else
				{
					WriteWordLine 0 4 "No UNC path to Central Credential Store entered"
				}
			}
		}
		If($Setting.TurnPasswordManagerOffState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Do not use Citrix Password Manager: " $Setting.TurnPasswordManagerOffState
		}
	}
	
	If($Setting.StreamingDeliveryProtocolState -ne "NotConfigured")
	{
		Write-Verbose "$(Get-Date): `t`t`tUser Workspace\Streamed Applications"
		WriteWordLine 0 2 "User Workspace\Streamed Applications"
		WriteWordLine 0 3 "Configure delivery protocol: " $Setting.StreamingDeliveryProtocolState
		If($Setting.StreamingDeliveryProtocolState -eq "Enabled")
		{
			WriteWordLine 0 4 "Streaming Delivery Protocol option: " 
			Switch ($Setting.StreamingDeliveryOption)
			{
				"Unknown"                {WriteWordLine 0 5 "Unknown"}
				"ForceServerAccess"      {WriteWordLine 0 5 "Do not allow applications to stream to the client"}
				"ForcedStreamedDelivery" {WriteWordLine 0 5 "Force applications to stream to the client"}
				Default {WriteWordLine 0 0 "Delivery protocol could not be determined: $($LoadBalancingPolicyConfiguration.StreamingDeliveryOption)"}
			}
		}
	}

	#Security
	If($Setting.SecureIcaEncriptionState -ne "NotConfigured")
	{
		Write-Verbose "$(Get-Date): `t`t`tSecurity\Encryption\SecureICA encryption"
		WriteWordLine 0 2 "Security\Encryption\SecureICA encryption: " $Setting.SecureIcaEncriptionState
		If($Setting.SecureIcaEncriptionState -eq "Enabled")
		{
			WriteWordLine 0 3 "Encryption level: " -nonewline
			Switch ($Setting.SecureIcaEncriptionLevel)
			{
				"Unknown" {WriteWordLine 0 0 "Unknown encryption"}
				"Basic"   {WriteWordLine 0 0 "Basic"}
				"LogOn"   {WriteWordLine 0 0 "RC5 (128 bit) logon only"}
				"Bits40"  {WriteWordLine 0 0 "RC5 (40 bit)"}
				"Bits56"  {WriteWordLine 0 0 "RC5 (56 bit)"}
				"Bits128" {WriteWordLine 0 0 "RC5 (128 bit)"}
				Default {WriteWordLine 0 0 "SecureICA minimum encryption level could not be determined: $($Setting.MinimumEncryptionLevel.Value)"}
			}
		}
	}
	
	If($Setting.SecureIcaEncriptionState -ne "NotConfigured")
	{
		Write-Verbose "$(Get-Date): `t`t`tService Level\Session Importance"
		WriteWordLine 0 2 "Service Level\Session Importance: " $Setting.SessionImportanceState
		If($Setting.SessionImportanceState -eq "Enabled")
		{
			WriteWordLine 0 3 "Importance level: " $Setting.SessionImportanceLevel
		}
	}
}

Function AbortScript
{
	$Script:Word.quit()
	Write-Verbose "$(Get-Date): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
	If(Test-Path variable:global:word)
	{
		Remove-Variable -Name word -Scope Global
	}
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	Write-Verbose "$(Get-Date): Script has been aborted"
	$ErrorActionPreference = $SaveEAPreference
	Exit
}

Function FindWordDocumentEnd
{
	#return focus to main document    
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument
	#move to the end of the current document
	$Script:Selection.EndKey($wdStory,$wdMove) | Out-Null
}

Function BuildTableForServer
{
	Param([Array]$xArray)
	
	If(-not ($xArray -is [Array]))
	{
		$xArray = (,$xArray)
	}
	[int]$MaxLength = 0
	[int]$TmpLength = 0
	ForEach($xName in $xArray)
	{
		$TmpLength = $xName.Length
		If($TmpLength -gt $MaxLength)
		{
			$MaxLength = $TmpLength
		}
	}
	Write-Verbose "$(Get-Date): `t`tMax length of server name is $($MaxLength)"
	$TableRange = $doc.Application.Selection.Range
	[int]$Columns = [Math]::Floor(60 / $MaxLength)
	If($xArray.count -lt $Columns)
	{
		[int]$Rows = 1
		#not enough array items to fill columns so use array count
		$MaxCells = $xArray.Count
		#reset column count so there are no empty columns
		$Columns = $xArray.Count 
	}
	Else
	{
		[int]$Rows = [Math]::Floor( ( $xArray.count + $Columns - 1 ) / $Columns)
		#more array items than columns so don't go past last column
		$MaxCells = $Columns
	}
	$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
	$Table.Style = $myHash.Word_TableGrid
	$Table.Borders.InsideLineStyle = $wdLineStyleSingle
	$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
	[int]$xRow = 1
	[int]$ArrayItem = 0
	While($xRow -le $Rows)
	{
		For($xCell=1; $xCell -le $MaxCells; $xCell++)
		{
			$Table.Cell($xRow,$xCell).Range.Text = $xArray[$ArrayItem]
			$ArrayItem++
		}
		$xRow++
	}
	$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)
	$Table.AutoFitBehavior($wdAutoFitContent)

	FindWordDocumentEnd
	$xArray = $Null
}

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
		[Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, ParameterSetName='Hashtable', Position=0)]
		[ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Hashtable,
		# Array of PSCustomObjects
		[Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, ParameterSetName='CustomObject', Position=0)]
		[ValidateNotNullOrEmpty()] [PSCustomObject[]] $CustomObject,
		# Array of Hashtable key names or PSCustomObject property names to include, in display order.
		# If not supplied then all Hashtable keys or all PSCustomObject properties will be displayed.
		[Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] [string[]] $Columns = $null,
		# Array of custom table header strings in display order.
		[Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] [string[]] $Headers = $null,
		# AutoFit table behavior.
		[Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] [int] $AutoFit = -1,
		# List view (no headers)
		[Switch] $List,
		# Grid lines
		[Switch] $NoGridLines,
		# Built-in Word table formatting style constant
		# Would recommend only $wdTableFormatContempory for normal usage (possibly $wdTableFormatList5 for List view)
		[Parameter(ValueFromPipelineByPropertyName=$true)] [int] $Format = 0
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'" -f $PSCmdlet.ParameterSetName);
		## Check if -Columns wasn't specified but -Headers were (saves some additional parameter sets!)
		If(($Columns -eq $null) -and ($Headers -ne $null)) 
		{
			Write-Warning "No columns specified and therefore, specified headers will be ignored.";
			$Columns = $null;
		}
		ElseIf(($Columns -ne $null) -and ($Headers -ne $null)) 
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
				If($Columns -eq $null) 
				{
					## Build the available columns from all availble PSCustomObject note properties
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
					If($Headers -ne $null) 
					{
                        [ref] $null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{ 
                        [ref] $null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
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
                    [ref] $null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end foreach
				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f ($CustomObject.Count));
			} ## end CustomObject

			Default 
			{   ## Hashtable
				If($Columns -eq $null) 
				{
					## Build the available columns from all available hashtable keys. Hopefully
					## all Hashtables have the same keys (they should for a table).
					$Columns = $Hashtable[0].Keys;
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Headers -ne $null) 
					{ 
                        [ref] $null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{
                        [ref] $null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
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
                    [ref] $null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
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
			$ConvertToTableArguments.Add("ApplyBorders", $true);
			$ConvertToTableArguments.Add("ApplyShading", $true);
			$ConvertToTableArguments.Add("ApplyFont", $true);
			$ConvertToTableArguments.Add("ApplyColor", $true);
			If(!$List) 
			{ 
				$ConvertToTableArguments.Add("ApplyHeadingRows", $true); 
			}
			$ConvertToTableArguments.Add("ApplyLastRow", $true);
			$ConvertToTableArguments.Add("ApplyFirstColumn", $true);
			$ConvertToTableArguments.Add("ApplyLastColumn", $true);
		}

		## Invoke ConvertToTable method - with named arguments - to convert Word range to a table
		## See http://msdn.microsoft.com/en-us/library/office/aa171893(v=office.11).aspx
		Write-Debug ("$(Get-Date): `t`tConverting range to table");
		## Store the table reference just in case we need to set alternate row coloring
		$WordTable = $WordRange.GetType().InvokeMember(
			"ConvertToTable",                               # Method name
			[System.Reflection.BindingFlags]::InvokeMethod, # Flags
			$null,                                          # Binder
			$WordRange,                                     # Target (self!)
			([Object[]]($ConvertToTableArguments.Values)),  ## Named argument values
			$null,                                          # Modifiers
			$null,                                          # Culture
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

		#the next line causes the heading row to flow across page breaks
		$WordTable.Rows.First.Headingformat = $wdHeadingFormatTrue;

		If(!$NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleSingle;
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
		[Parameter()] [AllowNull()] [string] $Font = $null,
		# Font color
		[Parameter()] [AllowNull()] $Color = $null,
		# Font size
		[Parameter()] [ValidateNotNullOrEmpty()] [int] $Size = 0,
		# Cell background color
		[Parameter()] [AllowNull()] $BackgroundColor = $null,
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
					If($BackgroundColor -ne $null) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Bold) { $Cell.Range.Font.Bold = $true; }
					If($Italic) { $Cell.Range.Font.Italic = $true; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Font -ne $null) { $Cell.Range.Font.Name = $Font; }
					If($Color -ne $null) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				} # end foreach
			} # end Collection
			'Cell' 
			{
				If($Bold) { $Cell.Range.Font.Bold = $true; }
				If($Italic) { $Cell.Range.Font.Italic = $true; }
				If($Underline) { $Cell.Range.Font.Underline = 1; }
				If($Font -ne $null) { $Cell.Range.Font.Name = $Font; }
				If($Color -ne $null) { $Cell.Range.Font.Color = $Color; }
				If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
				If($BackgroundColor -ne $null) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
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
					If($Font -ne $null) { $Cell.Range.Font.Name = $Font; }
					If($Color -ne $null) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($BackgroundColor -ne $null) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
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

Function ShowScriptOptions
{
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): Company Name : $($Script:CoName)"
	Write-Verbose "$(Get-Date): Cover Page   : $($CoverPage)"
	Write-Verbose "$(Get-Date): User Name    : $($UserName)"
	Write-Verbose "$(Get-Date): Save As PDF  : $($PDF)"
	Write-Verbose "$(Get-Date): Save As WORD : $($MSWORD)"
	Write-Verbose "$(Get-Date): Add DateTime : $($AddDateTime)"
	Write-Verbose "$(Get-Date): HW Inventory : $($Hardware)"
	Write-Verbose "$(Get-Date): SW Inventory : $Software"
	If(!$Summary)
	{
		Write-Verbose "$(Get-Date): Start Date   : $($StartDate)"
		Write-Verbose "$(Get-Date): End Date     : $($EndDate)"
	}
	Write-Verbose "$(Get-Date): Section      : $($Section)"
	Write-Verbose "$(Get-Date): Summary      : $($Summary)"
	Write-Verbose "$(Get-Date): Filename1    : $($Script:FileName1)"
	If($PDF)
	{
		Write-Verbose "$(Get-Date): Filename2    : $($Script:FileName2)"
	}
	Write-Verbose "$(Get-Date): OS Detected  : $($RunningOS)"
	Write-Verbose "$(Get-Date): PSUICulture  : $($PSUICulture)"
	Write-Verbose "$(Get-Date): PSCulture    : $($PSCulture)"
	Write-Verbose "$(Get-Date): Word version : $($Script:WordProduct)"
	Write-Verbose "$(Get-Date): Word language: $($Script:WordLanguageValue)"
	Write-Verbose "$(Get-Date): PoSH version : $($Host.Version)"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): Script start : $($Script:StartTime)"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
}

Function validStateProp( [object] $object, [string] $topLevel, [string] $secondLevel )
{
	#function created 8-jan-2014 by Michael B. Smith
	if( $object )
	{
		If( ( gm -Name $topLevel -InputObject $object ) )
		{
			If( ( gm -Name $secondLevel -InputObject $object.$topLevel ) )
			{
				Return $True
			}
		}
	}
	Return $False
}

Function SetupWord
{
	Write-Verbose "$(Get-Date): Setting up Word"
    
	# Setup word for output
	Write-Verbose "$(Get-Date): Create Word comObject.  Ignore the next message."
	$Script:Word = New-Object -comobject "Word.Application" -EA 0
	
	If(!$? -or $Script:Word -eq $Null)
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
	Write-Verbose "$(Get-Date): Word language value is $($Script:WordLanguageValue)"
	
	$Script:WordCultureCode = GetCulture $Script:WordLanguageValue
	
	SetWordHashTable $Script:WordCultureCode
	
	[int]$Script:WordVersion = [int]$Script:Word.Version
	If($Script:WordVersion -eq $wdWord2013)
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
	If([String]::IsNullOrEmpty($CoName))
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
			Write-Verbose "$(Get-Date): Updated company name to $($Script:CoName)"
		}
	}

	If($Script:WordCultureCode -ne "en-")
	{
		Write-Verbose "$(Get-Date): Check Default Cover Page for $($WordCultureCode)"
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
						If($Script:WordVersion -eq $wdWord2013)
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
		}

		If($CPChanged)
		{
			Write-Verbose "$(Get-Date): Changed Default Cover Page from Sideline to $($CoverPage)"
		}
	}

	Write-Verbose "$(Get-Date): Validate cover page $($CoverPage) for culture code $($Script:WordCultureCode)"
	[bool]$ValidCP = $False
	
	$ValidCP = ValidateCoverPage $Script:WordVersion $CoverPage $Script:WordCultureCode
	
	If(!$ValidCP)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Verbose "$(Get-Date): Word language value $($Script:WordLanguageValue)"
		Write-Verbose "$(Get-Date): Culture code $($Script:WordCultureCode)"
		Write-Error "`n`n`t`tFor $($Script:WordProduct), $($CoverPage) is not a valid Cover Page option.`n`n`t`tScript cannot continue.`n`n"
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
	#word 2010/2013
	$BuildingBlocksCollection = $Script:Word.Templates | Where {$_.name -eq "Built-In Building Blocks.dotx"}

	Write-Verbose "$(Get-Date): Attempt to load cover page $($CoverPage)"
	$part = $Null

	$BuildingBlocksCollection | 
	ForEach{
		If ($_.BuildingBlockEntries.Item($CoverPage).Name -eq $CoverPage) 
		{
			$BuildingBlocks = $_
		}
	}        

	If($BuildingBlocks -ne $Null)
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

		If($part -ne $Null)
		{
			$Script:CoverPagesExist = $True
		}
	}

	If(!$Script:CoverPagesExist)
	{
		Write-Verbose "$(Get-Date): Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
		Write-Warning "Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
		Write-Warning "This report will not have a Cover Page."
	}

	Write-Verbose "$(Get-Date): Create empty word doc"
	$Script:Doc = $Script:Word.Documents.Add()
	If($Script:Doc -eq $Null)
	{
		Write-Verbose "$(Get-Date): "
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tAn empty Word document could not be created.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}

	$Script:Selection = $Script:Word.Selection
	If($Script:Selection -eq $Null)
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
		Write-Verbose "$(Get-Date): Table of Contents - $($myHash.Word_TableOfContents)"
		$toc = $BuildingBlocks.BuildingBlockEntries.Item($myHash.Word_TableOfContents)
		If($toc -eq $Null)
		{
			Write-Verbose "$(Get-Date): "
			Write-Verbose "$(Get-Date): Table of Content - $($myHash.Word_TableOfContents) could not be retrieved."
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
	#Update document properties
	If($MSWORD -or $PDF)
	{
		If($Script:CoverPagesExist)
		{
			Write-Verbose "$(Get-Date): Set Cover Page Properties"
			_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Company" $Script:CoName
			_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Title" $Script:title
			_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Author" $username

			_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Subject" $SubjectTitle

			#Get the Coverpage XML part
			$cp = $Script:Doc.CustomXMLParts | Where {$_.NamespaceURI -match "coverPageProps$"}

			#get the abstract XML part
			$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "Abstract"}

			#set the text
			If([String]::IsNullOrEmpty($Script:CoName))
			{
				[string]$abstract = $AbstractTitle
			}
			Else
			{
				[string]$abstract = "$($AbstractTitle) for $Script:CoName"
			}

			$ab.Text = $abstract

			$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "PublishDate"}
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
		Write-Verbose "$(Get-Date): Running Word 2010 and detected operating system $($RunningOS)"
		$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
		$Script:Doc.SaveAs([REF]$Script:FileName1, [ref]$SaveFormat)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
			$Script:Doc.SaveAs([REF]$Script:FileName2, [ref]$saveFormat)
		}
	}
	ElseIf($Script:WordVersion -eq $wdWord2013)
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
		Write-Verbose "$(Get-Date): Running Word 2013 and detected operating system $($RunningOS)"
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
		Write-Verbose "$(Get-Date): Deleting $($Script:FileName1) since only $($Script:FileName2) is needed"
		Remove-Item $Script:FileName1
	}
	Write-Verbose "$(Get-Date): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
	If(Test-Path variable:global:word)
	{
		Remove-Variable -Name word -Scope Global
	}
	$SaveFormat = $Null
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
}

Function SetFileName1andFileName2
{
	Param([string]$OutputFileName)
	$pwdpath = $pwd.Path

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
}

#script begins

$Script:StartTime = get-date

If(!(Check-NeededPSSnapins "Citrix.XenApp.Commands")){
	#We're missing Citrix Snapins that we need
	$ErrorActionPreference = $SaveEAPreference
	Write-Error "Missing Citrix PowerShell Snap-ins Detected, check the console above for more information. Are you sure you are running this script on a XenApp 5 Server? Script will now close."
	Exit
}

#if software inventory is specified then verify SoftwareExclusions.txt exists
If($Software)
{
	If(!(Test-Path "$($pwd.path)\SoftwareExclusions.txt"))
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "Software inventory requested but $($pwd.path)\SoftwareExclusions.txt does not exist.  Script cannot continue."
		Exit
	}
	
	#file does exist but can we access it?
	$x = Get-Content "$($pwd.path)\SoftwareExclusions.txt" -EA 0
	If(!($?))
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "There was an error accessing or reading $($pwd.path)\SoftwareExclusions.txt.  Script cannot continue."
		Exit
	}
	$x = $Null
}

Write-Verbose "$(Get-Date): Getting initial Farm data"
$farm = Get-XAFarm -EA 0
If($? -and $Farm -ne $Null)
{
	Write-Verbose "$(Get-Date): Verify farm version"
	#first check to make sure this is a XenApp 5 farm
	If($Farm.ServerVersion.ToString().SubString(0,1) -ne "6")
	{
		If($Farm.ServerVersion.ToString().SubString(0,1) -eq "4")
		{
			$FarmOS = "2003"
		}
		Else
		{
			$FarmOS = "2008"
		}
		Write-Verbose "$(Get-Date): Farm OS is $($FarmOS)"
		#this is a XenApp 5 farm, script can proceed
		#XenApp 5 for server 2003 shows as version 4.6
		#XenApp 5 for server 2008 shows as version 5.0
	}
	Else
	{
		#this is not a XenApp 5 farm, script cannot proceed
		Write-Warning "This script is designed for XenApp 5 and should not be run on XenApp 6.x"
		Return 1
	}

	If($FarmOS -eq "2003")
	{
		#this is not a XenApp 5 for Windows Server 2008 farm, script cannot proceed
		Write-Warning "This script is designed for XenApp 5 for Windows Server 2008`nand should not be run on XenApp 5 for Windows Server 2003"
		Return 1
	}
	
	[string]$FarmName = $farm.FarmName
	[string]$Title = "Inventory Report for the $($FarmName) Farm"
	SetFileName1andFileName2 "$($farm.FarmName)"
} 
Else 
{
	$ErrorActionPreference = $SaveEAPreference
	Write-Warning "Farm information could not be retrieved"
	Write-Error "Farm information could not be retrieved.  Script cannot continue."
	Exit
}
$farm = $Null

#process the nodes in the Access Management Console (XA5/2008)

If(!$Summary -and ($Section -eq "All" -or $Section -eq "Farm"))
{
	# Get farm information
	Write-Verbose "$(Get-Date): Getting Farm Configuration data"
	$Server2008 = $False
	$ConfigLog = $False
	
	$farm = Get-XAFarmConfiguration -EA 0

	If($? -and $farm -ne $Null)
	{
		If($CoverPagesExist)
		{
			#only need the blank page inserted if there is a Table of Contents
			$selection.InsertNewPage()
		}
		WriteWordLine 1 0 "Farm Configuration Settings"
		
		WriteWordLine 2 0 "Farm-wide"

		Write-Verbose "$(Get-Date): `tFarm-wide"
		Write-Verbose "$(Get-Date): `t`tConnection Access Controls"
		WriteWordLine 0 1 "Connection Access Controls"
		
		Switch ($Farm.ConnectionAccessControls)
		{
			"AllowAnyConnection" {WriteWordLine 0 2 "Any connections"}
			"AllowOneTypeOnly"   {WriteWordLine 0 2 "Citrix Access Gateway, Citrix XenApp plug-in, and Web Interface connections only"}
			"AllowMultipleTypes" {WriteWordLine 0 2 "Citrix Access Gateway connections only"}
			Default {WriteWordLine 0 0 "Connection Access Controls could not be determined: $($Farm.ConnectionAccessControls)"}
		}

		Write-Verbose "$(Get-Date): `t`tConnection Limits"
		WriteWordLine 0 1 "Connection Limits" 
		WriteWordLine 0 2 "Connections per user"
		WriteWordLine 0 3 "Maximum connections per user: " -NoNewLine
		If($Farm.ConnectionLimitsMaximumPerUser -eq -1)
		{
			WriteWordLine 0 0 "No limit set"
		}
		Else
		{
			WriteWordLine 0 0 $Farm.ConnectionLimitsMaximumPerUser
		}
		If($Farm.ConnectionLimitsEnforceAdministrators)
		{
			WriteWordLine 0 3 "Enforce limit on administrators"
		}
		Else
		{
			WriteWordLine 0 3 "Do not enforce limit on administrators"
		}

		If($Farm.ConnectionLimitsLogOverLimits)
		{
			WriteWordLine 0 3 "Log over-the-limit denials"
		}
		Else
		{
			WriteWordLine 0 3 "Do not log over-the-limit denials"
		}

		Write-Verbose "$(Get-Date): `t`tHealth Monitoring & Recovery"
		WriteWordLine 0 1 "Health Monitoring & Recovery"
		WriteWordLine 0 2 "Limit server for load balancing"
		WriteWordLine 0 3 "Limit servers (%): " $Farm.HmrMaximumServerPercent

		Write-Verbose "$(Get-Date): `t`tConfiguration Logging"
		WriteWordLine 0 1 "Configuration Logging"
		If($Farm.ConfigLogEnabled)
		{
			$ConfigLog = $True

			WriteWordLine 0 2 "Database configuration"
			WriteWordLine 0 3 "Database type: " -nonewline
			Switch ($Farm.ConfigLogDatabaseType)
			{
				"SqlServer" {WriteWordLine 0 0 "Microsoft SQL Server"}
				"Oracle"    {WriteWordLine 0 0 "Oracle"}
				Default {WriteWordLine 0 0 "Database type could not be determined: $($Farm.ConfigLogDatabaseType)"}
			}
			If($Farm.ConfigLogDatabaseAuthenticationMode -eq "Native")
			{
				WriteWordLine 0 3 "Use SQL Server authentication"
			}
			Else
			{
				WriteWordLine 0 3 "Use Windows integrated security"
			}

			WriteWordLine 0 3 "Connection String: " -NoNewLine

			$StringMembers = "`n`t`t`t`t`t" + $Farm.ConfigLogDatabaseConnectionString.replace(";","`n`t`t`t`t`t")
			
			WriteWordLine 0 3 $StringMembers -NoNewLine
			WriteWordLine 0 0 "User name=" $Farm.ConfigLogDatabaseUserName

			WriteWordLine 0 3 "Log administrative tasks to Configuration Logging database: " -nonewline
			If($Farm.ConfigLogEnabled)
			{
				WriteWordLine 0 0 "Yes"
			}
			Else
			{
				WriteWordLine 0 0 "No"
			}
			WriteWordLine 0 3 "Allow changes to the farm when logging database is disconnected: " -nonewline
			
			If($Farm.ConfigLogChangesWhileDisconnectedAllowed)
			{
				WriteWordLine 0 0 "Yes"
			}
			Else
			{
				WriteWordLine 0 0 "No"
			}
			WriteWordLine 0 3 "Require admins to enter database credentials before clearing the log: " -nonewline
			If($Farm.ConfigLogCredentialsOnClearLogRequired)
			{
				WriteWordLine 0 0 "Yes"
			}
			Else
			{
				WriteWordLine 0 0 "No"
			}
		}
		Else
		{
			WriteWordLine 0 2 "Configuration logging is not enabled"
		}
		
		Write-Verbose "$(Get-Date): `t`tMemory/CPU"	
		WriteWordLine 0 1 "Memory/CPU"

		WriteWordLine 0 2 "Applications that memory optimization ignores: "
		If($Farm.MemoryOptimizationExcludedApplications)
		{
			ForEach($App in $Farm.MemoryOptimizationExcludedApplications)
			{
				WriteWordLine 0 3 $App
			}
		}
		Else
		{
			WriteWordLine 0 3 "No applications are listed"
		}

		WriteWordLine 0 2 "Optimization interval: " $Farm.MemoryOptimizationScheduleType

		If($Farm.MemoryOptimizationScheduleType -eq "Weekly")
		{
			WriteWordLine 0 2 "Day of week: " $Farm.MemoryOptimizationScheduleDayOfWeek
		}
		If($Farm.MemoryOptimizationScheduleType -eq "Monthly")
		{
			WriteWordLine 0 2 "Day of month: " $Farm.MemoryOptimizationScheduleDayOfMonth
		}

		WriteWordLine 0 2 "Optimization time: " $Farm.MemoryOptimizationScheduleTime
		WriteWordLine 0 2 "Memory optimization user: " -nonewline
		If($Farm.MemoryOptimizationLocalSystemAccountUsed)
		{
			WriteWordLine 0 0 "Use local system account"
		}
		Else
		{
			WriteWordLine 0 0 $Farm.MemoryOptimizationUser
		}
		
		Write-Verbose "$(Get-Date): `t`tXenApp"
		WriteWordLine 0 1 "XenApp"
		WriteWordLine 0 2 "General"
		WriteWordLine 0 3 "Respond to client broadcast messages"
		WriteWordLine 0 4 "Data collectors: " -nonewline
		If($Farm.RespondDataCollectors)
		{
			WriteWordLine 0 0 "Yes"
		}
		Else
		{
			WriteWordLine 0 0 "No"
		}
		WriteWordLine 0 4 "RAS servers: " -nonewline
		If($Farm.RespondRasServers)
		{
			WriteWordLine 0 0 "Yes"
		}
		Else
		{
			WriteWordLine 0 0 "No"
		}
		WriteWordLine 0 3 "Client time zones"
		WriteWordLine 0 4 "Use client's local time: " -nonewline
		If($Farm.ClientLocalTimeEnabled)
		{
			WriteWordLine 0 0 "Yes"
		}
		Else
		{
			WriteWordLine 0 0 "No"
		}
		WriteWordLine 0 4 "Estimate local time for clients: " -nonewline
		If($Farm.ClientLocalTimeEstimationEnabled)
		{
			WriteWordLine 0 0 "Yes"
		}
		Else
		{
			WriteWordLine 0 0 "No"
		}
		WriteWordLine 0 3 "XML Service DNS address resolution: " -nonewline
		If($Farm.DNSAddressResolution)
		{
			WriteWordLine 0 0 "Yes"
		}
		Else
		{
			WriteWordLine 0 0 "No"
		}
		WriteWordLine 0 3 "Novell Directory Services"
		WriteWordLine 0 4 "NDS preferred tree: " -NoNewLine
		If($Farm.NdsPreferredTree)
		{
			WriteWordLine 0 0 $Farm.NdsPreferredTree
		}
		Else
		{
			WriteWordLine 0 0 "No NDS Tree entered"
		}
		WriteWordLine 0 3 "Enhanced icon support: " -nonewline
		If($Farm.EnhancedIconEnabled)
		{
			WriteWordLine 0 0 "Yes"
		}
		Else
		{
			WriteWordLine 0 0 "No"
		}

		WriteWordLine 0 2 "Shadow Policies"
		WriteWordLine 0 3 "Merge shadowers in multiple policies: " -nonewline
		If($Farm.ShadowPoliciesMerge)
		{
			WriteWordLine 0 0 "Yes"
		}
		Else
		{
			WriteWordLine 0 0 "No"
		}

		Write-Verbose "$(Get-Date): `t`tSession Reliability"
		WriteWordLine 0 1 "Session Reliability"
		
		WriteWordLine 0 2 "Keep sessions open during loss of network connectivity: " -nonewline
		If($Farm.SessionReliabilityEnabled)
		{
			WriteWordLine 0 0 "Yes"
		}
		Else
		{
			WriteWordLine 0 0 "No"
		}
		WriteWordLine 0 2 "Port number (Default 2598): " $Farm.SessionReliabilityPort
		
		WriteWordLine 0 2 "Seconds to keep sessions open: " $Farm.SessionReliabilityTimeout

		Write-Verbose "$(Get-Date): `t`tCitrix Streaming Server"
		WriteWordLine 0 1 "Citrix Streaming Server"
		WriteWordLine 0 2 "Log application streaming events to event log: " -nonewline
		If($Farm.StreamingLogEvents)
		{
			WriteWordLine 0 0 "Yes"
		}
		Else
		{
			WriteWordLine 0 0 "No"
		}
		WriteWordLine 0 2 "Trust XenApp Plugin for Streamed Apps: " -nonewline
		If($Farm.StreamingTrustCLient)
		{
			WriteWordLine 0 0 "Yes"
		}
		Else
		{
			WriteWordLine 0 0 "No"
		}

		Write-Verbose "$(Get-Date): `t`tRestart Options"
		WriteWordLine 0 1 "Restart Options"
		WriteWordLine 0 2 "Message Options"
		WriteWordLine 0 3 "Send message to logged-on users before server restart: " -nonewline
		If($Farm.RestartSendMessage)
		{
			WriteWordLine 0 0 "Yes"
		}
		Else
		{
			WriteWordLine 0 0 "No"
		}
		WriteWordLine 0 3 "Send first message before restart: $($Farm.RestartMessageWait) minutes"
		WriteWordLine 0 3 "Send reminder message every: $($Farm.RestartMessageInterval) minutes"
		If($Farm.RestartCustomMessageEnabled)
		{
			WriteWordLine 0 3 "Additional text for restart message:"
			WriteWordLine 0 3 $Farm.RestartCustomMessage
		}
		If($Farm.RestartDisabledLogOnsInterval -gt 0)
		{
			WriteWordLine 0 3 "Disable logons before restart: " -nonewline
			WriteWordLine 0 0 "Yes"
			WriteWordLine 0 3 "Before restart, logons disabled by: $($Farm.RestartDisabledLogOnsInterval) minutes"
		}
		Else
		{
			WriteWordLine 0 3 "Disable logons before restart: No"
		}
		
		Write-Verbose "$(Get-Date): `t`tVirtual IP"
		WriteWordLine 0 1 "Virtual IP"
		WriteWordLine 0 2 "Address Configuration"
		WriteWordLine 0 3 "Virtual IP address ranges:"

		$VirtualIPs = Get-XAVirtualIPRange -EA 0
		If($? -and $VirtualIPs)
		{
			ForEach($VirtualIP in $VirtualIPs)
			{
				WriteWordLine 0 4 "IP Range: " $VirtualIP
			}
		}
		Else
		{
			WriteWordLine 0 4 "No virtual IP address range defined"
		}
		$VirtualIPs = $Null

		WriteWordLine 0 3 "Enable logging of IP address assignment and release: " -nonewline
		If($Farm.VirtualIPLoggingEnabled)
		{
			WriteWordLine 0 0 "Yes"
		}
		Else
		{
			WriteWordLine 0 0 "No"
		}
		WriteWordLine 0 2 "Process Configuration"
		WriteWordLine 0 3 "Virtual IP Processes"
		If($Farm.VirtualIPProcesses)
		{
			WriteWordLine 0 4 "Monitor the following processes:"
			ForEach($Process in $Farm.VirtualIPProcesses)
			{
				WriteWordLine 0 5 "Process: " $Process
			}
		}
		Else
		{
			WriteWordLine 0 4 "No virtual IP processes defined"
		}
		WriteWordLine 0 3 "Virtual Loopback Processes"
		If($Farm.VirtualIPLoopbackProcesses)
		{
			WriteWordLine 0 4 "Monitor the following processes:"
			ForEach($Process in $Farm.VirtualIPLoopbackProcesses)
			{
				WriteWordLine 0 5 "Process: " $Process
			}
		}
		Else
		{
			WriteWordLine 0 4 "No virtual IP Loopback processes defined"
		}
			
		$selection.InsertNewPage()
		Write-Verbose "$(Get-Date): `tServer Default"
		WriteWordLine 2 0 "Server Default"
		Write-Verbose "$(Get-Date): `t`tICA"
		WriteWordLine 0 1 "ICA"
		WriteWordLine 0 2 "Auto Client Reconnect"
		If($Farm.AcrEnabled)
		{
			WriteWordLine 0 3 "Reconnect automatically"
			WriteWordLine 0 3 "Log automatic reconnection attempts: " -NoNewLine

			If($Farm.AcrLogReconnections)
			{
				WriteWordLine 0 0 "Yes"
			}
			Else
			{
				WriteWordLine 0 0 "No"
			}
		}
		Else
		{
			WriteWordLine 0 4 "Require user authentication"
		}
		
		WriteWordLine 0 2 "Display"
		WriteWordLine 0 3 "Discard queued image that is replaced by another image: " -nonewline
		If($Farm.DisplayDiscardQueuedImages)
		{
			WriteWordLine 0 0 "Yes"
		}
		Else
		{
			WriteWordLine 0 0 "No"
		}
		WriteWordLine 0 3 "Cache image to make scrolling smoother: " -nonewline
		If($Farm.DisplayCacheImageForSmoothScrolling)
		{
			WriteWordLine 0 0 "Yes"
		}
		Else
		{
			WriteWordLine 0 0 "No"
		}
		WriteWordLine 0 3 "Maximum memory to use for each session's graphics (KB): " $Farm.DisplayMaximumGraphicsMemory
		WriteWordLine 0 3 "Degradation bias"
		If($Farm.DisplayDegradationBias -eq "Resolution")
		{
			WriteWordLine 0 4 "Degrade resolution first"
		}
		Else
		{
			WriteWordLine 0 4 "Degrade color depth first"
		}
		WriteWordLine 0 3 "Notify user of session degradation: " -nonewline
		If($Farm.DisplayNotifyUser)
		{
			WriteWordLine 0 0 "Yes"
		}
		Else
		{
			WriteWordLine 0 0 "No"
		}

		WriteWordLine 0 2 "Keep-Alive"
		If($Farm.KeepAliveEnabled)
		{
			WriteWordLine 0 3 "ICA Keep-Alive time-out value (seconds): " $Farm.KeepAliveTimeout
		}
		Else
		{
			WriteWordLine 0 3 "ICA Keep-Alive is not enabled"
		}
		
		Write-Verbose "$(Get-Date): `t`tLicense Server"
		WriteWordLine 0 1 "License Server"
		WriteWordLine 0 2 "Name: " $Farm.LicenseServerName
		WriteWordLine 0 2 "Port number (Default 27000): " $Farm.LicenseServerPortNumber
		
		Write-Verbose "$(Get-Date): `t`tMemory/CPU"
		WriteWordLine 0 1 "Memory/CPU"
		WriteWordLine 0 2 "CPU Utilization Management: " -NoNewLine
		WriteWordLine 0 0 "" -nonewline
		Switch ($Farm.CpuManagementLevel)
		{
			"NoManagement"  {WriteWordLine 0 0 "No CPU utilization management"}
			"Fair"          {WriteWordLine 0 0 "Fair sharing of CPU between sessions"}
			"ResourceBased" {WriteWordLine 0 0 "CPU Sharing based on Resource Allotments"}
			Default {WriteWordLine 0 0 "CPU Utilization Management could not be determined: $($Farm.CpuManagementLevel)"}
		}
		WriteWordLine 0 2 "Memory Optimization: " -nonewline
		If($Farm.MemoryOptimizationEnabled)
		{
			WriteWordLine 0 0 "Enabled"
		}
		Else
		{
			WriteWordLine 0 0 "Not Enabled"
		}
		
		Write-Verbose "$(Get-Date): `t`tHealth Monitoring & Recovery"
		WriteWordLine 0 1 "Health Monitoring & Recovery"
		If($Farm.HmrEnabled)
		{
			$HmrTests = Get-XAHmrTest -EA 0 | Sort TestName
			If($?)
			{
				ForEach($HmrTest in $HmrTests)
				{
					Write-Verbose "$(Get-Date): `t`t`tCreate Table for HMR Test $($HmrTest.TestName)"
					$TableRange = $doc.Application.Selection.Range
					[int]$Columns = 2
					If(![String]::IsNullOrEmpty($Hmrtest.Arguments))
					{
						[int]$Rows = 9
					}
					Else
					{
						[int]$Rows = 8
					}
					$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
					$Table.Style = $myHash.Word_TableGrid
					$Table.Borders.InsideLineStyle = $wdLineStyleNone
					$Table.Borders.OutsideLineStyle = $wdLineStyleNone
					[int]$xRow = 1
					$Table.Cell($xRow,1).Range.Text = "Test Name"
					$Table.Cell($xRow,2).Range.Text = $Hmrtest.TestName
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "Interval"
					$Table.Cell($xRow,2).Range.Text = $Hmrtest.Interval
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "Threshold"
					$Table.Cell($xRow,2).Range.Text = $Hmrtest.Threshold
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "Time-out"
					$Table.Cell($xRow,2).Range.Text = $Hmrtest.Timeout
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "Test File Name"
					$Table.Cell($xRow,2).Range.Text = $Hmrtest.FilePath
					If(![String]::IsNullOrEmpty($Hmrtest.Arguments))
					{
						$xRow++
						$Table.Cell($xRow,1).Range.Text = "Arguments"
						$Table.Cell($xRow,2).Range.Text = $Hmrtest.Arguments
					}
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "Recovery Action"
					Switch ($Hmrtest.RecoveryAction)
					{
						"AlertOnly"                     {$Table.Cell($xRow,2).Range.Text = "Alert Only"}
						"RemoveServerFromLoadBalancing" {$Table.Cell($xRow,2).Range.Text = "Remove Server from load balancing"}
						"RestartIma"                    {$Table.Cell($xRow,2).Range.Text = "Restart IMA"}
						"ShutdownIma"                   {$Table.Cell($xRow,2).Range.Text = "Shutdown IMA"}
						"RebootServer"                  {$Table.Cell($xRow,2).Range.Text = "Reboot Server"}
						Default {$Table.Cell($xRow,2).Range.Text = "Recovery Action could not be determined: $($Hmrtest.RecoveryAction)"}
					}
					If(![String]::IsNullOrEmpty($Hmrtest.Description))
					{
						$xRow++
						$Table.Cell($xRow,1).Range.Text = "Test Description"
						$Table.Cell($xRow,2).Range.Text = $Hmrtest.Description
					}

					$Table.Rows.SetLeftIndent($Indent3TabStops,$wdAdjustProportional)
					$Table.AutoFitBehavior($wdAutoFitContent)

					FindWordDocumentEnd
				}
			}
			Else
			{
				WriteWordLine 0 2 "Health Monitoring & Recovery Tests could not be retrieved"
			}
		}
		Else
		{
			WriteWordLine 0 2 "Health Monitoring & Recovery is not enabled"
		}

		Write-Verbose "$(Get-Date): `t`tXenApp"
		WriteWordLine 0 1 "XenApp"
		WriteWordLine 0 2 "Content redirection from server to client: " -nonewline
		If($Farm.ContentRedirectionEnabled)
		{
			WriteWordLine 0 0 "Yes"
		}
		Else
		{
			WriteWordLine 0 0 "No"
		}

		WriteWordLine 0 2 "Remote Console Connections"
		WriteWordLine 0 3 "Remote connections to the console: " -nonewline
		If($Farm.RemoteConsoleEnabled)
		{
			WriteWordLine 0 0 "Yes"
		}
		Else
		{
			WriteWordLine 0 0 "No"
		}

		Write-Verbose "$(Get-Date): `t`tSNMP"
		WriteWordLine 0 1 "SNMP"
		If($Farm.SnmpEnabled)
		{
			WriteWordLine 0 2 "Send session traps to selected SNMP agent on all farm servers"
			WriteWordLine 0 3 "SNMP agent session traps"
			WriteWordLine 0 4 "Logon`t`t`t: " -nonewline
			If($Farm.SnmpLogonEnabled)
			{
				WriteWordLine 0 0 "Yes"
			}
			Else
			{
				WriteWordLine 0 0 "No"
			}
			WriteWordLine 0 4 "Logoff`t`t`t: " -nonewline
			If($Farm.SnmpLogoffEnabled)
			{
				WriteWordLine 0 0 "Yes"
			}
			Else
			{
				WriteWordLine 0 0 "No"
			}
			WriteWordLine 0 4 "Disconnect`t`t: " -nonewline
			If($Farm.SnmpDisconnectEnabled)
			{
				WriteWordLine 0 0 "Yes"
			}
			Else
			{
				WriteWordLine 0 0 "No"
			}
			WriteWordLine 0 4 "Session limit per server`t: " -nonewline
			If($Farm.SnmpLimitEnabled)
			{
				WriteWordLine 0 0 " " $Farm.SnmpLimitPerServer
			}
			Else
			{
				WriteWordLine 0 0 "No"
			}
		}
		Else
		{
			WriteWordLine 0 2 "SNMP is not enabled"
		}

		Write-Verbose "$(Get-Date): `t`tSpeedScreen"
		WriteWordLine 0 1 "SpeedScreen"
		If($Farm.BrowserAccelerationEnabled)
		{
			WriteWordLine 0 2 "SpeedScreen Browser Acceleration is enabled"
			If($Farm.BrowserAccelerationCompressionEnabled)
			{
				WriteWordLine 0 3 "Compress JPEG images to improve bandwidth"
				WriteWordLine 0 4 "Image compression levels: " $Farm.BrowserAccelerationCompressionLevel
				If($Farm.BrowserAccelerationVariableImageCompression)
				{
					WriteWordLine 0 4 "Adjust compression level based on available bandwidth"
				}
				Else
				{
					WriteWordLine 0 4 "Do not adjust compression level based on available bandwidth"
				}
			}
			Else
			{
				WriteWordLine 0 3 "Do not compress JPEG images to improve bandwidth"
			}
		}
		Else
		{
			WriteWordLine 0 2 "SpeedScreen Browser Acceleration is disabled"
		}
		
		If($Farm.FlashAccelerationEnabled)
		{
			WriteWordLine 0 2 "Enable Adobe Flash Player"
			Switch ($Farm.FlashAccelerationOption)
			{
				"AllConnections" {WriteWordLine 0 3 "Accelerate for restricted bandwidth connections"}
				"Unknown"        {WriteWordLine 0 3 "Do not accelerate"}
				"NoOptimization" {WriteWordLine 0 3 "Accelerate for all connections"}
				Default {WriteWordLine 0 0 "Server-side acceleration could not be determined: $($Farm.FlashAccelerationOption)"}
			}
			
		}
		Else
		{
			WriteWordLine 0 2 "Adobe Flash is not enabled"
		}
		If($Farm.MultimediaAccelerationEnabled)
		{
			WriteWordLine 0 2 "SpeedScreen Multimedia Acceleration is enabled"
			If($Farm.MultimediaAccelerationDefaultBuffer)
			{
				WriteWordLine 0 3 "Use the Default buffer of 5 seconds"
			}
			Else
			{
				WriteWordLine 0 3 "Custom buffer in seconds: " $Farm.MultimediaAccelerationCustomBuffer
			}
		}
		Else
		{
			WriteWordLine 0 2 "SpeedScreen Multimedia Acceleration is disabled"
		}
		
		WriteWordLine 0 0 "Offline Access"
		Write-Verbose "$(Get-Date): `t`tUsers"
		WriteWordLine 0 1 "Users"
		If($Farm.OfflineAccounts)
		{
			WriteWordLine 0 2 "Configured users:"
			ForEach($User in $Farm.OfflineAccounts)
			{
				WriteWordLine 0 3 $User
			}
		}
		Else
		{
			WriteWordLine 0 2 "No users configured"
		}

		Write-Verbose "$(Get-Date): `t`tOffline License Settings"
		WriteWordLine 0 1 "Offline License Settings"
		WriteWordLine 0 2 "License period days: " $Farm.OfflineLicensePeriod

	} 
	ElseIf(!$?)
	{
		Write-Warning "Farm configuration information could not be retrieved"
	}
	Else
	{
		Write-Warning "No results returned for Farm configuration information"
	}
	$farm = $Null
	Write-Verbose "$(Get-Date): Finished getting Farm Configuration data"
	Write-Verbose "$(Get-Date): "
}

If($Section -eq "All" -or $Section -eq "Admins")
{
	Write-Verbose "$(Get-Date): Processing Administrators"
	Write-Verbose "$(Get-Date): `tSetting summary variables"
	[int]$TotalFullAdmins = 0
	[int]$TotalViewAdmins = 0
	[int]$TotalCustomAdmins = 0
	[int]$TotalAdmins = 0

	Write-Verbose "$(Get-Date): `tRetrieving Administrators"
	$Administrators = Get-XAAdministrator -EA 0 | Sort AdministratorName

	If($? -and $Administrators -ne $Null)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Administrators:"
		ForEach($Administrator in $Administrators)
		{
			Write-Verbose "$(Get-Date): `t`tProcessing administrator $($Administrator.AdministratorName)"
			If(!$Summary)
			{
				WriteWordLine 2 0 $Administrator.AdministratorName
				WriteWordLine 0 1 "Administrator type: " -nonewline
				Switch ($Administrator.AdministratorType)
				{
					"Unknown"  {WriteWordLine 0 0 "Unknown"}
					"Full"     {WriteWordLine 0 0 "Full Administration"; $TotalFullAdmins++}
					"ViewOnly" {WriteWordLine 0 0 "View Only"; $TotalViewAdmins++}
					"Custom"   {WriteWordLine 0 0 "Custom"; $TotalCustomAdmins++}
					Default    {WriteWordLine 0 0 "Administrator type could not be determined: $($Administrator.AdministratorType)"}
				}
				WriteWordLine 0 0 " Administrator"
				WriteWordLine 0 1 "Administrator account is " -NoNewLine
				If($Administrator.Enabled)
				{
					WriteWordLine 0 0 "Enabled" 
				} 
				Else
				{
					WriteWordLine 0 0 "Disabled" 
				}
				If(!([String]::IsNullOrEmpty($Administrator.EmailAddress) -and [String]::IsNullOrEmpty($Administrator.SmsNumber) -and [String]::IsNullOrEmpty($Administrator.SmsGateway)))
				{
					WriteWordLine 0 1 "Alert Contact Details"
					WriteWordLine 0 2 "E-mail`t`t: " $Administrator.EmailAddress
					WriteWordLine 0 2 "SMS Number`t: " $Administrator.SmsNumber
					WriteWordLine 0 2 "SMS Gateway`t: " $Administrator.SmsGateway
				}
				If($Administrator.AdministratorType -eq "Custom") 
				{
					WriteWordLine 0 1 "Farm Privileges:"
					ForEach($farmprivilege in $Administrator.FarmPrivileges) 
					{
						Write-Verbose "$(Get-Date): `t`t`tProcessing farm privilege $farmprivilege"
						Switch ($farmprivilege)
						{
							"Unknown"                   {WriteWordLine 0 2 "Unknown"}
							"ViewFarm"                  {WriteWordLine 0 2 "View farm management"}
							"EditZone"                  {WriteWordLine 0 2 "Edit zones"}
							"EditConfigurationLog"      {WriteWordLine 0 2 "Configure logging for the farm"}
							"EditFarmOther"             {WriteWordLine 0 2 "Edit all other farm settings"}
							"ViewAdmins"                {WriteWordLine 0 2 "View Citrix administrators"}
							"LogOnConsole"              {WriteWordLine 0 2 "Log on to console"}
							"LogOnWIConsole"            {WriteWordLine 0 2 "Logon on to Web Interface console"}
							"ViewLoadEvaluators"        {WriteWordLine 0 2 "View load evaluators"}
							"AssignLoadEvaluators"      {WriteWordLine 0 2 "Assign load evaluators"}
							"EditLoadEvaluators"        {WriteWordLine 0 2 "Edit load evaluators"}
							"ViewLoadBalancingPolicies" {WriteWordLine 0 2 "View load balancing policies"}
							"EditLoadBalancingPolicies" {WriteWordLine 0 2 "Edit load balancing policies"}
							"ViewPrinterDrivers"        {WriteWordLine 0 2 "View printer drivers"}
							"ReplicatePrinterDrivers"   {WriteWordLine 0 2 "Replicate printer drivers"}
							"EditUserPolicies"          {WriteWordLine 0 2 "Edit User Policies"}
							"ViewUserPolicies"          {WriteWordLine 0 2 "View User Policies"}
							"EditOtherPrinterSettings"  {WriteWordLine 0 2 "Edit All Other Printer Settings"}
							"EditPrinterDrivers"        {WriteWordLine 0 2 "Edit Printer Drivers"}
							"EditPrinters"              {WriteWordLine 0 2 "Edit Printers"}
							"ViewPrintersAndDrivers"    {WriteWordLine 0 2 "View Printers and Printer Drovers"}
							Default {WriteWordLine 0 2 "Farm privileges could not be determined: $($farmprivilege)"}
						}
					}
			
					Write-Verbose "$(Get-Date): `t`t`tProcessing folder privileges"
					WriteWordLine 0 1 "Folder Privileges:"
					ForEach($folderprivilege in $Administrator.FolderPrivileges) 
					{
						Write-Verbose "$(Get-Date): `t`t`t`tProcessing folder permissions for $($FolderPrivilege.FolderPath)"
						WriteWordLine 0 2 $FolderPrivilege.FolderPath
						ForEach($FolderPermission in $FolderPrivilege.FolderPrivileges)
						{
							Switch ($folderpermission)
							{
								"Unknown"                          {WriteWordLine 0 3 "Unknown"}
								"ViewApplications"                 {WriteWordLine 0 3 "View applications"}
								"EditApplications"                 {WriteWordLine 0 3 "Edit applications"}
								"TerminateProcessApplication"      {WriteWordLine 0 3 "Terminate process that is created as a result of launching a published application"}
								"AssignApplicationsToServers"      {WriteWordLine 0 3 "Assign applications to servers"}
								"ViewServers"                      {WriteWordLine 0 3 "View servers"}
								"EditOtherServerSettings"          {WriteWordLine 0 3 "Edit other server settings"}
								"RemoveServer"                     {WriteWordLine 0 3 "Remove a bad server from farm"}
								"TerminateProcess"                 {WriteWordLine 0 3 "Terminate processes on a server"}
								"ViewSessions"                     {WriteWordLine 0 3 "View ICA/RDP sessions"}
								"ConnectSessions"                  {WriteWordLine 0 3 "Connect sessions"}
								"DisconnectSessions"               {WriteWordLine 0 3 "Disconnect sessions"}
								"LogOffSessions"                   {WriteWordLine 0 3 "Log off sessions"}
								"ResetSessions"                    {WriteWordLine 0 3 "Reset sessions"}
								"SendMessages"                     {WriteWordLine 0 3 "Send messages to sessions"}
								"ViewWorkerGroups"                 {WriteWordLine 0 3 "View worker groups"}
								"AssignApplicationsToWorkerGroups" {WriteWordLine 0 3 "Assign applications to worker groups"}
								"AssignApplications"               {WriteWordLine 0 3 "Assign Application to Servers"}
								"EditServerSnmpSettings"           {WriteWordLine 0 3 "Edit SNMP Settings"}
								"EditLicenseServer"                {WriteWordLine 0 3 "Edit License Server Settings"}
								Default {WriteWordLine 0 3 "Folder permission could not be determined: $($folderpermission)"}
							}
						}
					}
				}		
				#WriteWordLine 0 0 " "
			}
			Else
			{
				WriteWordLine 0 0 $Administrator.AdministratorName
				$TotalAdmins++
			}
		}
	}
	ElseIf(!$?)
	{
		Write-Warning "Administrator information could not be retrieved"
	}
	Else
	{
		Write-Warning "No results returned for Administrator information"
	}
	$Administrators = $Null
	Write-Verbose "$(Get-Date): Finished Processing Administrators"
	Write-Verbose "$(Get-Date): "
}

If($Section -eq "All" -or $Section -eq "Apps")
{
	Write-Verbose "$(Get-Date): Processing Applications"

	Write-Verbose "$(Get-Date): `tSetting summary variables"
	[int]$TotalPublishedApps = 0
	[int]$TotalPublishedContent = 0
	[int]$TotalPublishedDesktops = 0
	[int]$TotalStreamedApps = 0
	[int]$TotalApps = 0
	$SessionSharingItems = @()

	Write-Verbose "$(Get-Date): `tRetrieving Applications"
	If($Summary)
	{
		$Applications = Get-XAApplication -EA 0 | Sort DisplayName
	}
	Else
	{
		$Applications = Get-XAApplication -EA 0 | Sort FolderPath, DisplayName
	}

	If($? -and $Applications -ne $Null)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Applications:"
		ForEach($Application in $Applications)
		{
			Write-Verbose "$(Get-Date): `t`tProcessing application $($Application.BrowserName)"
			If(!$Summary)
			{
				If($Application.ApplicationType -ne "ServerDesktop" -and $Application.ApplicationType -ne "Content")
				{
					#create array for appendix A
					#these items are taken from http://support.citrix.com/article/CTX159159
					#Some properties that must match on both Applications for Session Sharing to Function are:
					#
					#Color depth
					#Screen Size
					#Access Control Filters (for SmartAccess)
					#Sound (unexplained in article)
					#Drive Mapping (unexplained in article)
					#Printer Mapping (unexplained in article)
					#Encryption

					Write-Verbose "$(Get-Date): `t`t`tGather session sharing info for Appendix A"
					$obj = New-Object -TypeName PSObject
					$obj | Add-Member -MemberType NoteProperty -Name ApplicationName      -Value $Application.BrowserName
					$obj | Add-Member -MemberType NoteProperty -Name MaximumColorQuality  -Value $Application.ColorDepth
					$obj | Add-Member -MemberType NoteProperty -Name SessionWindowSize    -Value $Application.WindowType

					If($Application.AccessSessionConditionsEnabled)
					{
						$tmp = @()
						ForEach($filter in $Application.AccessSessionConditions)
						{
							$tmp += $filter
						}
						$obj | Add-Member -MemberType NoteProperty -Name AccessControlFilters -Value $tmp
					}
					Else
					{
						$obj | Add-Member -MemberType NoteProperty -Name AccessControlFilters -Value "None"
					}

					$obj | Add-Member -MemberType NoteProperty -Name Encryption           -Value $Application.EncryptionLevel
					$SessionSharingItems += $obj
				}
				
				[bool]$AppServerInfoResults = $False
				$AppServerInfo = Get-XAApplicationReport -BrowserName $Application.BrowserName -EA 0
				If($?)
				{
					$AppServerInfoResults = $True
				}
				[bool]$streamedapp = $False
				If($Application.ApplicationType -Contains "streamedtoclient" -or $Application.ApplicationType -Contains "streamedtoserver")
				{
					$streamedapp = $True
				}
			}
			Else
			{
				$TotalApps++
			}
			
			#name properties
			If(!$Summary)
			{
				WriteWordLine 2 0 $Application.DisplayName
				WriteWordLine 0 1 "Application name`t`t: " $Application.BrowserName
				WriteWordLine 0 1 "Disable application`t`t: " -NoNewLine
				#weird, if application is enabled, it is disabled!
				If($Application.Enabled) 
				{
					WriteWordLine 0 0 "No"
				} 
				Else
				{
					WriteWordLine 0 0 "Yes"
					WriteWordLine 0 1 "Hide disabled application`t: " -nonewline
					If($Application.HideWhenDisabled)
					{
						WriteWordLine 0 0 "Yes"
					}
					Else
					{
						WriteWordLine 0 0 "No"
					}
				}

				If(![String]::IsNullOrEmpty($Application.Description))
				{
					WriteWordLine 0 1 "Application description`t`t: " $Application.Description
				}
				
				#type properties
				WriteWordLine 0 1 "Application Type`t`t: " -nonewline
				Switch ($Application.ApplicationType)
				{
					"Unknown"                            {WriteWordLine 0 0 "Unknown"}
					"ServerInstalled"                    {WriteWordLine 0 0 "Installed application"; $TotalPublishedApps++}
					"ServerDesktop"                      {WriteWordLine 0 0 "Server desktop"; $TotalPublishedDesktops++}
					"Content"                            {WriteWordLine 0 0 "Content"; $TotalPublishedContent++}
					"StreamedToServer"                   {WriteWordLine 0 0 "Streamed to server"; $TotalStreamedApps++}
					"StreamedToClient"                   {WriteWordLine 0 0 "Streamed to client"; $TotalStreamedApps++}
					"StreamedToClientOrInstalled"        {WriteWordLine 0 0 "Streamed if possible, otherwise accessed from server as Installed application"; $TotalStreamedApps++}
					"StreamedToClientOrStreamedToServer" {WriteWordLine 0 0 "Streamed if possible, otherwise Streamed to server"; $TotalStreamedApps++}
					Default {WriteWordLine 0 0 "Application Type could not be determined: $($Application.ApplicationType)"}
				}
				If(![String]::IsNullOrEmpty($Application.FolderPath))
				{
					WriteWordLine 0 1 "Folder path`t`t`t: " $Application.FolderPath
				}
				If(![String]::IsNullOrEmpty($Application.ContentAddress))
				{
					WriteWordLine 0 1 "Content Address`t`t: " $Application.ContentAddress
				}
			
				#if a streamed app
				If($streamedapp)
				{
					WriteWordLine 0 1 "Citrix streaming app profile address`t`t: " 
					WriteWordLine 0 2 $Application.ProfileLocation
					WriteWordLine 0 1 "App to launch from Citrix stream app profile`t: " 
					WriteWordLine 0 2 $Application.ProfileProgramName
					If(![String]::IsNullOrEmpty($Application.ProfileProgramArguments))
					{
						WriteWordLine 0 1 "Extra command line parameters`t`t`t: " 
						WriteWordLine 0 2 $Application.ProfileProgramArguments
					}
					#if streamed, Offline access properties
					If($Application.OfflineAccessAllowed)
					{
						WriteWordLine 0 1 "Enable offline access`t`t`t`t: " -nonewline
						If($Application.OfflineAccessAllowed)
						{
							WriteWordLine 0 0 "Yes"
						}
						Else
						{
							WriteWordLine 0 0 "No"
						}
					}
					If($Application.CachingOption)
					{
						WriteWordLine 0 1 "Cache preference`t`t`t`t: " -nonewline
						Switch ($Application.CachingOption)
						{
							"Unknown"   {WriteWordLine 0 0 "Unknown"}
							"PreLaunch" {WriteWordLine 0 0 "Cache application prior to launching"}
							"AtLaunch"  {WriteWordLine 0 0 "Cache application during launch"}
							Default {WriteWordLine 0 0 "Could not be determined: $($Application.CachingOption)"}
						}
					}
				}
				
				#location properties
				If(!$streamedapp)
				{
					#requested by Pavel Stadler to put Command Line and Working Directory in a different sized font and make it bold
					If(![String]::IsNullOrEmpty($Application.CommandLineExecutable))
					{
						If($Application.CommandLineExecutable.Length -lt 40)
						{
							WriteWordLine 0 1 "Command Line`t`t`t: " -NoNewLine
							WriteWordLine 0 0 $Application.CommandLineExecutable "" "Courier New" 9 $False $True
						}
						Else
						{
							WriteWordLine 0 1 "Command Line: " 
							WriteWordLine 0 2 $Application.CommandLineExecutable "" "Courier New" 9 $False $True
						}
					}
					If(![String]::IsNullOrEmpty($Application.WorkingDirectory))
					{
						If($Application.WorkingDirectory.Length -lt 40)
						{
							WriteWordLine 0 1 "Working directory`t`t: " -NoNewLine
							WriteWordLine 0 0 $Application.WorkingDirectory "" "Courier New" 9 $False $True
						}
						Else
						{
							WriteWordLine 0 1 "Working directory: " 
							WriteWordLine 0 2 $Application.WorkingDirectory "" "Courier New" 9 $False $True
						}
					}
					
					#servers properties
					If($AppServerInfoResults)
					{
						If(![String]::IsNullOrEmpty($AppServerInfo.ServerNames))
						{
							If(![String]::IsNullOrEmpty($AppServerInfo.ServerNames))
							{
								WriteWordLine 0 1 "Servers:"
								$TempArray = $AppServerInfo.ServerNames | Sort
								BuildTableForServer $TempArray
								$TempArray = $Null
							}
						}
					}
					Else
					{
						WriteWordLine 0 2 "Unable to retrieve a list of Servers for this application"
					}
				}
			
				#users properties
				If($Application.AnonymousConnectionsAllowed)
				{
					WriteWordLine 0 1 "Allow anonymous users: " $Application.AnonymousConnectionsAllowed
				}
				Else
				{
					If($AppServerInfoResults)
					{
						WriteWordLine 0 1 "Users:"
						ForEach($user in $AppServerInfo.Accounts)
						{
							WriteWordLine 0 2 $user
						}
					}
					Else
					{
						WriteWordLine 0 2 "Unable to retrieve a list of Users for this application"
					}
				}
			
				#shortcut presentation properties
				#application icon is ignored
				If(![String]::IsNullOrEmpty($Application.ClientFolder))
				{
					If($Application.ClientFolder.Length -lt 30)
					{
						WriteWordLine 0 1 "Client application folder`t`t`t`t: " $Application.ClientFolder
					}
					Else
					{
						WriteWordLine 0 1 "Client application folder`t`t`t`t: " 
						WriteWordLine 0 2 $Application.ClientFolder
					}
				}
				If($Application.AddToClientStartMenu)
				{
					WriteWordLine 0 1 "Add to client's start menu"
					If($Application.StartMenuFolder)
					{
						WriteWordLine 0 2 "Start menu folder`t`t`t: " $Application.StartMenuFolder
					}
				}
				If($Application.AddToClientDesktop)
				{
					WriteWordLine 0 1 "Add shortcut to the client's desktop"
				}
			
				#access control properties
				If($Application.ConnectionsThroughAccessGatewayAllowed)
				{
					WriteWordLine 0 1 "Allow connections made through AGAE`t`t: " -nonewline
					If($Application.ConnectionsThroughAccessGatewayAllowed)
					{
						WriteWordLine 0 0 "Yes"
					} 
					Else
					{
						WriteWordLine 0 0 "No"
					}
				}
				If($Application.OtherConnectionsAllowed)
				{
					WriteWordLine 0 1 "Any connection`t`t`t`t`t: " -nonewline
					If($Application.OtherConnectionsAllowed)
					{
						WriteWordLine 0 0 "Yes"
					} 
					Else
					{
						WriteWordLine 0 0 "No"
					}
				}
				If($Application.AccessSessionConditionsEnabled)
				{
					WriteWordLine 0 1 "Any connection that meets any of the following filters: " $Application.AccessSessionConditionsEnabled
					WriteWordLine 0 1 "Access Gateway Filters:"
					$TableRange = $doc.Application.Selection.Range
					[int]$Columns = 2
					[int]$Rows = $Application.AccessSessionConditions.count + 1
					$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
					$Table.rows.first.headingformat = $wdHeadingFormatTrue
					$Table.Style = $myHash.Word_TableGrid
					$Table.Borders.InsideLineStyle = $wdLineStyleSingle
					$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
					[int]$xRow = 1
					Write-Verbose "$(Get-Date): `t`t`t`tFormat first row with column headings"
					$Table.Rows.First.Shading.BackgroundPatternColor = $wdColorGray15
					$Table.Cell($xRow,1).Range.Font.Bold = $True
					$Table.Cell($xRow,1).Range.Text = "Farm Name"
					$Table.Cell($xRow,2).Range.Font.Bold = $True
					$Table.Cell($xRow,2).Range.Text = "Filter"
					ForEach($AccessCondition in $Application.AccessSessionConditions)
					{
						[string]$Tmp = $AccessCondition
						[string]$AGFarm = $Tmp.substring(0, $Tmp.indexof(":"))
						[string]$AGFilter = $Tmp.substring($Tmp.indexof(":")+1)
						$xRow++
						Write-Verbose "$(Get-Date): `t`t`t`t`tProcessing row for Access Condition $($Tmp)"
						$Table.Cell($xRow,1).Range.Text = $AGFarm
						$Table.Cell($xRow,2).Range.Text = $AGFilter
					}

					$Table.Rows.SetLeftIndent($Indent2TabStops,$wdAdjustProportional)
					$Table.AutoFitBehavior($wdAutoFitContent)

					FindWordDocumentEnd
					$tmp = $Null
					$AGFarm = $Null
					$AGFilter = $Null
				}
			
				#content redirection properties
				If($AppServerInfoResults)
				{
					If($AppServerInfo.FileTypes)
					{
						WriteWordLine 0 1 "File type associations:"
						ForEach($filetype in $AppServerInfo.FileTypes)
						{
							WriteWordLine 0 3 $filetype
						}
					}
					Else
					{
						WriteWordLine 0 1 "File Type Associations for this application`t: None"
					}
				}
				Else
				{
					WriteWordLine 0 1 "Unable to retrieve the list of FTAs for this application"
				}
			
				#if streamed app, Alternate profiles
				If($streamedapp)
				{
					If($Application.AlternateProfiles)
					{
						WriteWordLine 0 1 "Primary application profile location`t`t: " $Application.AlternateProfiles
					}
				
					#if streamed app, User privileges properties
					If($Application.RunAsLeastPrivilegedUser)
					{
						WriteWordLine 0 1 "Run application as a least-privileged user account`t: " $Application.RunAsLeastPrivilegedUser
					}
				}
			
				#limits properties
				WriteWordLine 0 1 "Limit instances allowed to run in server farm`t: " -NoNewLine

				If($Application.InstanceLimit -eq -1)
				{
					WriteWordLine 0 0 "No limit set"
				}
				Else
				{
					WriteWordLine 0 0 $Application.InstanceLimit
				}
			
				WriteWordLine 0 1 "Allow only 1 instance of app for each user`t: " -NoNewLine
			
				If($Application.MultipleInstancesPerUserAllowed) 
				{
					WriteWordLine 0 0 "No"
				} 
				Else
				{
					WriteWordLine 0 0 "Yes"
				}
			
				If($Application.CpuPriorityLevel)
				{
					WriteWordLine 0 1 "Application importance`t`t`t`t: " -nonewline
					Switch ($Application.CpuPriorityLevel)
					{
						"Unknown"     {WriteWordLine 0 0 "Unknown"}
						"BelowNormal" {WriteWordLine 0 0 "Below Normal"}
						"Low"         {WriteWordLine 0 0 "Low"}
						"Normal"      {WriteWordLine 0 0 "Normal"}
						"AboveNormal" {WriteWordLine 0 0 "Above Normal"}
						"High"        {WriteWordLine 0 0 "High"}
						Default {WriteWordLine 0 0 "Application importance could not be determined: $($Application.CpuPriorityLevel)"}
					}
				}
				
				#client options properties
				WriteWordLine 0 1 "Enable legacy audio`t`t`t`t: " -nonewline
				Switch ($Application.AudioType)
				{
					"Unknown" {WriteWordLine 0 0 "Unknown"}
					"None"    {WriteWordLine 0 0 "Not Enabled"}
					"Basic"   {WriteWordLine 0 0 "Enabled"}
					Default {WriteWordLine 0 0 "Enable legacy audio could not be determined: $($Application.AudioType)"}
				}
				WriteWordLine 0 1 "Minimum requirement`t`t`t`t: " -nonewline
				If($Application.AudioRequired)
				{
					WriteWordLine 0 0 "Enabled"
				}
				Else
				{
					WriteWordLine 0 0 "Disabled"
				}
				If($Application.SslConnectionEnabled)
				{
					WriteWordLine 0 1 "Enable SSL and TLS protocols`t`t`t: " -nonewline
					If($Application.SslConnectionEnabled)
					{
						WriteWordLine 0 0 "Enabled"
					}
					Else
					{
						WriteWordLine 0 0 "Disabled"
					}
				}
				If($Application.EncryptionLevel)
				{
					WriteWordLine 0 1 "Encryption`t`t`t`t`t: " -nonewline
					Switch ($Application.EncryptionLevel)
					{
						"Unknown" {WriteWordLine 0 0 "Unknown"}
						"Basic"   {WriteWordLine 0 0 "Basic"}
						"LogOn"   {WriteWordLine 0 0 "128-Bit Login Only (RC-5)"}
						"Bits40"  {WriteWordLine 0 0 "40-Bit (RC-5)"}
						"Bits56"  {WriteWordLine 0 0 "56-Bit (RC-5)"}
						"Bits128" {WriteWordLine 0 0 "128-Bit (RC-5)"}
						Default {WriteWordLine 0 0 "Encryption could not be determined: $($Application.EncryptionLevel)"}
					}
				}
				If($Application.EncryptionRequired)
				{
					WriteWordLine 0 1 "Minimum requirement`t`t`t`t: " -nonewline
					If($Application.EncryptionRequired)
					{
						WriteWordLine 0 0 "Enabled"
					}
					Else
					{
						WriteWordLine 0 0 "Disabled"
					}
				}
			
				WriteWordLine 0 1 "Start app w/o waiting for printer creation`t: " -NoNewLine
				#another weird one, if True then this is Disabled
				If($Application.WaitOnPrinterCreation) 
				{
					WriteWordLine 0 0 "No"
				} 
				Else
				{
					WriteWordLine 0 0 "Yes"
				}
				
				#appearance properties
				If($Application.WindowType)
				{
					WriteWordLine 0 1 "Session window size`t`t`t`t: " $Application.WindowType
				}
				If($Application.ColorDepth)
				{
					WriteWordLine 0 1 "Maximum color quality`t`t`t`t: " -nonewline
					Switch ($Application.ColorDepth)
					{
						"Colors16"  {WriteWordLine 0 0 "16 colors"}
						"Colors256" {WriteWordLine 0 0 "256 colors"}
						"HighColor" {WriteWordLine 0 0 "High Color (16-bit)"}
						"TrueColor" {WriteWordLine 0 0 "True Color (24-bit)"}
						Default {WriteWordLine 0 0 "Maximum color quality could not be determined: $($Application.ColorDepth)"}
					}
				}
				If($Application.TitleBarHidden)
				{
					WriteWordLine 0 1 "Hide application title bar`t`t`t: " -nonewline
					If($Application.TitleBarHidden)
					{
						WriteWordLine 0 0 "Enabled"
					}
					Else
					{
						WriteWordLine 0 0 "Disabled"
					}
				}
				If($Application.MaximizedOnStartup)
				{
					WriteWordLine 0 1 "Maximize application at startup`t`t`t: " -nonewline
					If($Application.MaximizedOnStartup)
					{
						WriteWordLine 0 0 "Enabled"
					}
					Else
					{
						WriteWordLine 0 0 "Disabled"
					}
				}
				$AppServerInfo = $Null
			}
			Else
			{
				WriteWordLine 0 0 $Application.DisplayName
			}
		}
	}
	ElseIf($Applications -eq $Null)
	{
		Write-Verbose "$(Get-Date): There are no Applications published"
	}
	Else 
	{
		Write-Warning "No results returned for Application information"
	}
	$Applications = $Null
	Write-Verbose "$(Get-Date): Finished Processing Applications"
	Write-Verbose "$(Get-Date): "
}

If($Section -eq "All" -or $Section -eq "Servers")
{
	#servers
	Write-Verbose "$(Get-Date): Processing Servers"
	Write-Verbose "$(Get-Date): `tSetting summary variables"
	[int]$TotalServers = 0
	$ServerItems = @()

	Write-Verbose "$(Get-Date): `tRetrieving Servers"
	If($Summary)
	{
		$servers = Get-XAServer -EA 0 | Sort ServerName
	}
	Else
	{
		$servers = Get-XAServer -EA 0 | Sort FolderPath, ServerName
	}

	If($? -and $Servers -ne $Null)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Servers:"
		ForEach($server in $servers)
		{
			$TotalServers++
			Write-Verbose "$(Get-Date): `t`tProcessing server $($server.ServerName)"
			
			If(!$Summary)
			{
				[bool]$SvrOnline = $False
				Write-Verbose "$(Get-Date): `t`t`tTesting to see if $($server.ServerName) is online and reachable"
				If(Test-Connection -ComputerName $server.servername -quiet -EA 0)
				{
					$SvrOnline = $True
					If($Hardware -and $Software)
					{
						Write-Verbose "$(Get-Date): `t`t`t`t$($server.ServerName) is online.  Hardware inventory, Software Inventory, Citrix Services and Hotfix areas will be processed."
					}
					ElseIf($Hardware -and !($Software))
					{
						Write-Verbose "$(Get-Date): `t`t`t`t$($server.ServerName) is online.  Hardware inventory, Citrix Services and Hotfix areas will be processed."
					}
					ElseIf(!($Hardware) -and $Software)
					{
						Write-Verbose "$(Get-Date): `t`t`t`t$($server.ServerName) is online.  Software Inventory, Citrix Services and Hotfix areas will be processed."
					}
					Else
					{
						Write-Verbose "$(Get-Date): `t`t`t`t$($server.ServerName) is online.  Citrix Services and Hotfix areas will be processed."
					}
				}

				WriteWordLine 2 0 $server.ServerName
				WriteWordLine 0 1 "Product`t`t`t`t: " $server.CitrixProductName
				WriteWordLine 0 1 "Edition`t`t`t`t: " $server.CitrixEdition
				WriteWordLine 0 1 "Version`t`t`t`t: " $server.CitrixVersion
				WriteWordLine 0 1 "Service Pack`t`t`t: " $server.CitrixServicePack
				WriteWordLine 0 1 "Operating System Type`t`t: " -NoNewLine
				If($server.Is64Bit)
				{
					WriteWordLine 0 0 "64 bit"
				} 
				Else 
				{
					WriteWordLine 0 0 "32 bit"
				}
				WriteWordLine 0 1 "IP Address`t`t`t: " $server.IPAddresses
				WriteWordLine 0 1 "Logons`t`t`t`t: " -NoNewLine
				If($server.LogOnsEnabled)
				{
					WriteWordLine 0 0 "Enabled"
				} 
				Else 
				{
					WriteWordLine 0 0 "Disabled"
				}
				WriteWordLine 0 1 "Product Installation Date`t: " $server.CitrixInstallDate
				WriteWordLine 0 1 "Operating System Version`t: " $server.OSVersion -NoNewLine
				
				#is the server running server 2008?
				If($server.OSVersion.ToString().SubString(0,1) -eq "6")
				{
					$Server2008 = $True
				}

				WriteWordLine 0 0 " " $server.OSServicePack
				WriteWordLine 0 1 "Zone`t`t`t`t: " $server.ZoneName
				WriteWordLine 0 1 "Election Preference`t`t: " -nonewline
				Switch ($server.ElectionPreference)
				{
					"Unknown"           {WriteWordLine 0 0 "Unknown"}
					"MostPreferred"     {WriteWordLine 0 0 "Most Preferred"}
					"Preferred"         {WriteWordLine 0 0 "Preferred"}
					"DefaultPreference" {WriteWordLine 0 0 "Default Preference"}
					"NotPreferred"      {WriteWordLine 0 0 "Not Preferred"}
					"WorkerMode"        {WriteWordLine 0 0 "Worker Mode"}
					Default {WriteWordLine 0 0 "Server election preference could not be determined: $($server.ElectionPreference)"}
				}
				WriteWordLine 0 1 "Folder`t`t`t`t: " $server.FolderPath
				WriteWordLine 0 1 "Product Installation Path`t: " $server.CitrixInstallPath
				If($server.ICAPortNumber -gt 0)
				{
					WriteWordLine 0 1 "ICA Port Number`t`t: " $server.ICAPortNumber
				}
				$ServerConfig = Get-XAServerConfiguration -ServerName $Server.ServerName -EA 0
				If($?)
				{
					WriteWordLine 0 1 "Server Configuration Data:"
					
					$Tmp = "ICA"
					
					If($ServerConfig.AcrUseFarmSettings)
					{
						WriteWordLine 0 2 "$($Tmp)\Auto Client Reconnect: Server is using farm settings"
					}
					Else
					{
						WriteWordLine 0 2 "$($Tmp)\Auto Client Reconnect: Server is not using farm settings"
						If($ServerConfig.AcrEnabled)
						{
							WriteWordLine 0 3 "Reconnect automatically"
							WriteWordLine 0 4 "Log automatic reconnection attempts: " -nonewline
							If($ServerConfig.AcrLogReconnections)
							{
								WriteWordLine 0 0 "Yes"
							}
							Else
							{
								WriteWordLine 0 0 "No"
							}
						}
						Else
						{
							WriteWordLine 0 3 "Require user authentication"
						}
					}
					WriteWordLine 0 2 "$($Tmp)\Browser\Create browser listener on UDP network: " -nonewline
					If($ServerConfig.BrowserUdpListener)
					{
						WriteWordLine 0 0 "Yes"
					}
					Else
					{
						WriteWordLine 0 0 "No"
					}
					WriteWordLine 0 2 "$($Tmp)\Browser\Server responds to client broadcast messages: " -nonewline
					If($ServerConfig.BrowserRespondToClientBroadcasts)
					{
						WriteWordLine 0 0 "Yes"
					}
					Else
					{
						WriteWordLine 0 0 "No"
					}
					If($ServerConfig.DisplayUseFarmSettings)
					{
						WriteWordLine 0 2 "$($Tmp)\Display: Server is using farm settings"
					}
					Else
					{
						WriteWordLine 0 2 "$($Tmp)\Display: Server is not using farm settings"
						WriteWordLine 0 3 "Discard queued image that is replaced by another image: " -nonewline
						If($ServerConfig.DisplayDiscardQueuedImages)
						{
							WriteWordLine 0 0 "Yes"
						}
						Else
						{
							WriteWordLine 0 0 "No"
						}
						WriteWordLine 0 3 "Cache image to make scrolling smoother: " -nonewline
						If($ServerConfig.DisplayCacheImageForSmoothScrolling)
						{
							WriteWordLine 0 0 "Yes"
						}
						Else
						{
							WriteWordLine 0 0 "No"
						}
						WriteWordLine 0 3 "Maximum memory to use for each session's graphics (KB): " $ServerConfig.DisplayMaximumGraphicsMemory
						WriteWordLine 0 3 "Degradation bias: " 
						If($ServerConfig.DisplayDegradationBias -eq "Resolution")
						{
							WriteWordLine 0 4 "Degrade resolution first"
						}
						Else
						{
							WriteWordLine 0 4 "Degrade color depth first"
						}
						WriteWordLine 0 3 "Notify user of session degradation: " -nonewline
						If($ServerConfig.DisplayNotifyUser)
						{
							WriteWordLine 0 0 "Yes"
						}
						Else
						{
							WriteWordLine 0 0 "No"
						}
					}
					If($ServerConfig.KeepAliveUseFarmSettings)
					{
						WriteWordLine 0 2 "$($Tmp)\Keep-Alive: Server is using farm settings"
					}
					Else
					{
						WriteWordLine 0 2 "$($Tmp)\Keep-Alive: Server is not using farm settings"
						WriteWordLine 0 3 "ICA Keep-Alive time-out value seconds: " -NoNewLine
						If($ServerConfig.KeepAliveEnabled)
						{
							WriteWordLine 0 0 $ServerConfig.KeepAliveTimeout
						}
						Else
						{
							WriteWordLine 0 0 "Disabled"
						}
					}
					If($ServerConfig.PrinterBandwidth -eq -1)
					{
						WriteWordLine 0 2 "$($Tmp)\Printer Bandwidth\Unlimited client printer bandwidth"
					}
					Else
					{
						WriteWordLine 0 2 "$($Tmp)\Printer Bandwidth\Limit bandwidth to use (kbps): " $ServerConfig.PrinterBandwidth
					}

					If($ServerConfig.LicenseServerUseFarmSettings)
					{
						WriteWordLine 0 2 "License Server: Server is using farm settings"
					}
					Else
					{
						WriteWordLine 0 2 "License Server: Server is not using farm settings"
						WriteWordLine 0 3 "License server name: " $ServerConfig.LicenseServerName
						WriteWordLine 0 3 "License server port: " $ServerConfig.LicenseServerPortNumber
					}
					If($ServerConfig.HmrUseFarmSettings)
					{
						WriteWordLine 0 2 "Health Monitoring & Recovery: Server is using farm settings"
					}
					Else
					{
						WriteWordLine 0 2 "Health Monitoring & Recovery: Server is not using farm settings"
						WriteWordLine 0 3 "Run health monitoring tests on this server: " -nonewline
						If($ServerConfig.HmrEnabled)
						{
							WriteWordLine 0 0 "Yes"
						}
						Else
						{
							WriteWordLine 0 0 "No"
						}
						If($ServerConfig.HmrEnabled)
						{
							$HMRTests = Get-XAHmrTest -ServerName $Server.ServerName -EA 0
							If($?)
							{
								WriteWordLine 0 3 "Health Monitoring Tests:"
								ForEach($HMRTest in $HMRTests)
								{
									WriteWordLine 0 4 "Test Name`t: " $Hmrtest.TestName
									WriteWordLine 0 4 "Interval`t`t: " $Hmrtest.Interval
									WriteWordLine 0 4 "Threshold`t: " $Hmrtest.Threshold
									WriteWordLine 0 4 "Time-out`t: " $Hmrtest.Timeout
									WriteWordLine 0 4 "Test File Name`t: " $Hmrtest.FilePath
									If(![String]::IsNullOrEmpty($Hmrtest.Arguments))
									{
										WriteWordLine 0 4 "Arguments`t: " $Hmrtest.Arguments
									}
									WriteWordLine 0 4 "Recovery Action : " -nonewline
									Switch ($Hmrtest.RecoveryAction)
									{
										"AlertOnly"                     {WriteWordLine 0 0 "Alert Only"}
										"RemoveServerFromLoadBalancing" {WriteWordLine 0 0 "Remove Server from load balancing"}
										"RestartIma"                    {WriteWordLine 0 0 "Restart IMA"}
										"ShutdownIma"                   {WriteWordLine 0 0 "Shutdown IMA"}
										"RebootServer"                  {WriteWordLine 0 0 "Reboot Server"}
										Default {WriteWordLine 0 0 "Recovery Action could not be determined: $($Hmrtest.RecoveryAction)"}
									}
									WriteWordLine 0 0 ""
								}
							}
							Else
							{
								WriteWordLine 0 0 "Health Monitoring & Reporting data could not be retrieved for server " $Server.ServerName
							}
						}
					}
					If($ServerConfig.CpuUseFarmSettings)
					{
						WriteWordLine 0 2 "CPU Utilization Management: Server is using farm settings"
					}
					Else
					{
						WriteWordLine 0 2 "CPU Utilization Management: Server is not using farm settings"
						WriteWordLine 0 3 "CPU Utilization Management: " -nonewline
						Switch ($ServerConfig.CpuManagementLevel)
						{
							"NoManagement"  {WriteWordLine 0 0 "No CPU utilization management"}
							"Fair"          {WriteWordLine 0 0 "Fair sharing of CPU between sessions"}
							"ResourceBased" {WriteWordLine 0 0 "CPU Sharing based on Resource Allotments"}
							Default {WriteWordLine 0 0 "CPU Utilization Management could not be determined: $($Farm.CpuManagementLevel)"}
						}
					}
					If($ServerConfig.MemoryUseFarmSettings)
					{
						WriteWordLine 0 2 "Memory Optimization: Server is using farm settings"
					}
					Else
					{
						WriteWordLine 0 2 "Memory Optimization: Server is not using farm settings"
						WriteWordLine 0 3 "Memory Optimization: " -nonewline
						If($ServerConfig.MemoryOptimizationEnabled)
						{
							WriteWordLine 0 0 "Enabled"
						}
						Else
						{
							WriteWordLine 0 0 "Not Enabled"
						}
					}
					
					$Tmp = "XenApp"
					
					If($ServerConfig.ContentRedirectionUseFarmSettings)
					{
						WriteWordLine 0 2 "$($Tmp)/Content Redirection: Server is using farm settings"
					}
					Else
					{
						WriteWordLine 0 2 "$($Tmp)/Content Redirection: Server is not using farm settings"
						WriteWordLine 0 3 "Content redirection from server to client: " -nonewline
						If($ServerConfig.ContentRedirectionEnabled)
						{
							WriteWordLine 0 0 "Yes"
						}
						Else
						{
							WriteWordLine 0 0 "No"
						}
					}
					#ShadowLoggingEnabled is not stored by Citrix
					#WriteWordLine 0 3 "HDX Plug and Play/Shadow Logging/Log shadowing sessions: " $ServerConfig.ShadowLoggingEnabled
					
					If($ServerConfig.RemoteConsoleUseFarmSettings)
					{
						WriteWordLine 0 2 "$($Tmp)\Remote Console Connections: Server is using farm settings"
					}
					Else
					{
						WriteWordLine 0 2 "$($Tmp)\Remote Console Connections: Server is not using farm settings"
						WriteWordLine 0 3 "Remote connections to the console: " -nonewline
						If($ServerConfig.RemoteConsoleEnabled)
						{
							WriteWordLine 0 0 "Yes"
						}
						Else
						{
							WriteWordLine 0 0 "No"
						}
					}

					If($ServerConfig.SnmpUseFarmSettings)
					{
						WriteWordLine 0 2 "SNMP: Server is using farm settings"
					}
					Else
					{
						WriteWordLine 0 2 "SNMP: Server is not using farm settings"
						# SnmpEnabled is not working
						WriteWordLine 0 3 "Send session traps to selected SNMP agent on all farm servers"
						WriteWordLine 0 4 "SNMP agent session traps:"
						WriteWordLine 0 5 "Logon`t`t`t: " -nonewline
						If($ServerConfig.SnmpLogOnEnabled)
						{
							WriteWordLine 0 0 "Yes"
						}
						Else
						{
							WriteWordLine 0 0 "No"
						}
						WriteWordLine 0 5 "Logoff`t`t`t: " -nonewline
						If($ServerConfig.SnmpLogOffEnabled)
						{
							WriteWordLine 0 0 "Yes"
						}
						Else
						{
							WriteWordLine 0 0 "No"
						}
						WriteWordLine 0 5 "Disconnect`t`t: " -nonewline
						If($ServerConfig.SnmpDisconnectEnabled)
						{
							WriteWordLine 0 0 "Yes"
						}
						Else
						{
							WriteWordLine 0 0 "No"
						}
						WriteWordLine 0 5 "Session limit per server`t: " -nonewline
						If($ServerConfig.SnmpLimitEnabled)
						{
							WriteWordLine 0 0 " " $ServerConfig.SnmpLimitPerServer
						}
						Else
						{
							WriteWordLine 0 0 "No"
						}
					}
					
					$Tmp = "SpeedScreen"
					
					If($ServerConfig.BrowserAccelerationUseFarmSettings)
					{
						WriteWordLine 0 2 "$($Tmp)/Browser Acceleration: Server is using farm settings"
					}
					Else
					{
						WriteWordLine 0 2 "$($Tmp)/Browser Acceleration: Server is not using farm settings"
						WriteWordLine 0 3 "$($Tmp)/$($Tmp)Browser Acceleration: " -nonewline
						If($ServerConfig.BrowserAccelerationEnabled)
						{
							WriteWordLine 0 0 "Yes"
						}
						Else
						{
							WriteWordLine 0 0 "No"
						}
						If($ServerConfig.BrowserAccelerationEnabled)
						{
							WriteWordLine 0 4 "Compress JPEG images to improve bandwidth: " -nonewline
							If($ServerConfig.BrowserAccelerationCompressionEnabled)
							{
								WriteWordLine 0 0 "Yes"
							}
							Else
							{
								WriteWordLine 0 0 "No"
							}
							If($ServerConfig.BrowserAccelerationCompressionEnabled)
							{
								WriteWordLine 0 4 "Image compression level: " $ServerConfig.BrowserAccelerationCompressionLevel
								WriteWordLine 0 4 "Adjust compression level based on available bandwidth: " -nonewline
								If($ServerConfig.BrowserAccelerationVariableImageCompression)
								{
									WriteWordLine 0 0 "Yes"
								}
								Else
								{
									WriteWordLine 0 0 "No"
								}
							}
						}
					}
					
					$Tmp = "SpeedScreen"
					
					If($ServerConfig.FlashAccelerationUseFarmSettings)
					{
						WriteWordLine 0 2 "$($Tmp)/Flash: Server is using farm settings"
					}
					Else
					{
						WriteWordLine 0 2 "$($Tmp)/Flash: Server is not using farm settings"
						WriteWordLine 0 3 "Enable Adobe Flash Player: " -nonewline
						If($ServerConfig.FlashAccelerationEnabled)
						{
							WriteWordLine 0 0 "Yes"
						}
						Else
						{
							WriteWordLine 0 0 "No"
						}
						If($ServerConfig.FlashAccelerationEnabled)
						{
							Switch ($ServerConfig.FlashAccelerationOption)
							{
								"RestrictedBandwidth" {WriteWordLine 0 3 "Restricted bandwidth connections"}
								"NoOptimization"      {WriteWordLine 0 3 "Do not optimize"}
								"AllConnections"      {WriteWordLine 0 3 "All connections"}
								Default {WriteWordLine 0 0 "Server-side acceleration could not be determined: $($ServerConfig.FlashAccelerationOption)"}
							}
						}
					}
					If($ServerConfig.MultimediaAccelerationUseFarmSettings)
					{
						WriteWordLine 0 2 "$($Tmp)/Multimedia Acceleration: Server is using farm settings"
					}
					Else
					{
						WriteWordLine 0 2 "$($Tmp)/Multimedia Acceleration: Server is not using farm settings"
						WriteWordLine 0 3 "Multimedia acceleration: " -nonewline
						If($ServerConfig.MultimediaAccelerationEnabled)
						{
							WriteWordLine 0 0 "Yes"
						}
						Else
						{
							WriteWordLine 0 0 "No"
						}
						If($ServerConfig.MultimediaAccelerationEnabled)
						{
							If($ServerConfig.MultimediaAccelerationDefaultBuffer)
							{
								WriteWordLine 0 3 "Use the Default buffer of 5 seconds"
							}
							Else
							{
								WriteWordLine 0 3 "Custom buffer time in seconds: " $ServerConfig.MultimediaAccelerationCustomBuffer
							}
						}
					}
					WriteWordLine 0 2 "Virtual IP/Enable virtual IP for this server: " -nonewline
					If($ServerConfig.VirtualIPEnabled)
					{
						WriteWordLine 0 0 "Yes"
					}
					Else
					{
						WriteWordLine 0 0 "No"
					}
					WriteWordLine 0 2 "Virtual IP/Use farm setting for IP address logging: " -nonewline
					If($ServerConfig.VirtualIPUseFarmLoggingSettings)
					{
						WriteWordLine 0 0 "Yes"
					}
					Else
					{
						WriteWordLine 0 0 "No"
					}
					WriteWordLine 0 2 "Virtual IP/Enable logging of IP address assignment and release: " -nonewline
					If($ServerConfig.VirtualIPLoggingEnabled)
					{
						WriteWordLine 0 0 "Yes"
					}
					Else
					{
						WriteWordLine 0 0 "No"
					}
					WriteWordLine 0 2 "Virtual IP/Enable virtual loopback for this server: " -nonewline
					If($ServerConfig.VirtualIPLoopbackEnabled)
					{
						WriteWordLine 0 0 "Yes"
					}
					Else
					{
						WriteWordLine 0 0 "No"
					}
					WriteWordLine 0 2 "XML Service/Trust requests sent to the XML service: " -nonewline
					If($ServerConfig.XmlServiceTrustRequests)
					{
						WriteWordLine 0 0 "Yes"
					}
					Else
					{
						WriteWordLine 0 0 "No"
					}
					If($Server2008)
					{
						If($ServerConfig.RestartsEnabled)
						{
							WriteWordLine 0 2 "Automatic restarts are enabled"
							WriteWordLine 0 3 "Restart server from: " $ServerConfig.RestartFrom
							WriteWordLine 0 3 "Restart frequency in days: " $ServerConfig.RestartFrequency
						}
						Else
						{
							WriteWordLine 0 2 "Automatic restarts are not enabled"
						}
					}
					WriteWordLine 0 0 ""
				}
				Else
				{
					WriteWordLine 0 0 "Server configuration data could not be retrieved for server " $Server.ServerName
				}

				#create array for appendix B
				#this cannot be at the top of the loop like the other scripts
				#license server info is retrieved differently than other scripts
				Write-Verbose "$(Get-Date): `t`tGather server info for Appendix B"
				$obj = New-Object -TypeName PSObject
				$obj | Add-Member -MemberType NoteProperty -Name ServerName -Value $server.ServerName
				$obj | Add-Member -MemberType NoteProperty -Name ZoneName -Value $server.ZoneName
				$obj | Add-Member -MemberType NoteProperty -Name OSVersion -Value $server.OSVersion
				$obj | Add-Member -MemberType NoteProperty -Name CitrixVersion -Value $server.CitrixVersion
				$obj | Add-Member -MemberType NoteProperty -Name ProductEdition -Value $server.CitrixEdition
				
				If($ServerConfig.LicenseServerUseFarmSettings)
				{
					$obj | Add-Member -MemberType NoteProperty -Name LicenseServer -Value "Farm Setting"
				}
				Else
				{
					$obj | Add-Member -MemberType NoteProperty -Name LicenseServer -Value $ServerConfig.LicenseServerName
				}

				If($SvrOnline)
				{
					$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $server.ServerName)
					$RegKey= $Reg.OpenSubKey("SYSTEM\\CurrentControlSet\\Control\\Citrix\\Wfshell\\TWI")
					$SSDisabled = $RegKey.GetValue("SeamlessFlags")
					
					If($SSDisabled -eq 1)
					{
						$obj | Add-Member -MemberType NoteProperty -Name SessionSharing -Value "Disabled"
					}
					Else
					{
						$obj | Add-Member -MemberType NoteProperty -Name SessionSharing -Value "Enabled"
					}
				}
				Else
				{
					$obj | Add-Member -MemberType NoteProperty -Name SessionSharing -Value "Server Offline"
				}
				
				$ServerItems += $obj

				If($SvrOnline -and $Hardware)
				{
					GetComputerWMIInfo $server.ServerName
				}
				
				#applications published to server
				$Applications = Get-XAApplication -ServerName $server.ServerName -EA 0 | Sort FolderPath, DisplayName
				If($? -and $Applications)
				{
					WriteWordLine 0 1 "Published applications:"
					Write-Verbose "$(Get-Date): `t`tProcessing published applications for server $($server.ServerName)"
					Write-Verbose "$(Get-Date): `t`tCreate Word Table for server's published applications"
					$TableRange = $doc.Application.Selection.Range
					[int]$Columns = 2
					
					If($Applications -is [Array])
					{
						[int]$Rows = $Applications.count + 1
					}
					Else
					{
						[int]$Rows = 2
					}

					$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
					$Table.rows.first.headingformat = $wdHeadingFormatTrue
					$Table.Style = $myHash.Word_TableGrid
					$Table.Borders.InsideLineStyle = $wdLineStyleSingle
					$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
					[int]$xRow = 1
					Write-Verbose "$(Get-Date): `t`tFormat first row with column headings"
					$Table.Rows.First.Shading.BackgroundPatternColor = $wdColorGray15
					$Table.Cell($xRow,1).Range.Font.Bold = $True
					$Table.Cell($xRow,1).Range.Text = "Display name"
					$Table.Cell($xRow,2).Range.Font.Bold = $True
					$Table.Cell($xRow,2).Range.Text = "Folder path"
					ForEach($app in $Applications)
					{
						Write-Verbose "$(Get-Date): `t`t`tProcessing published application $($app.DisplayName)"
						$xRow++
						$Table.Cell($xRow,1).Range.Text = $app.DisplayName
						$Table.Cell($xRow,2).Range.Text = $app.FolderPath
					}
					Write-Verbose "$(Get-Date): `t`tMove table of published applications to the right"
					$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)
					$Table.AutoFitBehavior($wdAutoFitContent)

					FindWordDocumentEnd
					WriteWordLine 0 0 ""
				}

				#get list of applications installed on server
				# original work by Shaun Ritchie
				# modified by Jeff Wouters
				# modified by Webster
				# fixed, as usual, by Michael B. Smith
				If($SvrOnline -and $Software)
				{
					#section modified on 3-jan-2014 to add displayversion
					$InstalledApps = @()
					$JustApps = @()

					#Define the variable to hold the location of Currently Installed Programs
					$UninstallKey1="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall" 

					#Create an instance of the Registry Object and open the HKLM base key
					$reg=[microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$Server.ServerName) 

					#Drill down into the Uninstall key using the OpenSubKey Method
					$regkey1=$reg.OpenSubKey($UninstallKey1) 

					#Retrieve an array of string that contain all the subkey names
					If($regkey1 -ne $Null)
					{
						$subkeys1=$regkey1.GetSubKeyNames() 

						#Open each Subkey and use GetValue Method to return the required values for each
						ForEach($key in $subkeys1) 
						{
							$thisKey=$UninstallKey1+"\\"+$key 
							$thisSubKey=$reg.OpenSubKey($thisKey) 
							If(![String]::IsNullOrEmpty($($thisSubKey.GetValue("DisplayName")))) 
							{
								$obj = New-Object PSObject
								$obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $($thisSubKey.GetValue("DisplayName"))
								$obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value $($thisSubKey.GetValue("DisplayVersion"))
								$InstalledApps += $obj
							}
						}
					}			

					$UninstallKey2="SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall" 
					$regkey2=$reg.OpenSubKey($UninstallKey2)
					If($regkey2 -ne $Null)
					{
						$subkeys2=$regkey2.GetSubKeyNames()

						ForEach($key in $subkeys2) 
						{
							$thisKey=$UninstallKey2+"\\"+$key 
							$thisSubKey=$reg.OpenSubKey($thisKey) 
							if(![String]::IsNullOrEmpty($($thisSubKey.GetValue("DisplayName")))) 
							{
								$obj = New-Object PSObject
								$obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $($thisSubKey.GetValue("DisplayName"))
								$obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value $($thisSubKey.GetValue("DisplayVersion"))
								$InstalledApps += $obj
							}
						}
					}

					$InstalledApps = $InstalledApps | Sort DisplayName

					$tmp1 = SWExclusions
					If($Tmp1 -ne "")
					{
						$Func = ConvertTo-ScriptBlock $tmp1
						$tempapps = Invoke-Command {& $Func}
					}
					Else
					{
						$tempapps = $InstalledApps
					}
					
					$JustApps = $TempApps | Select DisplayName, DisplayVersion | Sort DisplayName -unique

					WriteWordLine 0 1 "Installed applications:"
					Write-Verbose "$(Get-Date): `t`tProcessing installed applications for server $($server.ServerName)"
					Write-Verbose "$(Get-Date): `t`tCreate Word Table for server's installed applications"
					$TableRange = $doc.Application.Selection.Range
					[int]$Columns = 2
					[int]$Rows = $JustApps.count + 1

					$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
					$Table.rows.first.headingformat = $wdHeadingFormatTrue
					$Table.Style = $myHash.Word_TableGrid
					$Table.Borders.InsideLineStyle = $wdLineStyleSingle
					$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
					[int]$xRow = 1
					Write-Verbose "$(Get-Date): `t`tFormat first row with column headings"
					$Table.Rows.First.Shading.BackgroundPatternColor = $wdColorGray15
					$Table.Cell($xRow,1).Range.Font.Bold = $True
					$Table.Cell($xRow,1).Range.Text = "Application name"
					$Table.Cell($xRow,2).Range.Font.Bold = $True
					$Table.Cell($xRow,2).Range.Text = "Application version"
					ForEach($app in $JustApps)
					{
						Write-Verbose "$(Get-Date): `t`t`tProcessing installed application $($app.DisplayName)"
						$xRow++
						$Table.Cell($xRow,1).Range.Text = $app.DisplayName
						$Table.Cell($xRow,2).Range.Text = $app.DisplayVersion
					}
					Write-Verbose "$(Get-Date): `t`tMove table of installed applications to the right"
					$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)
					$Table.AutoFitBehavior($wdAutoFitContent)

					FindWordDocumentEnd
					WriteWordLine 0 0 ""
				}

				#list citrix services
				Write-Verbose "$(Get-Date): `t`tTesting to see if $($server.ServerName) is online and reachable"
				If(Test-Connection -ComputerName $server.servername -quiet -EA 0)
				{
					Write-Verbose "$(Get-Date): `t`tProcessing Citrix services for server $($server.ServerName) by calling Get-Service"
					
					Try
					{
						#Iain Brighton optimization 5-Jun-2014
						#Replaced with a single call to retrieve services via WMI. The repeated
						## "Get-WMIObject Win32_Service -Filter" calls were the major delays in the script.
						## If we need to retrieve the StartUp type might as well just use WMI.
						$Services = Get-WMIObject Win32_Service -ComputerName $server.ServerName -EA 0 | Where {$_.DisplayName -like "*Citrix*"} | Sort DisplayName
					}
					
					Catch
					{
						$Services = $Null
					}
					
					WriteWordLine 0 1 "Citrix Services" -NoNewLine
					If($? -and $Services -ne $Null)
					{
						If($Services -is [array])
						{
							[int]$NumServices = $Services.count
						}
						Else
						{
							[int]$NumServices = 1
						}
						Write-Verbose "$(Get-Date): `t`tCreate Word Table for Citrix services"
						Write-Verbose "$(Get-Date): `t`t $NumServices Services found"
						
						If($MSWord -or $PDF)
						{
							WriteWordLine 0 0 " ($NumServices Services found)"
							## IB - replacement Services table generation utilising AddWordTable function

							## Create an array of hashtables to store our services
							[System.Collections.Hashtable[]] $ServicesWordTable = @();
							## Create an array of hashtables to store references of cells that we wish to highlight after the table has been added
							[System.Collections.Hashtable[]] $HighlightedCells = @();
							## Seed the $Services row index from the second row
							[int] $CurrentServiceIndex = 2;
						}
						
						ForEach($Service in $Services) 
						{
							#Write-Verbose "$(Get-Date): `t`t`t Processing service $($Service.DisplayName)";

							If($MSWord -or $PDF)
							{
								## Add the required key/values to the hashtable
								$WordTableRowHash = @{ DisplayName = $Service.DisplayName; Status = $Service.State; StartMode = $Service.StartMode; }

								## Add the hash to the array
								$ServicesWordTable += $WordTableRowHash;

								## Store "to highlight" cell references
								If($Service.State -like "Stopped" -and $Service.StartMode -like "Auto") 
								{
									$HighlightedCells += @{ Row = $CurrentServiceIndex; Column = 2; }
								}
								$CurrentServiceIndex++;
							}
						}
						
						If($MSWord -or $PDF)
						{
							## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
							$Table = AddWordTable -Hashtable $ServicesWordTable `
							-Columns DisplayName, Status, StartMode `
							-Headers "Display Name", "Status", "Startup Type" `
							-AutoFit $wdAutoFitContent;

							## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
							SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
							## IB - Set the required highlighted cells
							SetWordCellFormat -Coordinates $HighlightedCells -Table $Table -Bold -BackgroundColor $wdColorRed -Solid;

							#indent the entire table 1 tab stop
							$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)

							FindWordDocumentEnd
							$Table = $Null
						}
					}
					ElseIf(!$?)
					{
						Write-Warning "No services were retrieved."
						If($MSWord -or $PDF)
						{
							WriteWordLine 0 0 "Warning: No Services were retrieved" "" $Null 0 $False $True
							WriteWordLine 0 1 "If this is a trusted Forest, you may need to rerun the" "" $Null 0 $False $True
							WriteWordLine 0 1 "script with Admin credentials from the trusted Forest." "" $Null 0 $False $True
						}
					}
					Else
					{
						Write-Warning "Services retrieval was successful but no services were returned."
						WriteWordLine 0 0 "Services retrieval was successful but no services were returned." "" $Null 0 $False $True
					}
					
					#Citrix hotfixes installed
					Write-Verbose "$(Get-Date): `t`tGet list of Citrix hotfixes installed using Get-XAServerHotfix"
					Try
					{
						$hotfixes = (Get-XAServerHotfix -ServerName $server.ServerName -EA 0 | Where {$_.Valid -eq $True}) | Sort HotfixName
					}
					
					Catch
					{
						$hotfixes = $Null
					}
					
					If($? -and $hotfixes -ne $Null)
					{
						[int]$Rows = 1
						$Single_Row = (Get-Member -Type Property -Name Length -InputObject $hotfixes -EA 0) -eq $Null
						If(-not $Single_Row)
						{
							$Rows = $Hotfixes.length
						}
						$Rows++
						
						Write-Verbose "$(Get-Date): `t`tNumber of Citrix hotfixes is $($Rows-1)"
						$HotfixArray = @()
						
						If($MSWord -or $PDF)
						{
							WriteWordLine 0 0 ""
							WriteWordLine 0 1 "Citrix Installed Hotfixes ($($Rows-1)):"
							Write-Verbose "$(Get-Date): `t`tCreate Word Table for Citrix Hotfixes"
							Write-Verbose "$(Get-Date): `t`tAdd Citrix installed hotfix table to doc"
							## Create an array of hashtables to store our hotfixes
							[System.Collections.Hashtable[]] $hotfixesWordTable = @();
							## Seed the row index from the second row
							[int] $CurrentServiceIndex = 2;
						}

						ForEach($hotfix in $hotfixes)
						{
							$HotfixArray += $hotfix.HotfixName
							$InstallDate = $hotfix.InstalledOn.ToString()
							
							If($MSWord -or $PDF)
							{
								## Add the required key/values to the hashtable
								$WordTableRowHash = @{ HotfixName = $hotfix.HotfixName; InstalledBy = $hotfix.InstalledBy; InstallDate = $InstallDate.SubString(0,$InstallDate.IndexOf(" ")); HotfixType = $hotfix.HotfixType}

								## Add the hash to the array
								$HotfixesWordTable += $WordTableRowHash;

								$CurrentServiceIndex++;
							}
						}
						
						If($MSWord -or $PDF)
						{
							## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
							$Table = AddWordTable -Hashtable $HotfixesWordTable `
							-Columns HotfixName, InstalledBy, InstallDate, HotfixType `
							-Headers "Hotfix", "Installed By", "Install Date", "Type" `
							-AutoFit $wdAutoFitContent;

							SetWordCellFormat -Collection $Table -Size 10
							## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
							SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -Size 10 -BackgroundColor $wdColorGray15;

							#indent the entire table 1 tab stop
							$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)

							FindWordDocumentEnd
							$Table = $Null
							WriteWordLine 0 0 ""
						}
					}
					ElseIf(!$?)
					{
						Write-Warning "No Citrix hotfixes were retrieved"
						If($MSWORD -or $PDF)
						{
							WriteWordLine 0 0 "Warning: No Citrix hotfixes were retrieved" "" $Null 0 $False $True
						}
					}
					Else
					{
						Write-Warning "Citrix hotfix retrieval was successful but no Citrix hotfixes were returned."
						If($MSWORD -or $PDF)
						{
							WriteWordLine 0 0 "Citrix hotfix retrieval was successful but no Citrix hotfixes were returned." "" $Null 0 $False $True
						}
					}
				}
				Else
				{
					Write-Verbose "$(Get-Date): `t`t$($server.ServerName) is offline or unreachable.  Citrix Services and Hotfix areas skipped."
					WriteWordLine 0 0 "Server $($server.ServerName) was offline or unreachable at "$(Get-Date -Format u)
					WriteWordLine 0 0 "The Citrix Services and Hotfix areas were skipped."
				}

				WriteWordLine 0 0 "" 
				Write-Verbose "$(Get-Date): `t`t`tFinished Processing server $($server.ServerName)"
				Write-Verbose "$(Get-Date): "
			}
			Else
			{
				WriteWordLine 0 0 $server.ServerName
			}
		}
	}
	ElseIf(!$?)
	{
		Write-Warning "Server information could not be retrieved"
	}
	Else
	{
		Write-Warning "No results returned for Server information"
	}
	$servers = $Null
	Write-Verbose "$(Get-Date): Finished Processing Servers"
	Write-Verbose "$(Get-Date): "
}

If($Section -eq "All" -or $Section -eq "Zones")
{
	Write-Verbose "$(Get-Date): Processing Zones"
	Write-Verbose "$(Get-Date): `tSetting summary variables"
	[int]$TotalZones = 0

	Write-Verbose "$(Get-Date): `tRetrieving Zone Information"
	$Zones = Get-XAZone -EA 0 | Sort ZoneName
	If($? -and $Zones -ne $Null)
	{
		$ZoneSetting1 = $Null
		$ZoneSetting2 = $Null
		If(!$Summary)
		{
			Write-Verbose "$(Get-Date): `tRetrieving Global Zone Settings"
			$ZoneGlobal = Get-XAFarmConfiguration -EA 0
			If($?)
			{
				[bool]$ZoneSetting1 = $ZoneGlobal.ZonesShareLoadInformation
				[bool]$ZoneSetting2 = $ZoneGlobal.ZonesEnumerationFromDataCollectorsOnly
			}
			ElseIf($ZoneGlobal -eq $Null)
			{
				Write-Verbose "$(Get-Date): There are no Global Zone Settings available"
			}
			Else 
			{
				Write-Warning "No results returned for Global Zone settings"
			}
		}
		
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Zones:"
		ForEach($Zone in $Zones)
		{
			$TotalZones++
			Write-Verbose "$(Get-Date): `t`tProcessing zone $($Zone.ZoneName)"
			If(!$Summary)
			{
				WriteWordLine 2 0 $Zone.ZoneName
				WriteWordLine 0 1 "Current Data Collector: " $Zone.DataCollector
				$Servers = Get-XAServer -ZoneName $Zone.ZoneName -EA 0 | Sort ElectionPreference, ServerName
				If($? -and $Servers -ne $Null)
				{		
					WriteWordLine 0 1 "Servers in Zone"
			
					ForEach($Server in $Servers)
					{
						WriteWordLine 0 2 "Server Name and Preference: " $server.ServerName -NoNewLine
						WriteWordLine 0 0  " - " -nonewline
						Switch ($server.ElectionPreference)
						{
							"Unknown"           {WriteWordLine 0 0 "Unknown"}
							"MostPreferred"     {WriteWordLine 0 0 "Most Preferred"}
							"Preferred"         {WriteWordLine 0 0 "Preferred"}
							"DefaultPreference" {WriteWordLine 0 0 "Default Preference"}
							"NotPreferred"      {WriteWordLine 0 0 "Not Preferred"}
							"WorkerMode"        {WriteWordLine 0 0 "Worker Mode"}
							Default {WriteWordLine 0 0 "Zone preference could not be determined: $($server.ElectionPreference)"}
						}
					}
				}
				Else
				{
					WriteWordLine 0 1 "Unable to enumerate servers in the zone"
				}
				$Servers = $Null
			}
			Else
			{
				WriteWordLine 0 0 $Zone.ZoneName
			}
		}
		
		If($ZoneSetting1 -ne $Null)
		{
			Write-Verbose "$(Get-Date): `t`tProcessing global zone data"
			WriteWordLine 0 0 ""
			WriteWordLine 0 1 "Only zone data collectors enumerate Program Neighborhood: " -nonewline
			If($ZoneSetting1 -eq $True)
			{
				WriteWordLine 0 0 "Yes"
			}
			Else
			{
				WriteWordLine 0 0 "No"
			}
			WriteWordLine 0 1 "Share load information across zones: " -nonewline
			If($ZoneSetting2 -eq $True)
			{
				WriteWordLine 0 0 "Yes"
			}
			Else
			{
				WriteWordLine 0 0 "No"
			}
		}
	}
	ElseIf(!$?)
	{
		Write-Warning "Zone information could not be retrieved"
	}
	Else
	{
		Write-Warning "No results returned for Zone information"
	}
	$Servers = $Null
	$Zone = $Null
	Write-Verbose "$(Get-Date): Finished Processing Zones"
	Write-Verbose "$(Get-Date): "
}

#Process the nodes in the Advanced Configuration Console

If($Section -eq "All" -or $Section -eq "LoadEvals")
{
	Write-Verbose "$(Get-Date): Processing Load Evaluators"
	#load evaluators
	Write-Verbose "$(Get-Date): `tSetting summary variables"
	[int]$TotalLoadEvaluators = 0

	Write-Verbose "$(Get-Date): `tRetrieving Load Evaluators"
	$LoadEvaluators = Get-XALoadEvaluator -EA 0 | Sort LoadEvaluatorName

	If($? -and $LoadEvaluators -ne $Null)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Load Evaluators:"
		ForEach($LoadEvaluator in $LoadEvaluators)
		{
			$TotalLoadEvaluators++
			Write-Verbose "$(Get-Date): `t`tProcessing load evaluator $($LoadEvaluator.LoadEvaluatorName)"
			If(!$Summary)
			{
				WriteWordLine 2 0 $LoadEvaluator.LoadEvaluatorName
				WriteWordLine 0 1 "Description: " $LoadEvaluator.Description
				
				If($LoadEvaluator.IsBuiltIn)
				{
					WriteWordLine 0 1 "Built-in Load Evaluator"
				} 
				Else 
				{
					WriteWordLine 0 1 "User created load evaluator"
				}
			
				If($LoadEvaluator.ApplicationUserLoadEnabled)
				{
					WriteWordLine 0 1 "Application User Load Settings"
					WriteWordLine 0 2 "Report full load when the # of users for this application =: " $LoadEvaluator.ApplicationUserLoad
					WriteWordLine 0 2 "Application: " $LoadEvaluator.ApplicationBrowserName
				}
			
				If($LoadEvaluator.ContextSwitchesEnabled)
				{
					WriteWordLine 0 1 "Context Switches Settings"
					WriteWordLine 0 2 "Report full load when the # of context Switches per second is >= than: " $LoadEvaluator.ContextSwitches[1]
					WriteWordLine 0 2 "Report no load when the # of context Switches per second is <= to: " $LoadEvaluator.ContextSwitches[0]
				}
			
				If($LoadEvaluator.CpuUtilizationEnabled)
				{
					WriteWordLine 0 1 "CPU Utilization Settings"
					WriteWordLine 0 2 "Report full load when the processor utilization % is > than: " $LoadEvaluator.CpuUtilization[1]
					WriteWordLine 0 2 "Report no load when the processor utilization % is <= to: " $LoadEvaluator.CpuUtilization[0]
				}
			
				If($LoadEvaluator.DiskDataIOEnabled)
				{
					WriteWordLine 0 1 "Disk Data I/O Settings"
					WriteWordLine 0 2 "Report full load when the total disk I/O in kbps > than: " $LoadEvaluator.DiskDataIO[1]
					WriteWordLine 0 2 "Report no load when the total disk I/O in kbps <= to: " $LoadEvaluator.DiskDataIO[0]
				}
			
				If($LoadEvaluator.DiskOperationsEnabled)
				{
					WriteWordLine 0 1 "Disk Operations Settings"
					WriteWordLine 0 2 "Report full load when the total # of R/W operations per second is > than: " $LoadEvaluator.DiskOperations[1]
					WriteWordLine 0 2 "Report no load when the total # of R/W operations per second is <= to: " $LoadEvaluator.DiskOperations[0]
				}
			
				If($LoadEvaluator.IPRangesEnabled)
				{
					WriteWordLine 0 1 "IP Range Settings"
					If($LoadEvaluator.IPRangesAllowed)
					{
						WriteWordLine 0 2 "Allow " -NoNewLine
					} 
					Else 
					{
						WriteWordLine 0 2 "Deny " -NoNewLine
					}
					WriteWordLine 0 2 "client connections from the listed IP Ranges"
					ForEach($IPRange in $LoadEvaluator.IPRanges)
					{
						WriteWordLine 0 4 "IP Address Ranges: " $IPRange
					}
				}

				If($LoadEvaluator.LoadThrottlingEnabled)
				{
					WriteWordLine 0 1 "Load Throttling Settings"
					WriteWordLine 0 2 "Impact of logons on load: " -nonewline
					Switch ($LoadEvaluator.LoadThrottling)
					{
						"Unknown"    {WriteWordLine 0 0 "Unknown"}
						"Extreme"    {WriteWordLine 0 0 "Extreme"}
						"High"       {WriteWordLine 0 0 "High (Default)"}
						"MediumHigh" {WriteWordLine 0 0 "Medium High"}
						"Medium"     {WriteWordLine 0 0 "Medium"}
						"MediumLow"  {WriteWordLine 0 0 "Medium Low"}
						Default {WriteWordLine 0 0 "Impact of logons on load could not be determined: $($LoadEvaluator.LoadThrottling)"}
					}
				}
				
				If($LoadEvaluator.MemoryUsageEnabled)
				{
					WriteWordLine 0 1 "Memory Usage Settings"
					WriteWordLine 0 2 "Report full load when the memory usage is > than: " $LoadEvaluator.MemoryUsage[1]
					WriteWordLine 0 2 "Report no load when the memory usage is <= to: " $LoadEvaluator.MemoryUsage[0]
				}
			
				If($LoadEvaluator.PageFaultsEnabled)
				{
					WriteWordLine 0 1 "Page Faults Settings"
					WriteWordLine 0 2 "Report full load when the # of page faults per second is > than: " $LoadEvaluator.PageFaults[1]
					WriteWordLine 0 2 "Report no load when the # of page faults per second is <= to: " $LoadEvaluator.PageFaults[0]
				}
			
				If($LoadEvaluator.PageSwapsEnabled)
				{
					WriteWordLine 0 1 "Page Swaps Settings"
					WriteWordLine 0 2 "Report full load when the # of page swaps per second is > than: " $LoadEvaluator.PageSwaps[1]
					WriteWordLine 0 2 "Report no load when the # of page swaps per second is <= to: " $LoadEvaluator.PageSwaps[0]
				}
			
				If($LoadEvaluator.ScheduleEnabled)
				{
					WriteWordLine 0 1 "Scheduling Settings"
					WriteWordLine 0 2 "Sunday Schedule`t: " $LoadEvaluator.SundaySchedule
					WriteWordLine 0 2 "Monday Schedule`t: " $LoadEvaluator.MondaySchedule
					WriteWordLine 0 2 "Tuesday Schedule`t: " $LoadEvaluator.TuesdaySchedule
					WriteWordLine 0 2 "Wednesday Schedule`t: " $LoadEvaluator.WednesdaySchedule
					WriteWordLine 0 2 "Thursday Schedule`t: " $LoadEvaluator.ThursdaySchedule
					WriteWordLine 0 2 "Friday Schedule`t`t: " $LoadEvaluator.FridaySchedule
					WriteWordLine 0 2 "Saturday Schedule`t: " $LoadEvaluator.SaturdaySchedule
				}

				If($LoadEvaluator.ServerUserLoadEnabled)
				{
					WriteWordLine 0 1 "Server User Load Settings"
					WriteWordLine 0 2 "Report full load when the # of server users equals: " $LoadEvaluator.ServerUserLoad
				}
			
				WriteWordLine 0 0 ""
			}
			Else
			{
				WriteWordLine 0 0 $LoadEvaluator.LoadEvaluatorName
			}
		}
	}
	ElseIf(!$?)
	{
		Write-Warning "Load Evaluator information could not be retrieved"
	}
	Else
	{
		Write-Warning "No results returned for Load Evaluator information"
	}
	$LoadEvaluators = $Null
	Write-Verbose "$(Get-Date): Finished Processing Load Evaluators"
	Write-Verbose "$(Get-Date): "
}

If($Section -eq "All" -or $Section -eq "Policies")
{
	Write-Verbose "$(Get-Date): Processing Policies"
	Write-Verbose "$(Get-Date): `tSetting summary variables"
	[int]$TotalPolicies = 0

	Write-Verbose "$(Get-Date): `tRetrieving Policies"
	$Policies = Get-XAPolicy -EA 0 | Sort PolicyName
	If($? -and $Policies -ne $Null)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Policies:"
		ForEach($Policy in $Policies)
		{
			$TotalPolicies++
			Write-Verbose "$(Get-Date): `tProcessing policy $($Policy.PolicyName)"
			If(!$Summary)
			{
				WriteWordLine 2 0 $Policy.PolicyName
				WriteWordLine 0 1 "Description: " $Policy.Description
				WriteWordLine 0 1 "Enabled: " $Policy.Enabled
				WriteWordLine 0 1 "Priority: " $Policy.Priority

				$filter = Get-XAPolicyFilter -PolicyName $Policy.PolicyName -EA 0

				If($? -and $Filter -ne $Null)
				{
					If($Filter)
					{
						Write-Verbose "$(Get-Date): `t`tPolicy Filters"
						WriteWordLine 0 1 "Policy Filters:"
						
						If($Filter.AccessControlEnabled)
						{
							If($Filter.AllowConnectionsThroughAccessGateway)
							{
								WriteWordLine 0 2 "Apply to connections made through Access Gateway"
								If($Filter.AccessSessionConditions)
								{
									WriteWordLine 0 3 "Any connection that meets any of the following filters"
									$TableRange = $doc.Application.Selection.Range
									[int]$Columns = 2
									[int]$Rows = $Filter.AccessSessionConditions.count + 1
									$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
									$Table.rows.first.headingformat = $wdHeadingFormatTrue
									$Table.Style = $myHash.Word_TableGrid
									$Table.Borders.InsideLineStyle = $wdLineStyleSingle
									$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
									[int]$xRow = 1
									Write-Verbose "$(Get-Date): `t`t`t`tFormat first row with column headings"
									$Table.Rows.First.Shading.BackgroundPatternColor = $wdColorGray15
									$Table.Cell($xRow,1).Range.Font.Bold = $True
									$Table.Cell($xRow,1).Range.Text = "Farm Name"
									$Table.Cell($xRow,2).Range.Font.Bold = $True
									$Table.Cell($xRow,2).Range.Text = "Filter"
									ForEach($AccessCondition in $Filter.AccessSessionConditions)
									{
										[string]$Tmp = $AccessCondition
										[string]$AGFarm = $Tmp.substring(0, $Tmp.indexof(":"))
										[string]$AGFilter = $Tmp.substring($Tmp.indexof(":")+1)
										$xRow++
										Write-Verbose "$(Get-Date): `t`t`t`t`tProcessing row for Access Condition $($Tmp)"
										$Table.Cell($xRow,1).Range.Text = $AGFarm
										$Table.Cell($xRow,2).Range.Text = $AGFilter
									}

									$Table.Rows.SetLeftIndent($Indent3TabStops,$wdAdjustProportional)
									$Table.AutoFitBehavior($wdAutoFitContent)

									FindWordDocumentEnd
									$tmp = $Null
									$AGFarm = $Null
									$AGFilter = $Null
								}
								Else
								{
									WriteWordLine 0 3 "Any connection"
								}
								WriteWordLine 0 3 "Apply to all other connections: " -nonewline
								If($Filter.AllowOtherConnections)
								{
									WriteWordLine 0 0 "Yes"
								}
								Else
								{
									WriteWordLine 0 0 "No"
								}
							}
							Else
							{
								WriteWordLine 0 2 "Do not apply to connections made through Access Gateway"
							}
						}
						If($Filter.ClientIPAddressEnabled)
						{
							WriteWordLine 0 2 "Apply to all client IP addresses: " -nonewline
							If($Filter.ApplyToAllClientIPAddresses)
							{
								WriteWordLine 0 0 "Yes"
							}
							Else
							{
								WriteWordLine 0 0 "No"
							}
							If($Filter.AllowedIPAddresses)
							{
								WriteWordLine 0 2 "Allowed IP Addresses:"
								ForEach($Allowed in $Filter.AllowedIPAddresses)
								{
									WriteWordLine 0 3 $Allowed
								}
							}
							If($Filter.DeniedIPAddresses)
							{
								WriteWordLine 0 2 "Denied IP Addresses:"
								ForEach($Denied in $Filter.DeniedIPAddresses)
								{
									WriteWordLine 0 3 $Denied
								}
							}
						}
						If($Filter.ClientNameEnabled)
						{
							WriteWordLine 0 2 "Apply to all client names: " -nonewline
							If($Filter.ApplyToAllClientNames)
							{
								WriteWordLine 0 0 "Yes"
							}
							Else
							{
								WriteWordLine 0 0 "No"
							}
							If($Filter.AllowedClientNames)
							{
								WriteWordLine 0 2 "Allowed Client Names:"
								ForEach($Allowed in $Filter.AllowedClientNames)
								{
									WriteWordLine 0 3 $Allowed
								}
							}
							If($Filter.DeniedClientNames)
							{
								WriteWordLine 0 2 "Denied Client Names:"
								ForEach($Denied in $Filter.DeniedClientNames)
								{
									WriteWordLine 0 3 $Denied
								}
							}
						}
						If($Filter.ServerEnabled)
						{
							If($Filter.AllowedServerNames)
							{
								WriteWordLine 0 2 "Allowed Server Names:"
								ForEach($Allowed in $Filter.AllowedServerNames)
								{
									WriteWordLine 0 3 $Allowed
								}
							}
							If($Filter.DeniedServerNames)
							{
								WriteWordLine 0 2 "Denied Server Names:"
								ForEach($Denied in $Filter.DeniedServerNames)
								{
									WriteWordLine 0 3 $Denied
								}
							}
							If($Filter.AllowedServerFolders)
							{
								WriteWordLine 0 2 "Allowed Server Folders:"
								ForEach($Allowed in $Filter.AllowedServerFolders)
								{
									WriteWordLine 0 3 $Allowed
								}
							}
							If($Filter.DeniedServerFolders)
							{
								WriteWordLine 0 2 "Denied Server Folders:"
								ForEach($Denied in $Filter.DeniedServerFolders)
								{
									WriteWordLine 0 3 $Denied
								}
							}
						}
						If($Filter.AccountEnabled)
						{
							WriteWordLine 0 2 "Apply to all explicit (non-anonymous) users: " -nonewline
							If($Filter.ApplyToAllExplicitAccounts)
							{
								WriteWordLine 0 0 "Yes"
							}
							Else
							{
								WriteWordLine 0 0 "No"
							}
							WriteWordLine 0 2 "Apply to anonymous users: " -nonewline
							If($Filter.ApplyToAnonymousAccounts)
							{
								WriteWordLine 0 0 "Yes"
							}
							Else
							{
								WriteWordLine 0 0 "No"
							}
							If($Filter.AllowedAccounts)
							{
								WriteWordLine 0 2 "Allowed Accounts:"
								ForEach($Allowed in $Filter.AllowedAccounts)
								{
									WriteWordLine 0 3 $Allowed
								}
							}
							If($Filter.DeniedAccounts)
							{
								WriteWordLine 0 2 "Denied Accounts:"
								ForEach($Denied in $Filter.DeniedAccounts)
								{
									WriteWordLine 0 3 $Denied
								}
							}
						}
					}
					Else
					{
						WriteWordLine 0 1 "No filter information"
					}
				}
				ElseIf(!$?)
				{
					WriteWordLine 0 1 "Unable to retrieve Filter settings"
				}
				Else
				{
					WriteWordLine 0 1 "No Filter settings were found"
				}

				$Settings = Get-XAPolicyConfiguration -PolicyName $Policy.PolicyName -EA 0

				If($? -and $Settings -ne $Null)
				{
					Write-Verbose "$(Get-Date): `t`tPolicy Settings"
					WriteWordLine 0 1 "Policy Settings:"
					ForEach($Setting in $Settings)
					{
						Process2008Policies
					}
				}
				ElseIf(!$?)
				{
					WriteWordLine 0 1 "Unable to retrieve settings"
				}
				Else
				{
					WriteWordLine 0 1 "No policy settings were found"
				}
			
				$Settings = $Null
				$Filter = $Null
				Write-Verbose "$(Get-Date): `t`tFinished Processing policy $($Policy.PolicyName)"
				Write-Verbose "$(Get-Date): "
			}
			Else
			{
				WriteWordLine 0 0 $Policy.PolicyName
			}
		}
	}
	ElseIf($Policies -eq $Null)
	{
		Write-Verbose "$(Get-Date): There are no Policies created"
	}
	Else 
	{
		Write-Warning "No results returned for Citrix Policy information"
	}
	$Policies = $Null
	Write-Verbose "$(Get-Date): Finished Processing Policies"
	Write-Verbose "$(Get-Date): "
}

If($Section -eq "All" -or $Section -eq "Printers")
{
	Write-Verbose "$(Get-Date): Processing Print Drivers"
	#printer drivers
	Write-Verbose "$(Get-Date): `tSetting summary variables"
	[int]$TotalPrintDrivers = 0

	Write-Verbose "$(Get-Date): `tRetrieving Print Drivers"
	$PrinterDrivers = Get-XAPrinterDriver -EA 0 | Sort DriverName

	If($? -and $PrinterDrivers -ne $Null)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Print Drivers"
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 3
		[int]$Rows = $PrinterDrivers.count + 1
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$Table.rows.first.headingformat = $wdHeadingFormatTrue
		$Table.Style = $myHash.Word_TableGrid
		$Table.Borders.InsideLineStyle = $wdLineStyleSingle
		$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
		[int]$xRow = 1
		Write-Verbose "$(Get-Date): `t`tFormat first row with column headings"
		$Table.Rows.First.Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Text = "Driver"
		$Table.Cell($xRow,2).Range.Font.Bold = $True
		$Table.Cell($xRow,2).Range.Text = "Platform"
		$Table.Cell($xRow,3).Range.Font.Bold = $True
		$Table.Cell($xRow,3).Range.Text = "64 bit?"
		ForEach($PrinterDriver in $PrinterDrivers)
		{
			$TotalPrintDrivers++
			Write-Verbose "$(Get-Date): `t`tProcessing driver $($PrinterDriver.DriverName)"
			$xRow++
			$Table.Cell($xRow,1).Range.Text = $PrinterDriver.DriverName
			$Table.Cell($xRow,2).Range.Text = $PrinterDriver.OSVersion
			$Table.Cell($xRow,3).Range.Text = $PrinterDriver.Is64Bit
		}
		$Table.AutoFitBehavior($wdAutoFitContent)

		FindWordDocumentEnd
	}
	ElseIf($PrinterDrivers -eq $Null)
	{
		Write-Verbose "$(Get-Date): There are no Printer Drivers created"
	}
	Else 
	{
		Write-Warning "No results returned for Printer driver information"
	}
	$PrintDrivers = $Null
	Write-Verbose "$(Get-Date): Finished Processing Print Drivers"
	Write-Verbose "$(Get-Date): "

	Write-Verbose "$(Get-Date): Processing Printer Driver Mappings"
	#printer driver mappings
	Write-Verbose "$(Get-Date): `tSetting summary variables"
	[int]$TotalPrintDriverMappings = 0

	Write-Verbose "$(Get-Date): `tRetrieving Print Driver Mappings"
	$PrinterDriverMappings = Get-XAPrinterDriverMapping -EA 0 | Sort ClientDriverName

	If($? -and $PrinterDriverMappings -ne $Null)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Print Driver Mappings"
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 4
		[int]$Rows = $PrinterDriverMappings.count + 1
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$Table.rows.first.headingformat = $wdHeadingFormatTrue
		$Table.Style = $myHash.Word_TableGrid
		$Table.Borders.InsideLineStyle = $wdLineStyleSingle
		$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
		[int]$xRow = 1
		Write-Verbose "$(Get-Date): `t`tFormat first row with column headings"
		$Table.Rows.First.Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Text = "Client Driver"
		$Table.Cell($xRow,2).Range.Font.Bold = $True
		$Table.Cell($xRow,2).Range.Text = "Server Driver"
		$Table.Cell($xRow,3).Range.Font.Bold = $True
		$Table.Cell($xRow,3).Range.Text = "Platform"
		$Table.Cell($xRow,4).Range.Font.Bold = $True
		$Table.Cell($xRow,4).Range.Text = "64 bit?"
		ForEach($PrinterDriverMapping in $PrinterDriverMappings)
		{
			$TotalPrintDriverMappings++
			Write-Verbose "$(Get-Date): `t`tProcessing drive $($PrinterDriverMapping.ClientDriverName)"
			$xRow++
			$Table.Cell($xRow,1).Range.Text = $PrinterDriverMapping.ClientDriverName
			$Table.Cell($xRow,2).Range.Text = $PrinterDriverMapping.ServerDriverName
			$Table.Cell($xRow,3).Range.Text = $PrinterDriverMapping.OSVersion
			$Table.Cell($xRow,4).Range.Text = $PrinterDriverMapping.Is64Bit
		}
		$Table.AutoFitBehavior($wdAutoFitContent)

		FindWordDocumentEnd
	}
	ElseIf($PrinterDriverMappings -eq $Null)
	{
		Write-Verbose "$(Get-Date): There are no Printer Driver Mappings created"
	}
	Else 
	{
		Write-Warning "No results returned for Printer driver mapping information"
	}
	$PrintDriverMappings = $Null
	Write-Verbose "$(Get-Date): Finished Processing Printer Driver Mappings"
	Write-Verbose "$(Get-Date): "
}

If(!$Summary -and ($Section -eq "All" -or $Section -eq "ConfigLog"))
{
	Write-Verbose "$(Get-Date): Setting summary variables"
	[int]$TotalConfigLogItems = 0
	If($ConfigLog)
	{
		Write-Verbose "$(Get-Date): Processing the Configuration Logging Report"
		#Configuration Logging report
		#only process if $ConfigLog = $True and XA5ConfigLog.udl file exists
		#build connection string for Microsoft SQL Server
		#User ID is account that has access permission for the configuration logging database
		#Initial Catalog is the name of the Configuration Logging SQL Database
		If(Test-Path “$($pwd.path)\XA5ConfigLog.udl”)
		{
			Write-Verbose "$(Get-Date): `tRetrieving Configuration Logging Data"
			$ConnectionString = Get-Content “$($pwd.path)\XA5ConfigLog.udl” | select-object -last 1
			$ConfigLogReport = get-CtxConfigurationLogReport -connectionstring $ConnectionString -TimePeriodFrom $StartDate -TimePeriodTo $EndDate -EA 0

			If($? -and $ConfigLogReport -ne $Null)
			{
				Write-Verbose "$(Get-Date): `tProcessing $($ConfigLogReport.Count) configuration logging items"
				$selection.InsertNewPage()
				WriteWordLine 1 0 "Configuration Log Report:"
				WriteWordLine 0 0 " For date range $($StartDate) through $($EndDate)"
				$TableRange   = $doc.Application.Selection.Range
				[int]$Columns = 6
				If($ConfigLogReport -is [array])
				{
					[int]$Rows = $ConfigLogReport.Count +1
				}
				Else
				{
					[int]$Rows = 2
				}
				Write-Verbose "$(Get-Date): `t`tAdd Configuration Logging table to doc"
				$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
				$Table.rows.first.headingformat = $wdHeadingFormatTrue
				$Table.Style = $myHash.Word_TableGrid
				$Table.Borders.InsideLineStyle = $wdLineStyleNone
				$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
				$Table.Rows.First.Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(1,1).Range.Font.Bold = $True
				$Table.Cell(1,1).Range.Text = "Date"
				$Table.Cell(1,2).Range.Font.Bold = $True
				$Table.Cell(1,2).Range.Text = "Account"
				$Table.Cell(1,3).Range.Font.Bold = $True
				$Table.Cell(1,3).Range.Text = "Change description"
				$Table.Cell(1,4).Range.Font.Bold = $True
				$Table.Cell(1,4).Range.Text = "Type of change"
				$Table.Cell(1,5).Range.Font.Bold = $True
				$Table.Cell(1,5).Range.Text = "Type of item"
				$Table.Cell(1,6).Range.Font.Bold = $True
				$Table.Cell(1,6).Range.Text = "Name of item"
				$xRow = 1
				ForEach($Item in $ConfigLogReport)
				{
					$xRow++
					Write-Verbose "$(Get-Date): `t`t`tAdding row for $($Item.Description)"
					If($xRow % 2 -eq 0)
					{
						$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray05
						$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray05
						$Table.Cell($xRow,3).Shading.BackgroundPatternColor = $wdColorGray05
						$Table.Cell($xRow,4).Shading.BackgroundPatternColor = $wdColorGray05
						$Table.Cell($xRow,5).Shading.BackgroundPatternColor = $wdColorGray05
						$Table.Cell($xRow,6).Shading.BackgroundPatternColor = $wdColorGray05
					}
					$Table.Cell($xRow,1).Range.Font.size = 9
					$Table.Cell($xRow,1).Range.Text = $Item.Date
					$Table.Cell($xRow,2).Range.Font.size = 9
					$Table.Cell($xRow,2).Range.Text = $Item.Account
					$Table.Cell($xRow,3).Range.Font.size = 9
					$Table.Cell($xRow,3).Range.Text = $Item.Description
					$Table.Cell($xRow,4).Range.Font.size = 9
					$Table.Cell($xRow,4).Range.Text = $Item.TaskType
					$Table.Cell($xRow,5).Range.Font.size = 9
					$Table.Cell($xRow,5).Range.Text = $Item.ItemType
					$Table.Cell($xRow,6).Range.Font.size = 9
					$Table.Cell($xRow,6).Range.Text = $Item.ItemName
				}
				$Table.AutoFitBehavior($wdAutoFitContent)

				FindWordDocumentEnd
			} 
			Else 
			{
				WriteWordLine 0 0 "Configuration log report could not be retrieved"
			}
			$ConfigLogReport = $Null
		}
		Else 
		{
			$selection.InsertNewPage()
			WriteWordLine 1 0 "Configuration Logging is enabled but the XA5ConfigLog.udl file was not found"
		}
		Write-Verbose "$(Get-Date): Finished Processing the Configuration Logging Report"
	}
}
Write-Verbose "$(Get-Date): "

If(!$Summary -and ($Section -eq "All"))
{
	Write-Verbose "$(Get-Date): Create Appendix A Session Sharing Items"
	#	The Session Sharing Key is generated by the XML Broker in XenApp 5.  
	#	Web Interface or StoreFront send the following information to the XML Broker:"
	#	Audio Quality (Policy Setting)"
	#	Client Printer Port Mapping (Policy Setting)"
	#	Client Printer Spooling (Policy Setting)"
	#	Color Depth (Application Setting)"
	#	COM Port Mapping (Policy Setting)"
	#	Domain Name (Logon)"
	#	Encryption Level (Application Setting and Policy Setting.  Policy wins.)"
	#	Special Folder Redirection (Policy Setting)"
	#	User Name (Logon)"
	#	Virtual COM Port Emulation (Policy Setting)"
	#
	#	This table consists of the above application settings plus
	#	the application settings from CTX159159
	#	Color depth
	#	Screen Size
	#	Access Control Filters (for SmartAccess)
	#	Encryption
	#
	#	In addition, a XenApp server can have Session Sharing disable in a registry key
	#	To disable session sharing, the following registry key must be present.
	#	This information has been added to the Server Appendix B section
	#
	#	Add the following value to disable this feature (this value does not exist by default):
	#	HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Citrix\Wfshell\TWI\:
	#	Type: REG_DWORD
	#	Value: SeamlessFlags = 1

	Write-Verbose "$(Get-Date): Create Appendix A Session Sharing Items"
	Write-Verbose "$(Get-Date): `tAdd Session Sharing Items table to doc"
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Appendix A - Session Sharing Items from CTX159159"
		## Create an array of hashtables to store our services
		[System.Collections.Hashtable[]] $ItemsWordTable = @();
		## Seed the row index from the second row
		[int] $CurrentServiceIndex = 2;
	}
	
	ForEach($Item in $SessionSharingItems)
	{
		If($MSWord -or $PDF)
		{
			## Add the required key/values to the hashtable
			$WordTableRowHash = @{ ApplicationName = $Item.ApplicationName;
			MaximumColorQuality = $Item.MaximumColorQuality;
			SessionWindowSize = $Item.SessionWindowSize; 
			AccessControlFilters = $Item.AccessControlFilters;
			Encryption = $Item.Encryption}

			## Add the hash to the array
			$ItemsWordTable += $WordTableRowHash;

			$CurrentServiceIndex++;
		}
	}

	If($MSWord -or $PDF)
	{
		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		$Table = AddWordTable -Hashtable $ItemsWordTable `
		-Columns ApplicationName, MaximumColorQuality, SessionWindowSize, AccessControlFilters, Encryption `
		-Headers "Application Name", "Maximum color quality", "Session window size", "Access Control Filters", "Encryption" `
		-AutoFit $wdAutoFitContent;

		## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$TableRange = $Null
		$Table = $Null
	}
	
	Write-Verbose "$(Get-Date): Finished Create Appendix A - Session Sharing Items"
	Write-Verbose "$(Get-Date): "

	Write-Verbose "$(Get-Date): Create Appendix B Server Major Items"
	Write-Verbose "$(Get-Date): `tAdd Major Server Items table to doc"
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Appendix B - Server Major Items"
		## Create an array of hashtables to store our services
		[System.Collections.Hashtable[]] $ItemsWordTable = @();
		## Seed the row index from the second row
		[int] $CurrentServiceIndex = 2;
	}

	ForEach($Item in $ServerItems)
	{
		$Tmp = $Null
		If([String]::IsNullOrEmpty($Item.LicenseServer))
		{
			$Tmp = "Set by policy"
		}
		Else
		{
			$Tmp = $Item.LicenseServer
		}
		If($MSWord -or $PDF)
		{
			## Add the required key/values to the hashtable
			$WordTableRowHash = @{ ServerName = $Item.ServerName;
			ZoneName = $Item.ZoneName;
			OSVersion = $Item.OSVersion;
			CitrixVersion = $Item.CitrixVersion;
			ProductEdition = $Item.ProductEdition;
			LicenseServer = $Tmp
			SessionSharing = $Item.SessionSharing}
			## Add the hash to the array
			$ItemsWordTable += $WordTableRowHash;

			$CurrentServiceIndex++;
		}
	}

	If($MSWord -or $PDF)
	{
		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		$Table = AddWordTable -Hashtable $ItemsWordTable `
		-Columns ServerName, ZoneName, OSVersion, CitrixVersion, ProductEdition, LicenseServer, SessionSharing `
		-Headers "Server Name", "Zone Name", "OS Version", "Citrix Version", "Product Edition", "License Server", "Session Sharing" `
		-AutoFit $wdAutoFitContent;

		## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	
	Write-Verbose "$(Get-Date): Finished Create Appendix B - Server Major Items"
	Write-Verbose "$(Get-Date): "
}

If($Section -eq "All")
{
	#summary page
	Write-Verbose "$(Get-Date): Create Summary Page"
	$selection.InsertNewPage()
	WriteWordLine 1 0 "Summary Page"
	If(!$Summary)
	{
		Write-Verbose "$(Get-Date): Add administrator summary info"
		WriteWordLine 0 0 "Administrators"
		WriteWordLine 0 1 "Total Full Administrators`t: " $TotalFullAdmins
		WriteWordLine 0 1 "Total View Administrators`t: " $TotalViewAdmins
		WriteWordLine 0 1 "Total Custom Administrators`t: " $TotalCustomAdmins
		WriteWordLine 0 2 "Total Administrators`t: " ($TotalFullAdmins + $TotalViewAdmins + $TotalCustomAdmins)
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): Add application summary info"
		WriteWordLine 0 0 "Applications"
		WriteWordLine 0 1 "Total Published Applications`t: " $TotalPublishedApps
		WriteWordLine 0 1 "Total Published Content`t`t: " $TotalPublishedContent
		WriteWordLine 0 1 "Total Published Desktops`t: " $TotalPublishedDesktops
		WriteWordLine 0 1 "Total Streamed Applications`t: " $TotalStreamedApps
		WriteWordLine 0 2 "Total Applications`t: " ($TotalPublishedApps + $TotalPublishedContent + $TotalPublishedDesktops + $TotalStreamedApps)
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): Add server summary info"
		WriteWordLine 0 0 "Servers"
		WriteWordLine 0 2 "Total Servers`t`t: " $TotalServers
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): Add zone summary info"
		WriteWordLine 0 0 "Zones"
		WriteWordLine 0 2 "Total Zones`t`t: " $TotalZones
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): Add load evaluator summary info"
		WriteWordLine 0 0 "Load Evaluators"
		WriteWordLine 0 2 "Total Load Evaluators`t: " $TotalLoadEvaluators
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): Add policy summary info"
		WriteWordLine 0 0 "Policies"
		WriteWordLine 0 2 "Total Policies`t`t: " $TotalPolicies
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): Add print driver summary info"
		WriteWordLine 0 0 "Print Drivers"
		WriteWordLine 0 2 "Total Print Drivers`t: " $TotalPrintDrivers
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): Add print driver mapping summary info"
		WriteWordLine 0 0 "Print Driver Mappingss"
		WriteWordLine 0 2 "Total Prt Drvr Mappings: " $TotalPrintDriverMappings
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): Add configuration logging summary info"
		WriteWordLine 0 0 "Configuration Logging"
		WriteWordLine 0 2 "Total Config Log Items`t: " $TotalConfigLogItems 
		WriteWordLine 0 0 ""
	}
	Else
	{
		Write-Verbose "$(Get-Date): Add administrator summary info"
		WriteWordLine 0 0 "Administrators"
		WriteWordLine 0 1 "Total Administrators`t: " $TotalAdmins
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): Add application summary info"
		WriteWordLine 0 0 "Applications"
		WriteWordLine 0 1 "Total Applications`t: " $TotalApps
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): Add server summary info"
		WriteWordLine 0 0 "Servers"
		WriteWordLine 0 1 "Total Servers`t`t: " $TotalServers
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): Add zone summary info"
		WriteWordLine 0 0 "Zones"
		WriteWordLine 0 1 "Total Zones`t`t: " $TotalZones
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): Add load evaluator summary info"
		WriteWordLine 0 0 "Load Evaluators"
		WriteWordLine 0 1 "Total Load Evaluators`t: " $TotalLoadEvaluators
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): Add policy summary info"
		WriteWordLine 0 0 "Policies"
		WriteWordLine 0 1 "Total Policies`t`t: " $TotalPolicies
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): Add print driver summary info"
		WriteWordLine 0 0 "Print Drivers"
		WriteWordLine 0 1 "Total Print Drivers`t: " $TotalPrintDrivers
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): Add print driver mapping summary info"
		WriteWordLine 0 0 "Print Driver Mappingss"
		WriteWordLine 0 1 "Total Prt Drvr Mappings: " $TotalPrintDriverMappings
		WriteWordLine 0 0 ""
	}

	Write-Verbose "$(Get-Date): `tFinished Create Summary Page"
	Write-Verbose "$(Get-Date): "
}

Write-Verbose "$(Get-Date): Finishing up document"
#end of document processing

$AbstractTitle = "Citrix XenApp 5 Inventory"
$SubjectTitle = "XenApp 5 Farm Inventory"
UpdateDocumentProperties $AbstractTitle $SubjectTitle

If($MSWORD -or $PDF)
{
    SaveandCloseDocumentandShutdownWord
}

Write-Verbose "$(Get-Date): Script has completed"
Write-Verbose "$(Get-Date): "

If($PDF)
{
	If(Test-Path "$($Script:FileName2)")
	{
		Write-Verbose "$(Get-Date): $($Script:FileName2) is ready for use"
	}
	Else
	{
		Write-Warning "$(Get-Date): Unable to save the output file, $($Script:FileName2)"
		Write-Error "Unable to save the output file, $($Script:FileName2)"
	}
}
Else
{
	If(Test-Path "$($Script:FileName1)")
	{
		Write-Verbose "$(Get-Date): $($Script:FileName1) is ready for use"
	}
	Else
	{
		Write-Warning "$(Get-Date): Unable to save the output file, $($Script:FileName1)"
		Write-Error "Unable to save the output file, $($Script:FileName1)"
	}
}

Write-Verbose "$(Get-Date): "

#http://poshtips.com/measuring-elapsed-time-in-powershell/
Write-Verbose "$(Get-Date): Script started: $($Script:StartTime)"
Write-Verbose "$(Get-Date): Script ended: $(Get-Date)"
$runtime = $(Get-Date) - $Script:StartTime
$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds", `
	$runtime.Days, `
	$runtime.Hours, `
	$runtime.Minutes, `
	$runtime.Seconds,
	$runtime.Milliseconds)
Write-Verbose "$(Get-Date): Elapsed time: $($Str)"
$runtime = $Null
$Str = $Null
$ErrorActionPreference = $SaveEAPreference
