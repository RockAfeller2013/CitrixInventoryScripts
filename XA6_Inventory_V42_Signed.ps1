#Requires -Version 3.0
#This File is in Unicode format.  Do not edit in an ASCII editor.

<#
.SYNOPSIS
	Creates a complete inventory of a Citrix XenApp 6 farm using Microsoft Word 2010 or 2013.
.DESCRIPTION
	Creates a complete inventory of a Citrix XenApp 6 farm using Microsoft Word and PowerShell.
	Creates a Word document or PDF named after the XenApp 6 farm.
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
	Only Word 2010, 2013 and 2016 are supported.
	(default cover pages in Word en-US)
	
	Valid input is:
		Alphabet (Word 2010. Works)
		Annual (Word 2010. Doesn't work well for this report)
		Austere (Word 2010. Works)
		Austin (Word 2010/2013/2016. Doesn't work in 2013 or 2016, mostly works in 2010 but 
						Subtitle/Subject & Author fields need to be moved 
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
		Ion (Dark) (Word 2013/2016. Top date doesn't fit, box needs to be manually resized or font 
						changed to 8 point)
		Ion (Light) (Word 2013/2016. Top date doesn't fit, box needs to be manually resized or font 
						changed to 8 point)
		Mod (Word 2010. Works)
		Motion (Word 2010/2013/2016. Works if top date is manually changed to 36 point)
		Newsprint (Word 2010. Works but date is not populated)
		Perspective (Word 2010. Works)
		Pinstripes (Word 2010. Works)
		Puzzle (Word 2010. Top date doesn't fit, box needs to be manually resized or font 
					changed to 14 point)
		Retrospect (Word 2013/2016. Works)
		Semaphore (Word 2013/2016. Works)
		Sideline (Word 2010/2013/2016. Doesn't work in 2013 or 2016, works in 2010)
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
.PARAMETER PDF
	SaveAs PDF file instead of DOCX file.
	This parameter is disabled by default.
	For Word 2007, the Microsoft add-in for saving as a PDF muct be installed.
	For Word 2007, please see http://www.microsoft.com/en-us/download/details.aspx?id=9943
	The PDF file is roughly 5X to 10X larger than the DOCX file.
.PARAMETER Text
	Creates a formatted text file with a .txt extension.
	This parameter is reserved for a future update.
	This parameter is disabled by default.
.PARAMETER MSWord
	SaveAs DOCX file
	This parameter is set True if no other output format is selected.
.PARAMETER HTML
	Creates an HTML file with an .html extension.
	This parameter is disabled by default.
	This parameter is reserved for a future update and no output is created at this time.
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
	This parameter cannot be used with either the Hardware, Software, StartDate or EndDate parameters.
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
		LBPolicies (Load Balancing Policies)
		LoadEvals (Load Evaluators)
		Policies
		Servers
		WGs (Worker Groups)
		Zones
		All
	This parameter defaults to All sections.
.EXAMPLE
	PS C:\PSScript > .\XA6_Inventory_V42.ps1
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XA6_Inventory_V42.ps1 -PDF 
	
	Will use all Default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XA6_Inventory_V42.ps1 -TEXT

	This parameter is reserved for a future update and no output is created at this time.
	
	Will use all default values and save the document as a formatted text file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XA6_Inventory_V42.ps1 -HTML

	This parameter is reserved for a future update and no output is created at this time.
	
	Will use all default values and save the document as an HTML file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XA65_Inventory_V42.ps1 -Summary
	
	Creates a Summary report with no detail.
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XA65_Inventory_V42.ps1 -PDF -Summary 
	
	Creates a Summary report with no detail.
	Will use all Default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XA6_Inventory_V42.ps1 -Hardware 
	
	Will use all Default values and add additional information for each server about its hardware.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XA6_Inventory_V42.ps1 -StartDate "01/01/2014" -EndDate "01/02/2014" 
	
	Will use all Default values and add additional information for each server about its installed applications.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	Will return all Configuration Logging entries from "01/01/2014 00:00:00" through "01/02/2014 "00:00:00".
.EXAMPLE
	PS C:\PSScript > .\XA6_Inventory_V42.ps1 -StartDate "01/01/2014" -EndDate "01/01/2014" 
	
	Will use all Default values and add additional information for each server about its installed applications.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	Will return all Configuration Logging entries from "01/01/2014 00:00:00" through "01/01/2014 "00:00:00".  In other words, nothing is returned.
.EXAMPLE
	PS C:\PSScript > .\XA6_Inventory_V42.ps1 -StartDate "01/01/2014 21:00:00" -EndDate "01/01/2014 22:00:00" 
	
	Will use all Default values and add additional information for each server about its installed applications.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	Will return all Configuration Logging entries from 9PM to 10PM on 01/01/2014.
.EXAMPLE
	PS C:\PSScript .\XA6_Inventory_V42.ps1 -CompanyName "Carl Webster Consulting" -CoverPage "Mod" -UserName "Carl Webster"

	Will use:
		Carl Webster Consulting for the Company Name.
		Mod for the Cover Page format.
		Carl Webster for the User Name.
.EXAMPLE
	PS C:\PSScript .\XA6_Inventory_V42.ps1 -CN "Carl Webster Consulting" -CP "Mod" -UN "Carl Webster"

	Will use:
		Carl Webster Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Carl Webster for the User Name (alias UN).
.EXAMPLE
	PS C:\PSScript > .\XA6_Inventory_V42.ps1 -Section Policies
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	Processes only the Policies section of the report.
.EXAMPLE
	PS C:\PSScript > .\XA6_Inventory_V42.ps1 -AddDateTime
	
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
	Output filename will be XA6FarmName_2014-06-01_1800.docx
.EXAMPLE
	PS C:\PSScript > .\XA6_Inventory_V42.ps1 -PDF -AddDateTime
	
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
	Output filename will be XA6FarmName_2014-06-01_1800.pdf
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  This script creates a Word or PDF document.
.NOTES
	NAME: XA6_Inventory_V42.ps1
	VERSION: 4.24
	AUTHOR: Carl Webster (with a lot of help from Michael B. Smith, Jeff Wouters and Iain Brighton)
	LASTEDIT: October 5, 2015
#>


#thanks to @jeffwouters and Michael B. Smith for helping me with these parameters
[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "Word") ]

Param(
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[Switch]$MSWord=$False,

	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[Switch]$PDF=$False,

	[parameter(ParameterSetName="Text",Mandatory=$False)] 
	[Switch]$Text=$False,

	[parameter(ParameterSetName="HTML",Mandatory=$False)] 
	[Switch]$HTML=$False,

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
#This script is designed to be run on a XenApp 6 server

#Modifications by Carl Webster, CTP and independent consultant
#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com
#originally released to the Citrix community on September 30, 2011
#Version 4.24 5-Oct-2015
#	Added support for Word 2016
#Version 4.23 17-Aug-2015
#	Updated for CTX129229 that was updated August 2015
#Version 4.22 25-Jul-2015
#	Updated for CTX129229 dated 1-Apr-2015
#	Add checking for KB3014783 for Server 2008 R2 w/o SP1
#	Added most current hardware inventory code
#	Cleaned up extraneous console output
#Version 4.21 18-Dec-2014
#	Updated for CTX129229 dated 18-Dec-2014
#	Fix wrong variable name for saving as PDF for Word 2013
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
#	Move Citrix and Microsoft hotfix tables to new table functions
#	Move hardware info to new table functions
#	Add more write statements and error handling to the Configuration Logging report section
#	Add parameters for MSWord, Text and HTML for future updates
#	Add Section parameter
#	Valid Section options are:
#		Admins (Administrators)
#		Apps (Applications)
#		ConfigLog (Configuration Logging)
#		LBPolicies (Load Balancing Policies)
#		LoadEvals (Load Evaluators)
#		Policies
#		Servers
#		WGs (Worker Groups)
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
If($Text -eq $Null)
{
	$Text = $False
}
If($HTML -eq $Null)
{
	$HTML = $False
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
If(!(Test-Path Variable:Text))
{
	$Text = $False
}
If(!(Test-Path Variable:HTML))
{
	$HTML = $False
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
	If($MSWord -eq $Null)
	{
		Write-Verbose "$(Get-Date): MSWord is Null"
	}
	ElseIf($PDF -eq $Null)
	{
		Write-Verbose "$(Get-Date): PDF is Null"
	}
	ElseIf($Text -eq $Null)
	{
		Write-Verbose "$(Get-Date): Text is Null"
	}
	ElseIf($HTML -eq $Null)
	{
		Write-Verbose "$(Get-Date): HTML is Null"
	}
	Else
	{
		Write-Verbose "$(Get-Date): MSWord is $($MSWord)"
		Write-Verbose "$(Get-Date): PDF is $($PDF)"
		Write-Verbose "$(Get-Date): Text is $($Text)"
		Write-Verbose "$(Get-Date): HTML is $($HTML)"
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
	"LBPolicies" {$ValidSection = $True}
	"LoadEvals" {$ValidSection = $True}
	"Policies" {$ValidSection = $True}
	"Servers" {$ValidSection = $True}
	"WGs" {$ValidSection = $True}
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
	`t`tLBPolicies
	`t`tLoadEvals
	`t`tPolicies
	`t`tServers
	`t`tWGs
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

	[string]$RunningOS = (Get-WmiObject -class Win32_OperatingSystem -EA 0).Caption
}

#region code for -hardware switch
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
		WriteWordLine 3 0 "Computer Information: $($RemoteComputerName)"
		WriteWordLine 4 0 "General Computer"
	}
	ElseIf($Text)
	{
		Line 0 "Computer Information: $($RemoteComputerName)"
		Line 1 "General Computer"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 3 0 "Computer Information: $($RemoteComputerName)"
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
		$ComputerItems = $Results | Select Manufacturer, Model, Domain, `
		@{N="TotalPhysicalRam"; E={[math]::round(($_.TotalPhysicalMemory / 1GB),0)}}, `
		NumberOfProcessors, NumberOfLogicalProcessors
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
		ElseIf($Text)
		{
			Line 2 "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
			Line 2 ""
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results Returned for Computer information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for Computer information" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "No results Returned for Computer information"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for Computer information" "" $Null 0 $False $True
		}
	}
	
	#Get Disk info
	Write-Verbose "$(Get-Date): `t`t`tDrive information"

	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Drive(s)"
	}
	ElseIf($Text)
	{
		Line 1 "Drive(s)"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 2 0 "Drive(s)"
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
		ElseIf($Text)
		{
			Line 2 "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results Returned for Drive information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for Drive information" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "No results Returned for Drive information"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for Drive information" "" $Null 0 $False $True
		}
	}
	

	#Get CPU's and stepping
	Write-Verbose "$(Get-Date): `t`t`tProcessor information"

	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Processor(s)"
	}
	ElseIf($Text)
	{
		Line 1 "Processor(s)"
	}
	ElseIf($HTML)
	{
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
		ElseIf($Text)
		{
			Line 2 "Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "Get-WmiObject win32_Processor failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results Returned for Processor information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for Processor information" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "No results Returned for Processor information"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for Processor information" "" $Null 0 $False $True
		}
	}

	#Get Nics
	Write-Verbose "$(Get-Date): `t`t`tNIC information"

	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Network Interface(s)"
	}
	ElseIf($Text)
	{
		Line 1 "Network Interface(s)"
	}
	ElseIf($HTML)
	{
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
					ElseIf($Text)
					{
						Line 2 "Error retrieving NIC information"
						Line 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
						Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
						Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
						Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
					}
					ElseIf($HTML)
					{
						WriteHTMLLine 0 2 "Error retrieving NIC information" "" $Null 0 $False $True
						WriteHTMLLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" "" $Null 0 $False $True
						WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
						WriteHTMLLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
						WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
					}
				}
				Else
				{
					Write-Verbose "$(Get-Date): No results Returned for NIC information"
					If($MSWORD -or $PDF)
					{
						WriteWordLine 0 2 "No results Returned for NIC information" "" $Null 0 $False $True
					}
					ElseIf($Text)
					{
						Line 2 "No results Returned for NIC information"
					}
					ElseIf($HTML)
					{
						WriteHTMLLine 0 2 "No results Returned for NIC information" "" $Null 0 $False $True
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
		ElseIf($Text)
		{
			Line 2 "Error retrieving NIC configuration information"
			Line 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "Error retrieving NIC configuration information" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results Returned for NIC configuration information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for NIC configuration information" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "No results Returned for NIC configuration information"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for NIC configuration information" "" $Null 0 $False $True
		}
	}
	
	If($MSWORD -or $PDF)
	{
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 0 0 ""
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
		$ItemInformation += @{ Data = "Physical Processors (sockets)"; Value = $Item.NumberOfProcessors; }
		$ItemInformation += @{ Data = "Logical Processors (cores w/HT)"; Value = $Item.NumberOfLogicalProcessors; }
		$Table = AddWordTable -Hashtable $ItemInformation `
		-Columns Data,Value `
		-List `
		-AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 2 ""
	}
	ElseIf($Text)
	{
		Line 2 "Manufacturer`t: " $Item.manufacturer
		Line 2 "Model`t`t: " $Item.model
		Line 2 "Domain`t`t: " $Item.domain
		Line 2 "Total Ram`t: $($Item.totalphysicalram) GB"
		Line 2 "Physical Processors (sockets): " $Item.NumberOfProcessors
		Line 2 "Logical Processors (cores w/HT): " $Item.NumberOfLogicalProcessors
		Line 2 ""
	}
	ElseIf($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Manufacturer",($htmlsilver -bor $htmlbold),$Item.manufacturer,$htmlwhite)
		$rowdata += @(,('Model',($htmlsilver -bor $htmlbold),$Item.model,$htmlwhite))
		$rowdata += @(,('Domain',($htmlsilver -bor $htmlbold),$Item.domain,$htmlwhite))
		$rowdata += @(,('Total Ram',($htmlsilver -bor $htmlbold),"$($Item.totalphysicalram) GB",$htmlwhite))
		$rowdata += @(,('Physical Processors (sockets)',($htmlsilver -bor $htmlbold),$Item.NumberOfProcessors,$htmlwhite))
		$rowdata += @(,('Logical Processors (cores w/HT)',($htmlsilver -bor $htmlbold),$Item.NumberOfLogicalProcessors,$htmlwhite))

		$msg = "General Computer"
		$columnWidths = @("150px","200px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}

Function OutputDriveItem
{
	Param([object]$Drive)
	
	$xDriveType = ""
	Switch ($drive.drivetype)
	{
		0	{$xDriveType = "Unknown"}
		1	{$xDriveType = "No Root Directory"}
		2	{$xDriveType = "Removable Disk"}
		3	{$xDriveType = "Local Disk"}
		4	{$xDriveType = "Network Drive"}
		5	{$xDriveType = "Compact Disc"}
		6	{$xDriveType = "RAM Disk"}
		Default {$xDriveType = "Unknown"}
	}
	
	$xVolumeDirty = ""
	If(![String]::IsNullOrEmpty($drive.volumedirty))
	{
		If($drive.volumedirty)
		{
			$xVolumeDirty = "Yes"
		}
		Else
		{
			$xVolumeDirty = "No"
		}
	}

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
			$DriveInformation += @{ Data = "Volume is Dirty"; Value = $xVolumeDirty; }
		}
		If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
		{
			$DriveInformation += @{ Data = "Volume Serial Number"; Value = $Drive.volumeserialnumber; }
		}
		$DriveInformation += @{ Data = "Drive Type"; Value = $xDriveType; }
		$Table = AddWordTable -Hashtable $DriveInformation `
		-Columns Data,Value `
		-List `
		-AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells `
		-Bold `
		-BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 2 ""
	}
	ElseIf($Text)
	{
		Line 2 "Caption`t`t: " $drive.caption
		Line 2 "Size`t`t: $($drive.drivesize) GB"
		If(![String]::IsNullOrEmpty($drive.filesystem))
		{
			Line 2 "File System`t: " $drive.filesystem
		}
		Line 2 "Free Space`t: $($drive.drivefreespace) GB"
		If(![String]::IsNullOrEmpty($drive.volumename))
		{
			Line 2 "Volume Name`t: " $drive.volumename
		}
		If(![String]::IsNullOrEmpty($drive.volumedirty))
		{
			Line 2 "Volume is Dirty`t: " $xVolumeDirty
		}
		If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
		{
			Line 2 "Volume Serial #`t: " $drive.volumeserialnumber
		}
		Line 2 "Drive Type`t: " $xDriveType
		Line 2 ""
	}
	ElseIf($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Caption",($htmlsilver -bor $htmlbold),$Drive.caption,$htmlwhite)
		$rowdata += @(,('Size',($htmlsilver -bor $htmlbold),"$($drive.drivesize) GB",$htmlwhite))

		If(![String]::IsNullOrEmpty($drive.filesystem))
		{
			$rowdata += @(,('File System',($htmlsilver -bor $htmlbold),$Drive.filesystem,$htmlwhite))
		}
		$rowdata += @(,('Free Space',($htmlsilver -bor $htmlbold),"$($drive.drivefreespace) GB",$htmlwhite))
		If(![String]::IsNullOrEmpty($drive.volumename))
		{
			$rowdata += @(,('Volume Name',($htmlsilver -bor $htmlbold),$Drive.volumename,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($drive.volumedirty))
		{
			$rowdata += @(,('Volume is Dirty',($htmlsilver -bor $htmlbold),$xVolumeDirty,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
		{
			$rowdata += @(,('Volume Serial Number',($htmlsilver -bor $htmlbold),$Drive.volumeserialnumber,$htmlwhite))
		}
		$rowdata += @(,('Drive Type',($htmlsilver -bor $htmlbold),$xDriveType,$htmlwhite))

		$msg = ""
		$columnWidths = @("150px","200px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}

Function OutputProcessorItem
{
	Param([object]$Processor)
	
	$xAvailability = ""
	Switch ($processor.availability)
	{
		1	{$xAvailability = "Other"}
		2	{$xAvailability = "Unknown"}
		3	{$xAvailability = "Running or Full Power"}
		4	{$xAvailability = "Warning"}
		5	{$xAvailability = "In Test"}
		6	{$xAvailability = "Not Applicable"}
		7	{$xAvailability = "Power Off"}
		8	{$xAvailability = "Off Line"}
		9	{$xAvailability = "Off Duty"}
		10	{$xAvailability = "Degraded"}
		11	{$xAvailability = "Not Installed"}
		12	{$xAvailability = "Install Error"}
		13	{$xAvailability = "Power Save - Unknown"}
		14	{$xAvailability = "Power Save - Low Power Mode"}
		15	{$xAvailability = "Power Save - Standby"}
		16	{$xAvailability = "Power Cycle"}
		17	{$xAvailability = "Power Save - Warning"}
		Default	{$xAvailability = "Unknown"}
	}

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
			$ProcessorInformation += @{ Data = "Number of Logical Processors (cores w/HT)"; Value = $Processor.numberoflogicalprocessors; }
		}
		$ProcessorInformation += @{ Data = "Availability"; Value = $xAvailability; }
		$Table = AddWordTable -Hashtable $ProcessorInformation `
		-Columns Data,Value `
		-List `
		-AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 2 "Name`t`t`t: " $processor.name
		Line 2 "Description`t`t: " $processor.description
		Line 2 "Max Clock Speed`t`t: $($processor.maxclockspeed) MHz"
		If($processor.l2cachesize -gt 0)
		{
			Line 2 "L2 Cache Size`t`t: $($processor.l2cachesize) KB"
		}
		If($processor.l3cachesize -gt 0)
		{
			Line 2 "L3 Cache Size`t`t: $($processor.l3cachesize) KB"
		}
		If($processor.numberofcores -gt 0)
		{
			Line 2 "# of Cores`t`t: " $processor.numberofcores
		}
		If($processor.numberoflogicalprocessors -gt 0)
		{
			Line 2 "# of Logical Procs (cores w/HT)`t: " $processor.numberoflogicalprocessors
		}
		Line 2 "Availability`t`t: " $xAvailability
		Line 2 ""
	}
	ElseIf($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Name",($htmlsilver -bor $htmlbold),$Processor.name,$htmlwhite)
		$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$Processor.description,$htmlwhite))

		$rowdata += @(,('Max Clock Speed',($htmlsilver -bor $htmlbold),"$($processor.maxclockspeed) MHz",$htmlwhite))
		If($processor.l2cachesize -gt 0)
		{
			$rowdata += @(,('L2 Cache Size',($htmlsilver -bor $htmlbold),"$($processor.l2cachesize) KB",$htmlwhite))
		}
		If($processor.l3cachesize -gt 0)
		{
			$rowdata += @(,('L3 Cache Size',($htmlsilver -bor $htmlbold),"$($processor.l3cachesize) KB",$htmlwhite))
		}
		If($processor.numberofcores -gt 0)
		{
			$rowdata += @(,('Number of Cores',($htmlsilver -bor $htmlbold),$Processor.numberofcores,$htmlwhite))
		}
		If($processor.numberoflogicalprocessors -gt 0)
		{
			$rowdata += @(,('Number of Logical Processors (cores w/HT)',($htmlsilver -bor $htmlbold),$Processor.numberoflogicalprocessors,$htmlwhite))
		}
		$rowdata += @(,('Availability',($htmlsilver -bor $htmlbold),$xAvailability,$htmlwhite))

		$msg = "Processor(s)"
		$columnWidths = @("150px","200px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}

Function OutputNicItem
{
	Param([object]$Nic, [object]$ThisNic)
	
	$xAvailability = ""
	Switch ($processor.availability)
	{
		1	{$xAvailability = "Other"}
		2	{$xAvailability = "Unknown"}
		3	{$xAvailability = "Running or Full Power"}
		4	{$xAvailability = "Warning"}
		5	{$xAvailability = "In Test"}
		6	{$xAvailability = "Not Applicable"}
		7	{$xAvailability = "Power Off"}
		8	{$xAvailability = "Off Line"}
		9	{$xAvailability = "Off Duty"}
		10	{$xAvailability = "Degraded"}
		11	{$xAvailability = "Not Installed"}
		12	{$xAvailability = "Install Error"}
		13	{$xAvailability = "Power Save - Unknown"}
		14	{$xAvailability = "Power Save - Low Power Mode"}
		15	{$xAvailability = "Power Save - Standby"}
		16	{$xAvailability = "Power Cycle"}
		17	{$xAvailability = "Power Save - Warning"}
		Default	{$xAvailability = "Unknown"}
	}

	$xIPAddress = @()
	ForEach($IPAddress in $Nic.ipaddress)
	{
		$xIPAddress += "$($IPAddress)"
	}

	$xIPSubnet = @()
	ForEach($IPSubnet in $Nic.ipsubnet)
	{
		$xIPSubnet += "$($IPSubnet)"
	}

	If($nic.dnsdomainsuffixsearchorder -ne $Null -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
	{
		$nicdnsdomainsuffixsearchorder = $nic.dnsdomainsuffixsearchorder
		$xnicdnsdomainsuffixsearchorder = @()
		ForEach($DNSDomain in $nicdnsdomainsuffixsearchorder)
		{
			$xnicdnsdomainsuffixsearchorder += "$($DNSDomain)"
		}
	}
	
	If($nic.dnsserversearchorder -ne $Null -and $nic.dnsserversearchorder.length -gt 0)
	{
		$nicdnsserversearchorder = $nic.dnsserversearchorder
		$xnicdnsserversearchorder = @()
		ForEach($DNSServer in $nicdnsserversearchorder)
		{
			$xnicdnsserversearchorder += "$($DNSServer)"
		}
	}

	$xdnsenabledforwinsresolution = ""
	If($nic.dnsenabledforwinsresolution)
	{
		$xdnsenabledforwinsresolution = "Yes"
	}
	Else
	{
		$xdnsenabledforwinsresolution = "No"
	}
	
	$xTcpipNetbiosOptions = ""
	Switch ($nic.TcpipNetbiosOptions)
	{
		0	{$xTcpipNetbiosOptions = "Use NetBIOS setting from DHCP Server"}
		1	{$xTcpipNetbiosOptions = "Enable NetBIOS"}
		2	{$xTcpipNetbiosOptions = "Disable NetBIOS"}
		Default	{$xTcpipNetbiosOptions = "Unknown"}
	}
	
	$xwinsenablelmhostslookup = ""
	If($nic.winsenablelmhostslookup)
	{
		$xwinsenablelmhostslookup = "Yes"
	}
	Else
	{
		$xwinsenablelmhostslookup = "No"
	}

	If($MSWORD -or $PDF)
	{
		[System.Collections.Hashtable[]] $NicInformation = @()
		$NicInformation += @{ Data = "Name"; Value = $ThisNic.Name; }
		If($ThisNic.Name -ne $nic.description)
		{
			$NicInformation += @{ Data = "Description"; Value = $Nic.description; }
		}
		$NicInformation += @{ Data = "Connection ID"; Value = $ThisNic.NetConnectionID; }
		$NicInformation += @{ Data = "Manufacturer"; Value = $Nic.manufacturer; }
		$NicInformation += @{ Data = "Availability"; Value = $xAvailability; }
		$NicInformation += @{ Data = "Physical Address"; Value = $Nic.macaddress; }
		If($xIPAddress.Count -gt 1)
		{
			$NicInformation += @{ Data = "IP Address"; Value = $xIPAddress[0]; }
			$NicInformation += @{ Data = "Default Gateway"; Value = $Nic.Defaultipgateway; }
			$NicInformation += @{ Data = "Subnet Mask"; Value = $xIPSubnet[0]; }
			$cnt = -1
			ForEach($tmp in $xIPAddress)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$NicInformation += @{ Data = "IP Address"; Value = $tmp; }
					$NicInformation += @{ Data = "Subnet Mask"; Value = $xIPSubnet[$cnt]; }
				}
			}
		}
		Else
		{
			$NicInformation += @{ Data = "IP Address"; Value = $xIPAddress; }
			$NicInformation += @{ Data = "Default Gateway"; Value = $Nic.Defaultipgateway; }
			$NicInformation += @{ Data = "Subnet Mask"; Value = $xIPSubnet; }
		}
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
			$NicInformation += @{ Data = "DNS Search Suffixes"; Value = $xnicdnsdomainsuffixsearchorder[0]; }
			$cnt = -1
			ForEach($tmp in $xnicdnsdomainsuffixsearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$NicInformation += @{ Data = ""; Value = $tmp; }
				}
			}
		}
		$NicInformation += @{ Data = "DNS WINS Enabled"; Value = $xdnsenabledforwinsresolution; }
		If($nic.dnsserversearchorder -ne $Null -and $nic.dnsserversearchorder.length -gt 0)
		{
			$NicInformation += @{ Data = "DNS Servers"; Value = $xnicdnsserversearchorder[0]; }
			$cnt = -1
			ForEach($tmp in $xnicdnsserversearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$NicInformation += @{ Data = ""; Value = $tmp; }
				}
			}
		}
		$NicInformation += @{ Data = "NetBIOS Setting"; Value = $xTcpipNetbiosOptions; }
		$NicInformation += @{ Data = "WINS: Enabled LMHosts"; Value = $xwinsenablelmhostslookup; }
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

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($Text)
	{
		Line 2 "Name`t`t`t: " $ThisNic.Name
		If($ThisNic.Name -ne $nic.description)
		{
			Line 2 "Description`t`t: " $nic.description
		}
		Line 2 "Connection ID`t`t: " $ThisNic.NetConnectionID
		Line 2 "Manufacturer`t`t: " $ThisNic.manufacturer
		Line 2 "Availability`t`t: " $xAvailability
		Line 2 "Physical Address`t: " $nic.macaddress
		Line 2 "IP Address`t`t: " $xIPAddress[0]
		$cnt = -1
		ForEach($tmp in $xIPAddress)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 5 "" $tmp
			}
		}
		Line 2 "Default Gateway`t`t: " $Nic.Defaultipgateway
		Line 2 "Subnet Mask`t`t: " $xIPSubnet[0]
		$cnt = -1
		ForEach($tmp in $xIPSubnet)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 5 "" $tmp
			}
		}
		If($nic.dhcpenabled)
		{
			$DHCPLeaseObtainedDate = $nic.ConvertToDateTime($nic.dhcpleaseobtained)
			$DHCPLeaseExpiresDate = $nic.ConvertToDateTime($nic.dhcpleaseexpires)
			Line 2 "DHCP Enabled`t`t: " $nic.dhcpenabled
			Line 2 "DHCP Lease Obtained`t: " $dhcpleaseobtaineddate
			Line 2 "DHCP Lease Expires`t: " $dhcpleaseexpiresdate
			Line 2 "DHCP Server`t`t:" $nic.dhcpserver
		}
		If(![String]::IsNullOrEmpty($nic.dnsdomain))
		{
			Line 2 "DNS Domain`t`t: " $nic.dnsdomain
		}
		If($nic.dnsdomainsuffixsearchorder -ne $Null -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
		{
			[int]$x = 1
			Line 2 "DNS Search Suffixes`t:" $xnicdnsdomainsuffixsearchorder[0]
			$cnt = -1
			ForEach($tmp in $xnicdnsdomainsuffixsearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{ Data = ""; Value = $tmp; }
				}
			}
		}
		Line 2 "DNS WINS Enabled`t: " $xdnsenabledforwinsresolution
		If($nic.dnsserversearchorder -ne $Null -and $nic.dnsserversearchorder.length -gt 0)
		{
			[int]$x = 1
			Line 2 "DNS Servers`t`t:" $xnicdnsserversearchorder[0]
			$cnt = -1
			ForEach($tmp in $xnicdnsserversearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{ Data = ""; Value = $tmp; }
				}
			}
		}
		Line 2 "NetBIOS Setting`t`t: " $xTcpipNetbiosOptions
		Line 2 "WINS:"
		Line 3 "Enabled LMHosts`t: " $xwinsenablelmhostslookup
		If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
		{
			Line 3 "Host Lookup File`t: " $nic.winshostlookupfile
		}
		If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
		{
			Line 3 "Primary Server`t`t: " $nic.winsprimaryserver
		}
		If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
		{
			Line 3 "Secondary Server`t: " $nic.winssecondaryserver
		}
		If(![String]::IsNullOrEmpty($nic.winsscopeid))
		{
			Line 3 "Scope ID`t`t: " $nic.winsscopeid
		}
	}
	ElseIf($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Name",($htmlsilver -bor $htmlbold),$ThisNic.Name,$htmlwhite)
		If($ThisNic.Name -ne $nic.description)
		{
			$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$Nic.description,$htmlwhite))
		}
		$rowdata += @(,('Connection ID',($htmlsilver -bor $htmlbold),$ThisNic.NetConnectionID,$htmlwhite))
		$rowdata += @(,('Manufacturer',($htmlsilver -bor $htmlbold),$Nic.manufacturer,$htmlwhite))
		$rowdata += @(,('Availability',($htmlsilver -bor $htmlbold),$xAvailability,$htmlwhite))
		$rowdata += @(,('Physical Address',($htmlsilver -bor $htmlbold),$Nic.macaddress,$htmlwhite))
		$rowdata += @(,('IP Address',($htmlsilver -bor $htmlbold),$xIPAddress[0],$htmlwhite))
		$cnt = -1
		ForEach($tmp in $xIPAddress)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$rowdata += @(,('IP Address',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
			}
		}
		$rowdata += @(,('Default Gateway',($htmlsilver -bor $htmlbold),$Nic.Defaultipgateway,$htmlwhite))
		$rowdata += @(,('Subnet Mask',($htmlsilver -bor $htmlbold),$xIPSubnet[0],$htmlwhite))
		$cnt = -1
		ForEach($tmp in $xIPSubnet)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$rowdata += @(,('Subnet Mask',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
			}
		}
		If($nic.dhcpenabled)
		{
			$DHCPLeaseObtainedDate = $nic.ConvertToDateTime($nic.dhcpleaseobtained)
			$DHCPLeaseExpiresDate = $nic.ConvertToDateTime($nic.dhcpleaseexpires)
			$rowdata += @(,('DHCP Enabled',($htmlsilver -bor $htmlbold),$Nic.dhcpenabled,$htmlwhite))
			$rowdata += @(,('DHCP Lease Obtained',($htmlsilver -bor $htmlbold),$dhcpleaseobtaineddate,$htmlwhite))
			$rowdata += @(,('DHCP Lease Expires',($htmlsilver -bor $htmlbold),$dhcpleaseexpiresdate,$htmlwhite))
			$rowdata += @(,('DHCP Server',($htmlsilver -bor $htmlbold),$Nic.dhcpserver,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.dnsdomain))
		{
			$rowdata += @(,('DNS Domain',($htmlsilver -bor $htmlbold),$Nic.dnsdomain,$htmlwhite))
		}
		If($nic.dnsdomainsuffixsearchorder -ne $Null -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
		{
			$rowdata += @(,('DNS Search Suffixes',($htmlsilver -bor $htmlbold),$xnicdnsdomainsuffixsearchorder[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xnicdnsdomainsuffixsearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}
		}
		$rowdata += @(,('DNS WINS Enabled',($htmlsilver -bor $htmlbold),$xdnsenabledforwinsresolution,$htmlwhite))
		If($nic.dnsserversearchorder -ne $Null -and $nic.dnsserversearchorder.length -gt 0)
		{
			$rowdata += @(,('DNS Servers',($htmlsilver -bor $htmlbold),$xnicdnsserversearchorder[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xnicdnsserversearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}
		}
		$rowdata += @(,('NetBIOS Setting',($htmlsilver -bor $htmlbold),$xTcpipNetbiosOptions,$htmlwhite))
		$rowdata += @(,('WINS: Enabled LMHosts',($htmlsilver -bor $htmlbold),$xwinsenablelmhostslookup,$htmlwhite))
		If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
		{
			$rowdata += @(,('Host Lookup File',($htmlsilver -bor $htmlbold),$Nic.winshostlookupfile,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
		{
			$rowdata += @(,('Primary Server',($htmlsilver -bor $htmlbold),$Nic.winsprimaryserver,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
		{
			$rowdata += @(,('Secondary Server',($htmlsilver -bor $htmlbold),$Nic.winssecondaryserver,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.winsscopeid))
		{
			$rowdata += @(,('Scope ID',($htmlsilver -bor $htmlbold),$Nic.winsscopeid,$htmlwhite))
		}

		$msg = "Network Interface(s)"
		$columnWidths = @("150px","200px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}
#endregion

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
	
Function ConvertNumberToTime
{
	Param([int]$val = 0)
	
	#this is stored as a number between 0 (00:00 AM) and 1439 (23:59 PM)
	#180 = 3AM
	#900 = 3PM
	#1027 = 5:07 PM
	#[int] (1027/60) = 17 or 5PM
	#1027 % 60 leaves 7 or 7 minutes
	
	#thanks to MBS for the next line
	[int]$hour = [System.Math]::Floor(([int] $val) / ([int] 60))
	[int]$minute = $val % 60
	[string]$Strminute = $minute.ToString()
	[string]$tempminute = ""
	If($Strminute.length -lt 2)
	{
		$tempMinute = "0" + $Strminute
	}
	Else
	{
		$tempminute = $strminute
	}
	[string]$AMorPM = "AM"
	If($Hour -ge 0 -and $Hour -le 11)
	{
		$AMorPM = "AM"
	}
	Else
	{
		$AMorPM = "PM"
		If($Hour -ge 12)
		{
			$Hour = $Hour - 12
		}
	}
	Return "$($hour):$($tempminute) $($AMorPM)"
}

Function ConvertIntegerToDate
{
	#thanks to MBS for helping me on this Function
	Param([int]$DateAsInteger = 0)
	
	#this is stored as an integer but is actually a bitmask
	#01/01/2013 = 131924225 = 11111011101 00000001 00000001
	#01/17/2013 = 131924241 = 11111011101 00000001 00010001
	#
	# last 8 bits are the day
	# previous 8 bits are the month
	# the rest (up to 16) are the year
	
	[int]$year  = [Math]::Floor($DateAsInteger / 65536)
	[int]$month = [Math]::Floor($DateAsInteger / 256) % 256
	[int]$day   = $DateAsInteger % 256

	Return "$Month/$Day/$Year"
}
	
Function Check-LoadedModule
#Function created by Jeff Wouters
#@JeffWouters on Twitter
#modified by Michael B. Smith to handle when the module doesn't exist on server
#modified by @andyjmorgan
#bug fixed by @schose
#bug fixed by Peter Bosen
#This Function handles all three scenarios:
#
# 1. Module is already imported into current session
# 2. Module is not already imported into current session, it does exists on the server and is imported
# 3. Module does not exist on the server

{
	Param([parameter(Mandatory = $True)][alias("Module")][string]$ModuleName)
	#$LoadedModules = Get-Module | Select Name
	#following line changed at the recommendation of @andyjmorgan
	$LoadedModules = Get-Module |% { $_.Name.ToString() }
	#bug reported on 21-JAN-2013 by @schose 
	#the following line did not work if the citrix.grouppolicy.commands.psm1 module
	#was manually loaded from a non Default folder
	#$ModuleFound = (!$LoadedModules -like "*$ModuleName*")
	
	[bool]$ModuleFound = ($LoadedModules -like "*$ModuleName*")
	If(!$ModuleFound) 
	{
		$module = Import-Module -Name $ModuleName -PassThru -EA 0 4>$Null
		If($module -and $?)
		{
			# module imported properly
			Return $True
		}
		Else
		{
			# module import failed
			Return $False
		}
	}
	Else
	{
		#module already imported into current session
		Return $True
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
	$loadedSnapins += Get-pssnapin | % {$_.name}
	$registeredSnapins += Get-pssnapin -Registered | % {$_.name}

	ForEach ($Snapin in $Snapins)
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
				Add-PSSnapin -Name $snapin -EA 0 *>$Null
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
	$prop = $properties | ForEach { 
		$propname = $_.GetType().InvokeMember("Name","GetProperty",$Null,$_,$Null)
		If($propname -eq $Name) 
		{
			Return $_
		}
	} #ForEach

	#set the value
	$Prop.GetType().InvokeMember("Value","SetProperty",$Null,$prop,$Value)
}

Function AbortScript
{
	$Script:Word.quit()
	Write-Verbose "$(Get-Date): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
	If(Test-Path variable:global:word)
	{
		Remove-Variable -Name word -Scope Global 4>$Null
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

Function ProcessCitrixPolicies
{
	Param([string]$xDriveName)

	If($xDriveName -eq "")
	{
		If($Summary)
		{
			WriteWordLine 0 0 ""
			WriteWordLine 0 0 "IMA Policies"
			$Policies = Get-CtxGroupPolicy -EA 0 | Sort Type,PolicyName
		}
		Else
		{
			$Policies = Get-CtxGroupPolicy -EA 0 | Sort Type,Priority
		}
	}
	Else
	{
		If($Summary)
		{
			WriteWordLine 0 0 ""
			WriteWordLine 0 0 "Active Directory Policies"
			$Policies = Get-CtxGroupPolicy -DriveName $xDriveName -EA 0 | Sort Type,PolicyName
		}
		Else
		{
			$Policies = Get-CtxGroupPolicy -DriveName $xDriveName -EA 0 | Sort Type,Priority
		}
	}

	If($? -and $Policies -ne $Null)
	{
		ForEach($Policy in $Policies)
		{
			Write-Verbose "$(Get-Date): `tStarted $($Policy.PolicyName)`t$($Policy.Type)"
			If(!$Summary)
			{
				If($xDriveName -eq "")
				{
					$Global:TotalIMAPolicies++
					WriteWordLine 2 0 $Policy.PolicyName
					WriteWordLine 0 1 "IMA Farm based policy"
				}
				Else
				{
					$Global:TotalADPolicies++
					#requested by Pavel Stadler to show which AD Policy a Citrix policy is contained in
					WriteWordLine 2 0 "$($Policy.PolicyName) in $($CtxGPO)"
					WriteWordLine 0 1 "Active Directory based policy"
				}

				WriteWordLine 0 1 "Type`t`t: " $Policy.Type
					
				If($Policy.Type -eq "Computer")
				{
					$Global:TotalComputerPolicies++
				}
				Else
				{
					$Global:TotalUserPolicies++
				}
				
				If(![String]::IsNullOrEmpty($Policy.Description))
				{
					WriteWordLine 0 1 "Description`t: " $Policy.Description
				}
				WriteWordLine 0 1 "Enabled`t`t: " $Policy.Enabled
				WriteWordLine 0 1 "Priority`t`t: " $Policy.Priority

				If($xDriveName -eq "")
				{
					$filters = Get-CtxGroupPolicyFilter -PolicyName $Policy.PolicyName -EA 0
				}
				Else
				{
					$filters = Get-CtxGroupPolicyFilter -PolicyName $Policy.PolicyName -DriveName $xDriveName -EA 0
				}

				If($? -and $Filters -ne $Null)
				{
					If(![String]::IsNullOrEmpty($filters))
					{
						WriteWordLine 0 1 "Filter(s)`t`t:"
						ForEach($Filter in $Filters)
						{
							WriteWordLine 0 2 "Filter name`t: " $filter.FilterName
							WriteWordLine 0 2 "Filter type`t: " -nonewline
							Switch($filter.FilterType)
							{
								"User"           {WriteWordLine 0 0 "User or Group"}
								"WorkerGroup"    {WriteWordLine 0 0 "Worker Group"}
								"OU"             {WriteWordLine 0 0 "Organization Unit"}
								"ClientName"     {WriteWordLine 0 0 "Client Name"}
								"ClientIP"       {WriteWordLine 0 0 "Client IP Address"}
								"BranchRepeater" {WriteWordLine 0 0 "Branch Repeater"}
								"AccessControl"  {WriteWordLine 0 0 "Access Control"}
								Default {WriteWordLine 0 3 "Policy Filter Type could not be determined: $($filter.FilterType)"}
							}
							WriteWordLine 0 2 "Filter enabled`t: " $filter.Enabled
							WriteWordLine 0 2 "Filter mode`t: " $filter.Mode
							If(![String]::IsNullOrEmpty($filter.FilterValue))
							{
								WriteWordLine 0 2 "Filter value`t: " $filter.FilterValue
							}
							WriteWordLine 0 2 ""
						}
					}
					Else
					{
						WriteWordLine 0 1 "Filter(s)`t`t: None"
					}
				}
				Else
				{
					WriteWordLine 0 1 "Unable to retrieve Filter settings"
				}

				If($xDriveName -eq "")
				{
					$Settings = Get-CtxGroupPolicyConfiguration -PolicyName $Policy.PolicyName -EA 0
				}
				Else
				{
					$Settings = Get-CtxGroupPolicyConfiguration -PolicyName $Policy.PolicyName -DriveName $xDriveName -EA 0
				}

				If($? -and $Settings -ne $Null)
				{
					ForEach($Setting in $Settings)
					{
						If($Setting.Type -eq "Computer")
						{
							Write-Verbose "$(Get-Date): `t`tComputer settings"
							Write-Verbose "$(Get-Date): `t`t`tICA"
							WriteWordLine 0 1 "Computer settings:"
							If($Setting.IcaListenerTimeout.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\ICA listener connection timeout (milliseconds): " $Setting.IcaListenerTimeout.Value
							}
							If($Setting.IcaListenerPortNumber.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\ICA listener port number: " $Setting.IcaListenerPortNumber.Value
							}
							If($Setting.AutoClientReconnect.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Auto Client Reconnect\Auto client reconnect: " $Setting.AutoClientReconnect.State
							}
							If($Setting.AutoClientReconnectAuthenticationRequired.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Auto Client Reconnect\Auto client reconnect authorization: " 
								Switch ($Setting.AutoClientReconnectAuthenticationRequired.Value)
								{
									"DoNotRequireAuthentication" {WriteWordLine 0 3 "Do not require authentication"}
									"RequireAuthentication"      {WriteWordLine 0 3 "Require authentication"}
									Default {WriteWordLine 0 0 "Auto client reconnect authorization could not be determined: $($Setting.AutoClientReconnectAuthenticationRequired.Value)"}
								}
							}
							If($Setting.AutoClientReconnectLogging.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Auto Client Reconnect\Auto client reconnect logging: "
								Switch ($Setting.AutoClientReconnectLogging.Value)
								{
									"DoNotLogAutoReconnectEvents" {WriteWordLine 0 3 "Do Not Log auto-reconnect events"}
									"LogAutoReconnectEvents"      {WriteWordLine 0 3 "Log auto-reconnect events"}
									Default {WriteWordLine 0 3 "Auto client reconnect logging could not be determined: $($Setting.AutoClientReconnectLogging.Value)"}
								}
							}
							If($Setting.IcaRoundTripCalculation.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\End User Monitoring\ICA round trip calculation: " $Setting.IcaRoundTripCalculation.State
							}
							If($Setting.IcaRoundTripCalculationInterval.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\End User Monitoring\ICA round trip calculation interval (seconds): " $Setting.IcaRoundTripCalculationInterval.Value
							}
							If($Setting.IcaRoundTripCalculationWhenIdle.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\End User Monitoring\ICA round trip calculations for idle connections: " 
								WriteWordLine 0 3 $Setting.IcaRoundTripCalculationWhenIdle.State
							}
							If($Setting.DisplayMemoryLimit.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Graphics\Display memory limit (KB): " $Setting.DisplayMemoryLimit.Value
							}
							If($Setting.DisplayDegradePreference.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Graphics\Display mode degrade preference: "
								
								Switch ($Setting.DisplayDegradePreference.Value)
								{
									"ColorDepth" {WriteWordLine 0 3 "Degrade color depth first"}
									"Resolution" {WriteWordLine 0 3 "Degrade resolution first"}
									Default {WriteWordLine 0 3 "Display mode degrade preference could not be determined: $($Setting.DisplayDegradePreference.Value)"}
								}
							}
							If($Setting.ImageCaching.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Graphics\Image caching: " $Setting.ImageCaching.State
							}
							If($Setting.MaximumColorDepth.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Graphics\Maximum allowed color depth: "
								Switch ($Setting.MaximumColorDepth.Value)
								{
									"BitsPerPixel8"  {WriteWordLine 0 3 "8 Bits Per Pixel"}
									"BitsPerPixel15" {WriteWordLine 0 3 "15 Bits Per Pixel"}
									"BitsPerPixel16" {WriteWordLine 0 3 "16 Bits Per Pixel"}
									"BitsPerPixel24" {WriteWordLine 0 3 "24 Bits Per Pixel"}
									"BitsPerPixel32" {WriteWordLine 0 3 "32 Bits Per Pixel"}
									Default {WriteWordLine 0 3 "Maximum allowed color depth could not be determined: $($Setting.MaximumColorDepth.Value)"}
								}
							}
							If($Setting.DisplayDegradeUserNotification.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Graphics\Notify user when display mode is degraded: " $Setting.DisplayDegradeUserNotification.State
							}
							If($Setting.QueueingAndTossing.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Graphics\Queueing and tossing: " $Setting.QueueingAndTossing.State
							}
							If($Setting.IcaKeepAliveTimeout.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Keep Alive\ICA keep alive timeout (seconds): " $Setting.IcaKeepAliveTimeout.Value
							}
							If($Setting.IcaKeepAlives.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Keep Alive\ICA keep alives: "
								Switch ($Setting.IcaKeepAlives.Value)
								{
									"DoNotSendKeepAlives" {WriteWordLine 0 3 "Do not send ICA keep alive messages"}
									"SendKeepAlives"      {WriteWordLine 0 3 "Send ICA keep alive messages"}
									Default {WriteWordLine 0 3 "ICA keep alives could not be determined: $($Setting.IcaKeepAlives.Value)"}
								}
							}
							If($Setting.MultimediaAcceleration.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Multimedia\HDX MediaStream Multimedia Acceleration: " $Setting.MultimediaAcceleration.State
							}
							If($Setting.MultimediaAccelerationDefaultBufferSize.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Multimedia\HDX MediaStream Multimedia Acceleration "
								WriteWordLine 0 3 "Default buffer Size (seconds): " $Setting.MultimediaAccelerationDefaultBufferSize.Value
							}
							If($Setting.MultimediaAccelerationUseDefaultBufferSize.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Multimedia\HDX MediaStream Multimedia Acceleration "
								WriteWordLine 0 3 "Default buffer Size Use: " $Setting.MultimediaAccelerationUseDefaultBufferSize.State
							}
							If($Setting.MultimediaConferencing.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Multimedia\Multimedia conferencing: " $Setting.MultimediaConferencing.State
							}
							If($Setting.PromptForPassword.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Security\Prompt for password: " $Setting.PromptForPassword.State
							}
							If($Setting.IdleTimerInterval.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Server Limits\Server idle timer interval (milliseconds): " $Setting.IdleTimerInterval.Value
							}
							If($Setting.SessionReliabilityConnections.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Session Reliability\Session reliability connections: " $Setting.SessionReliabilityConnections.State
							}
							If($Setting.SessionReliabilityPort.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Session Reliability\Session reliability port number: " $Setting.SessionReliabilityPort.Value
							}
							If($Setting.SessionReliabilityTimeout.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Session Reliability\Session reliability timeout (seconds): " $Setting.SessionReliabilityTimeout.Value
							}
							If($Setting.Shadowing.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Shadowing\Shadowing: " $Setting.Shadowing.State
							}
							Write-Verbose "$(Get-Date): `t`t`tLicensing"
							If($Setting.LicenseServerHostName.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Licensing\License server host name: " $Setting.LicenseServerHostName.Value
							}
							If($Setting.LicenseServerPort.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Licensing\License server port: " $Setting.LicenseServerPort.Value
							}
							Write-Verbose "$(Get-Date): `t`t`tServer Settings"
							If($Setting.ConnectionAccessControl.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Connection access control: "
								Switch ($Setting.ConnectionAccessControl.Value)
								{
									"AllowAny"                     {WriteWordLine 0 3 "Any connections"}
									"AllowTicketedConnectionsOnly" {WriteWordLine 0 3 "Citrix Access Gateway, Citrix Receiver, and Web Interface connections only"}
									"AllowAccessGatewayOnly"       {WriteWordLine 0 3 "Citrix Access Gateway connections only"}
									Default {WriteWordLine 0 3 "Connection access control could not be determined: $($Setting.ConnectionAccessControl.Value)"}
								}
							}
							If($Setting.DnsAddressResolution.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\DNS address resolution: " $Setting.DnsAddressResolution.State
							}
							If($Setting.FullIconCaching.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Full icon caching: " $Setting.FullIconCaching.State
							}
							If($Setting.ProductEdition.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\XenApp product edition: " $Setting.ProductEdition.Value
							}
							If($Setting.UserSessionLimit.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Connection Limits\Limit user sessions: " $Setting.UserSessionLimit.Value
							}
							If($Setting.UserSessionLimitAffectsAdministrators.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Connection Limits\Limits on administrator sessions: " $Setting.UserSessionLimitAffectsAdministrators.State
							}
							If($Setting.UserSessionLimitLogging.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Connection Limits\Logging of logon limit events: " $Setting.UserSessionLimitLogging.State
							}
							If($Setting.HealthMonitoring.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Health Monitoring and Recovery\Health monitoring: " $Setting.HealthMonitoring.State
							}
							If($Setting.HealthMonitoringTests.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Health Monitoring and Recovery\Health monitoring tests: " 
								[xml]$XML = $Setting.HealthMonitoringTests.Value
								ForEach($Test in $xml.hmrtests.tests.test)
								{
									Write-Verbose "$(Get-Date): `t`t`t`tCreate Table for HMR Test $($test.name)"
									$TableRange = $doc.Application.Selection.Range
									[int]$Columns = 2
									[int]$Rows = $test.attributes.count
									$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
									$Table.Style = $myHash.Word_TableGrid
									$Table.Borders.InsideLineStyle = $wdLineStyleNone
									$Table.Borders.OutsideLineStyle = $wdLineStyleNone
									[int]$xRow = 1
									$Table.Cell($xRow,1).Range.Text = "Name"
									$Table.Cell($xRow,2).Range.Text = $test.name
									$xRow++
									$Table.Cell($xRow,1).Range.Text = "File Location"
									$Table.Cell($xRow,2).Range.Text = $test.file
									If($test.HasAttribute("arguments"))
									{
										$xRow++
										$Table.Cell($xRow,1).Range.Text = "Arguments"
										$Table.Cell($xRow,2).Range.Text = $test.arguments
									}
									If(![String]::IsNullOrEmpty($test.Description))
									{
										$xRow++
										$Table.Cell($xRow,1).Range.Text = "Description"
										$Table.Cell($xRow,2).Range.Text = $test.description
									}
									$xRow++
									$Table.Cell($xRow,1).Range.Text = "Interval"
									$Table.Cell($xRow,2).Range.Text = $test.interval
									$xRow++
									$Table.Cell($xRow,1).Range.Text = "Time-out"
									$Table.Cell($xRow,2).Range.Text = $test.timeout
									$xRow++
									$Table.Cell($xRow,1).Range.Text = "Threshold"
									$Table.Cell($xRow,2).Range.Text = $test.threshold
									$xRow++
									$Table.Cell($xRow,1).Range.Text = "Recovery Action"
									Switch ($test.RecoveryAction)
									{
										"AlertOnly"                     {$Table.Cell($xRow,2).Range.Text = "Alert Only"}
										"RemoveServerFromLoadBalancing" {$Table.Cell($xRow,2).Range.Text = "Remove Server from load balancing"}
										"RestartIma"                    {$Table.Cell($xRow,2).Range.Text = "Restart IMA"}
										"ShutdownIma"                   {$Table.Cell($xRow,2).Range.Text = "Shutdown IMA"}
										"RebootServer"                  {$Table.Cell($xRow,2).Range.Text = "Reboot Server"}
										Default {$Table.Cell($xRow,2).Range.Text = "Recovery Action could not be determined: $($test.RecoveryAction)"}
									}

									$Table.Rows.SetLeftIndent($Indent3TabStops,$wdAdjustProportional)
									$Table.AutoFitBehavior($wdAutoFitContent)

									FindWordDocumentEnd
								}
								$XML = $Null
								$xRow = $Null
								$Columns = $Null
								$Row = $Null
							}
							If($Setting.MaximumServersOfflinePercent.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Health Monitoring and Recovery"
								WriteWordLine 0 3 "Maximum % of servers with logon control: " $Setting.MaximumServersOfflinePercent.Value
							}
							If($Setting.CpuManagementServerLevel.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Memory/CPU\CPU management server level: "
								Switch ($Setting.CpuManagementServerLevel.Value)
								{
									"NoManagement" {WriteWordLine 0 3 "No CPU utilization management"}
									"Fair"         {WriteWordLine 0 3 "Fair sharing of CPU between sessions"}
									"Preferential" {WriteWordLine 0 3 "Preferential Load Balancing"}
									Default {WriteWordLine 0 3 "CPU management server level could not be determined: $($Setting.CpuManagementServerLevel.Value)"}
								}
							}
							If($Setting.MemoryOptimization.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Memory/CPU\Memory optimization: " $Setting.MemoryOptimization.State
							}
							If($Setting.MemoryOptimizationExcludedPrograms.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Memory/CPU\Memory optimization application exclusion list: "
								$array = $Setting.MemoryOptimizationExcludedPrograms.Values
								ForEach($element in $array)
								{
									WriteWordLine 0 3 $element
								}
								$array = $Null
							}
							If($Setting.MemoryOptimizationIntervalType.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Memory/CPU\Memory optimization interval: " -nonewline
								Switch ($Setting.MemoryOptimizationIntervalType.Value)
								{
									"AtStartup" {WriteWordLine 0 0 "Only at startup time"}
									"Daily"     {WriteWordLine 0 0 "Daily"}
									"Weekly"    {WriteWordLine 0 0 "Weekly"}
									"Monthly"   {WriteWordLine 0 0 "Monthly"}
									Default {WriteWordLine 0 0 " could not be determined: $($Setting.MemoryOptimizationIntervalType.Value)"}
								}
							}
							If($Setting.MemoryOptimizationDayOfMonth.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Memory/CPU\Memory optimization schedule: "
								WriteWordLine 0 3 "day of month: " $Setting.MemoryOptimizationDayOfMonth.Value
							}
							If($Setting.MemoryOptimizationDayOfWeek.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Memory/CPU\Memory optimization schedule: "
								WriteWordLine 0 3 "day of week: " $Setting.MemoryOptimizationDayOfWeek.Value
							}
							If($Setting.MemoryOptimizationTime.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Memory/CPU\Memory optimization schedule: "
								WriteWordLine 0 3 "time: " -nonewline
								$tmp = ConvertNumberToTime $Setting.MemoryOptimizationTime.Value
								WriteWordLine 0 0 $tmp
								$tmp = $Null
							}
							If($Setting.OfflineClientTrust.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Offline Applications\Offline app client trust: " $Setting.OfflineClientTrust.State
							}
							If($Setting.OfflineEventLogging.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Offline Applications\Offline app event logging: " $Setting.OfflineEventLogging.State
							}
							If($Setting.OfflineLicensePeriod.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Offline Applications\Offline app license period - Days: " $Setting.OfflineLicensePeriod.Value
							}
							If($Setting.OfflineUsers.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Offline Applications\Offline app users: " 
								$array = $Setting.OfflineUsers.Values
								ForEach($element in $array)
								{
									WriteWordLine 0 3 $element
								}
								$array = $Null
							}
							If($Setting.RebootCustomMessage.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Reboot Behavior\Reboot custom warning: " $Setting.RebootCustomMessage.State
							}
							If($Setting.RebootCustomMessageText.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Reboot Behavior\Reboot custom warning text: " 
								WriteWordLine 0 3 $Setting.RebootCustomMessageText.Value
							}
							If($Setting.RebootDisableLogOnTime.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Reboot Behavior\Reboot logon disable time: "
								Switch ($Setting.RebootDisableLogOnTime.Value)
								{
									"DoNotDisableLogOnsBeforeReboot" {WriteWordLine 0 3 "Do not disable logons before reboot"}
									"Disable5MinutesBeforeReboot"    {WriteWordLine 0 3 "Disable 5 minutes before reboot"}
									"Disable10MinutesBeforeReboot"   {WriteWordLine 0 3 "Disable 10 minutes before reboot"}
									"Disable15MinutesBeforeReboot"   {WriteWordLine 0 3 "Disable 15 minutes before reboot"}
									"Disable30MinutesBeforeReboot"   {WriteWordLine 0 3 "Disable 30 minutes before reboot"}
									"Disable60MinutesBeforeReboot"   {WriteWordLine 0 3 "Disable 60 minutes before reboot"}
									Default {WriteWordLine 0 3 "Reboot logon disable time could not be determined: $($Setting.RebootDisableLogOnTime.Value)"}
								}
							}
							If($Setting.RebootScheduleFrequency.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Reboot Behavior\Reboot schedule frequency - Days: " $Setting.RebootScheduleFrequency.Value
							}
							If($Setting.RebootScheduleStartDate.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Reboot Behavior\Reboot schedule start date: " -nonewline
								$Tmp = ConvertIntegerToDate $Setting.RebootScheduleStartDate.Value
								WriteWordLine 0 0 $Tmp
								$tmp = $Null
							}
							If($Setting.RebootScheduleTime.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Reboot Behavior\Reboot schedule time: " -nonewline
								$tmp = ConvertNumberToTime $Setting.RebootScheduleTime.Value 						
								WriteWordLine 0 0 $Tmp
								$tmp = $Null
							}
							If($Setting.RebootWarningInterval.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Reboot Behavior\Reboot warning interval: "
								Switch ($Setting.RebootWarningInterval.Value)
								{
									"Every1Minute"   {WriteWordLine 0 3 "Every 1 Minute"}
									"Every3Minutes"  {WriteWordLine 0 3 "Every 3 Minutes"}
									"Every5Minutes"  {WriteWordLine 0 3 "Every 5 Minutes"}
									"Every10Minutes" {WriteWordLine 0 3 "Every 10 Minutes"}
									"Every15Minutes" {WriteWordLine 0 3 "Every 15 Minutes"}
									Default {WriteWordLine 0 3 "Reboot warning interval could not be determined: $($Setting.RebootWarningInterval.Value)"}
								}
							}
							If($Setting.RebootWarningStartTime.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Reboot Behavior\Reboot warning start time: "
								Switch ($Setting.RebootWarningStartTime.Value)
								{
									"Start5MinutesBeforeReboot"  {WriteWordLine 0 3 "Start 5 Minutes Before Reboot"}
									"Start10MinutesBeforeReboot" {WriteWordLine 0 3 "Start 10 Minutes Before Reboot"}
									"Start15MinutesBeforeReboot" {WriteWordLine 0 3 "Start 15 Minutes Before Reboot"}
									"Start30MinutesBeforeReboot" {WriteWordLine 0 3 "Start 30 Minutes Before Reboot"}
									"Start60MinutesBeforeReboot" {WriteWordLine 0 3 "Start 60 Minutes Before Reboot"}
									Default {WriteWordLine 0 3 "Reboot warning start time could not be determined: $($Setting.RebootWarningStartTime.Value)"}
								}
							}
							If($Setting.RebootWarningMessage.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Reboot Behavior\Reboot warning to users: " $Setting.RebootWarningMessage.State
							}
							If($Setting.ScheduledReboots.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Reboot Behavior\Scheduled reboots: " $Setting.ScheduledReboots.State
							}
							Write-Verbose "$(Get-Date): `t`t`tVirtual IP"
							If($Setting.EnhancedCompatibility.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Virtual IP\Virtual IP enhanced compatibility: " $Setting.EnhancedCompatibility.State
							}
							If($Setting.FilterAdapterAddressesPrograms.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Virtual IP\Virtual IP filter adapter addresses programs list: " 
								$array = $Setting.FilterAdapterAddressesPrograms.Values
								ForEach($element in $array)
								{
									WriteWordLine 0 3 $element
								}
								$array = $Null
							}
							If($Setting.VirtualLoopbackSupport.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Virtual IP\Virtual IP loopback support: " $Setting.VirtualLoopbackSupport.State
							}
							If($Setting.VirtualLoopbackPrograms.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Virtual IP\Virtual IP virtual loopback programs list: " 
								$array = $Setting.VirtualLoopbackPrograms.Values
								ForEach($element in $array)
								{
									WriteWordLine 0 3 $element
								}
								$array = $Null
							}
							Write-Verbose "$(Get-Date): `t`t`tXML Service"
							If($Setting.TrustXmlRequests.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "XML Service\Trust XML requests: " $Setting.TrustXmlRequests.State
							}
							If($Setting.XmlServicePort.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "XML Service\XML service port: " $Setting.XmlServicePort.Value
							}
						}
						Else
						{
							Write-Verbose "$(Get-Date): `t`tUser Settings"
							WriteWordLine 0 1 "User settings:"
							Write-Verbose "$(Get-Date): `t`t`tICA"
							If($Setting.ClipboardRedirection.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Client clipboard redirection: " $Setting.ClipboardRedirection.State
							}
							If($Setting.DesktopLaunchForNonAdmins.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Desktop launches: " $Setting.DesktopLaunchForNonAdmins.State
							}
							If($Setting.NonPublishedProgramLaunching.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Launching of non-published programs during client connection: " $Setting.NonPublishedProgramLaunching.State
							}
							If($Setting.OemChannels.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\OEM Channels: " $Setting.OemChannels.State
							}
							Write-Verbose "$(Get-Date): `t`t`tICA\Audio"
							If($Setting.AudioQuality.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Audio\Audio quality: "
								Switch ($Setting.AudioQuality.Value)
								{
									"Low"    {WriteWordLine 0 3 "Low - for low-speed connections"}
									"Medium" {WriteWordLine 0 3 "Medium - optimized for speech"}
									"High"   {WriteWordLine 0 3 "High - high definition audio"}
									Default {WriteWordLine 0 3 "Audio quality could not be determined: $($Setting.AudioQuality.Value)"}
								}
							}
							If($Setting.ClientAudioRedirection.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Audio\Client audio redirection: " $Setting.ClientAudioRedirection.State
							}
							If($Setting.MicrophoneRedirection.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Audio\Client microphone redirection: " $Setting.MicrophoneRedirection.State
							}
							Write-Verbose "$(Get-Date): `t`t`tICA\Bandwidth"
							If($Setting.AudioBandwidthLimit.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\Audio redirection bandwidth limit (Kbps): " $Setting.AudioBandwidthLimit.Value
							}
							If($Setting.AudioBandwidthPercent.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\Audio redirection bandwidth limit %: " $Setting.AudioBandwidthPercent.Value
							}
							If($Setting.ClipboardBandwidthLimit.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\Clipboard redirection bandwidth limit (Kbps): " $Setting.ClipboardBandwidthLimit.Value
							}
							If($Setting.ClipboardBandwidthPercent.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\Clipboard redirection bandwidth limit %: " $Setting.ClipboardBandwidthPercent.Value
							}
							If($Setting.ComPortBandwidthLimit.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\COM port redirection bandwidth limit (Kbps): " $Setting.ComPortBandwidthLimit.Value
							}
							If($Setting.ComPortBandwidthPercent.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\COM port redirection bandwidth limit %: " $Setting.ComPortBandwidthPercent.Value
							}
							If($Setting.FileRedirectionBandwidthLimit.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\File redirection bandwidth limit (Kbps): " $Setting.FileRedirectionBandwidthLimit.Value
							}
							If($Setting.FileRedirectionBandwidthPercent.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\File redirection bandwidth limit %: " $Setting.FileRedirectionBandwidthPercent.Value
							}
							If($Setting.LptBandwidthLimit.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\LPT port redirection bandwidth limit (Kbps): " $Setting.LptBandwidthLimit.Value
							}
							If($Setting.LptBandwidthLimitPercent.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\LPT port redirection bandwidth limit %: " $Setting.LptBandwidthLimitPercent.Value
							}
							If($Setting.OemChannelBandwidthLimit.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\OEM channels bandwidth limit - Value: " $Setting.OemChannelBandwidthLimit.Value
							}
							If($Setting.OemChannelBandwidthPercent.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\OEM channels bandwidth limit percent - Value: " $Setting.OemChannelBandwidthPercent.Value
							}
							If($Setting.OverallBandwidthLimit.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\Overall session bandwidth limit (Kbps): " $Setting.OverallBandwidthLimit.Value
							}
							If($Setting.PrinterBandwidthLimit.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\Printer redirection bandwidth limit (Kbps): " $Setting.PrinterBandwidthLimit.Value
							}
							If($Setting.PrinterBandwidthPercent.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\Printer redirection bandwidth limit %: " $Setting.PrinterBandwidthPercent.Value
							}
							If($Setting.TwainBandwidthLimit.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\TWAIN device redirection bandwidth limit (Kbps): " $Setting.TwainBandwidthLimit.Value
							}
							If($Setting.TwainBandwidthPercent.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\TWAIN device redirection bandwidth limit %: " $Setting.TwainBandwidthPercent.Value
							}
							Write-Verbose "$(Get-Date): `t`t`tICA\Desktop"
							If($Setting.DesktopWallpaper.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Desktop UI\Desktop wallpaper: " $Setting.DesktopWallpaper.State
							}
							If($Setting.MenuAnimation.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Desktop UI\Menu animation: " $Setting.MenuAnimation.State
							}
							If($Setting.WindowContentsVisibleWhileDragging.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Desktop UI\View window contents while dragging: " $Setting.WindowContentsVisibleWhileDragging.State
							}
							Write-Verbose "$(Get-Date): `t`t`tICA\File Redirection"
							If($Setting.AutoConnectDrives.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\File Redirection\Auto connect client drives: " $Setting.AutoConnectDrives.State
							}
							If($Setting.ClientDriveRedirection.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\File Redirection\Client drive redirection: " $Setting.ClientDriveRedirection.State
							}
							If($Setting.ClientFixedDrives.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\File Redirection\Client fixed drives: " $Setting.ClientFixedDrives.State
							}
							If($Setting.ClientFloppyDrives.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\File Redirection\Client floppy drives: " $Setting.ClientFloppyDrives.State
							}
							If($Setting.ClientNetworkDrives.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\File Redirection\Client network drives: " $Setting.ClientNetworkDrives.State
							}
							If($Setting.ClientOpticalDrives.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\File Redirection\Client optical drives: " $Setting.ClientOpticalDrives.State
							}
							If($Setting.ClientRemoveableDrives.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\File Redirection\Client removable drives: " $Setting.ClientRemoveableDrives.State
							}
							If($Setting.HostToClientRedirection.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\File Redirection\Host to client redirection: " $Setting.HostToClientRedirection.State
							}
							If($Setting.ClientDriveLetterPreservation.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\File Redirection\Preserve client drive letters - Value: " $Setting.ClientDriveLetterPreservation.State
							}
							If($Setting.SpecialFolderRedirection.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\File Redirection\Special folder redirection: " $Setting.SpecialFolderRedirection.State
							}
							If($Setting.AsynchronousWrites.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\File Redirection\Use asynchronous writes: " $Setting.AsynchronousWrites.State
							}
							Write-Verbose "$(Get-Date): `t`t`tICA\Graphics"
							If($Setting.LossyCompressionLevel.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Graphics\Image compression\Lossy compression level: " -nonewline
								Switch ($Setting.LossyCompressionLevel.Value)
								{
									"None"   {WriteWordLine 0 0 "None"}
									"Low"    {WriteWordLine 0 0 "Low"}
									"Medium" {WriteWordLine 0 0 "Medium"}
									"High"   {WriteWordLine 0 0 "High"}
									Default {WriteWordLine 0 0 "Lossy compression level could not be determined: $($Setting.LossyCompressionLevel.Value)"}
								}
							}
							If($Setting.LossyCompressionThreshold.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Graphics\Image compression\Lossy compression "
								WriteWordLine 0 3 "threshold value (Kbps): " $Setting.LossyCompressionThreshold.Value
							}
							If($Setting.ProgressiveCompressionLevel.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Graphics\Image compression\Progressive compression level: " -nonewline
								Switch ($Setting.ProgressiveCompressionLevel.Value)
								{
									"UltraHigh" {WriteWordLine 0 0 "Ultra high"}
									"VeryHigh"  {WriteWordLine 0 0 "Very high"}
									"High"      {WriteWordLine 0 0 "High"}
									"Normal"    {WriteWordLine 0 0 "Normal"}
									"Low"       {WriteWordLine 0 0 "Low"}
									"None"      {WriteWordLine 0 0 "None"}
									Default {WriteWordLine 0 0 "Progressive compression level could not be determined: $($Setting.ProgressiveCompressionLevel.Value)"}
								}
							}
							If($Setting.ProgressiveCompressionThreshold.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Graphics\Image compression\Progressive compression "
								WriteWordLine 0 3 "threshold value (Kbps): " $Setting.ProgressiveCompressionThreshold.Value
							}
							If($Setting.ProgressiveHeavyweightCompression.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Graphics\Image compression\Progressive heavyweight compression: " 
								WriteWordLine 0 3 $Setting.ProgressiveHeavyweightCompression.State
							}
							Write-Verbose "$(Get-Date): `t`t`tICA\Multimedia"
							If($Setting.FlashAcceleration.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Multimedia\HDX MediaStream for Flash (client side)"
								WriteWordLine 0 3 "Flash acceleration: " $Setting.FlashAcceleration.State
							}
							If($Setting.FlashEventLogging.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Multimedia\HDX MediaStream for Flash (client side)"
								WriteWordLine 0 3 "Flash event logging: " $Setting.FlashEventLogging.State
							}
							If($Setting.FlashLatencyThreshold.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Multimedia\HDX MediaStream for Flash (client side)"
								WriteWordLine 0 3 "Flash latency threshold (milliseconds): " $Setting.FlashLatencyThreshold.Value
							}
							If($Setting.FlashServerSideContentFetchingWhitelist.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Multimedia\HDX MediaStream for Flash (client side)"
								WriteWordLine 0 3 "Flash server-side content fetching whitelist: "
								$Values = $Setting.FlashServerSideContentFetchingWhitelist.Values
								ForEach($Value in $Values)
								{
									WriteWordLine 0 4 $Value
								}
								$Values = $Null
							}
							If($Setting.FlashUrlBlacklist.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Multimedia\HDX MediaStream for Flash (client side)"
								WriteWordLine 0 3 "Flash URL blacklist " 
								$Values = $Setting.FlashUrlBlacklist.Values
								ForEach($Value in $Values)
								{
									WriteWordLine 0 4 $Value
								}
								$Values = $Null
							}
							If($Setting.AllowSpeedFlash.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Multimedia\HDX Multimedia For Flash (server side)"
								WriteWordLine 0 3 "Flash quality adjustment: "
								Switch ($Setting.AllowSpeedFlash.Value)
								{
									"NoOptimization"      {WriteWordLine 0 3 "Do not optimize Flash animation options"}
									"AllConnections"      {WriteWordLine 0 3 "Optimize Flash animation options for all connections"}
									"RestrictedBandwidth" {WriteWordLine 0 3 "Optimize Flash animation options for low bandwidth connections only"}
									Default {WriteWordLine 0 3 "Flash quality adjustment could not be determined: $($Setting.AllowSpeedFlash.Value)"}
								}
							}
							Write-Verbose "$(Get-Date): `t`t`tICA\Ports"
							If($Setting.ClientComPortsAutoConnection.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Ports\Auto connect client COM ports: " $Setting.ClientComPortsAutoConnection.State
							}
							If($Setting.ClientLptPortsAutoConnection.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Ports\Auto connect client LPT ports: " $Setting.ClientLptPortsAutoConnection.State
							}
							If($Setting.ClientComPortRedirection.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Ports\Client COM port redirection: " $Setting.ClientComPortRedirection.State
							}
							If($Setting.ClientLptPortRedirection.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Ports\Client LPT port redirection: " $Setting.ClientLptPortRedirection.State
							}
							Write-Verbose "$(Get-Date): `t`t`tICA\Printing"
							If($Setting.ClientPrinterRedirection.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Client printer redirection: " $Setting.ClientPrinterRedirection.State
							}
							If($Setting.DefaultClientPrinter.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Default printer - client's Default printer: " 
								Switch ($Setting.DefaultClientPrinter.Value)
								{
									"ClientDefault" {WriteWordLine 0 3 "Set Default printer to the client's main printer"}
									"DoNotAdjust"   {WriteWordLine 0 3 "Do not adjust the user's Default printer"}
									Default {WriteWordLine 0 0 "Default printer could not be determined: $($Setting.DefaultClientPrinter.Value)"}
								}
							}
							If($Setting.AutoCreationEventLogPreference.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Printer auto-creation event log preference: " 
								Switch ($Setting.AutoCreationEventLogPreference.Value)
								{
									"LogErrorsOnly"        {WriteWordLine 0 3 "Log errors only"}
									"LogErrorsAndWarnings" {WriteWordLine 0 3 "Log errors and warnings"}
									"DoNotLog"             {WriteWordLine 0 3 "Do not log errors or warnings"}
									Default {WriteWordLine 0 3 "Printer auto-creation event log preference could not be determined: $($Setting.AutoCreationEventLogPreference.Value)"}
								}
							}
							If($Setting.SessionPrinters.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Session printers\Session printers: "
								$valArray = $Setting.SessionPrinters.Values
								ForEach($printer in $valArray)
								{
									$prArray = $printer.Split(',')
									ForEach($element in $prArray)
									{
										If($element.SubString(0, 2) -eq "\\")
										{
											$index = $element.SubString(2).IndexOf('\')
											If($index -ge 0)
											{
												$server = $element.SubString(0, $index + 2)
												$share  = $element.SubString($index + 3)
												WriteWordLine 0 3 "Server`t`t: $server"
												WriteWordLine 0 3 "Shared Name`t: $share"
											}
										}
										Else
										{
											$tmp = $element.SubString(0, 4)
											Switch ($tmp)
											{
												"copi" 
												{
													$txt = "Copy count`t:"
													$index = $element.SubString(0).IndexOf('=')
													If($index -ge 0)
													{
														$tmp2 = $element.SubString($index + 1)
														WriteWordLine 0 3 "$txt $tmp2"
													}
												}
												"coll"
												{
													$txt = "Collate`t`t:"
													$index = $element.SubString(0).IndexOf('=')
													If($index -ge 0)
													{
														$tmp2 = $element.SubString($index + 1)
														WriteWordLine 0 3 "$txt $tmp2"
													}
												}
												"prin"
												{
													$txt = "Print Quality`t:"
													$index = $element.SubString(0).IndexOf('=')
													If($index -ge 0)
													{
														$tmp2 = $element.SubString($index + 1)
														WriteWordLine 0 3 "$txt " -nonewline
														Switch ($tmp2)
														{
															"150" {WriteWordLine 0 0 "150 dpi"}
															"300" {WriteWordLine 0 0 "300 dpi"}
															"600" {WriteWordLine 0 0 "600 dpi"}
															"75"  {WriteWordLine 0 0 "75 dpi"}
															Default {WriteWordLine 0 0 "Print Quality could not be determined: $($tmp2)"}
														}
													}
												}
												"orie"
												{
													$txt = "Orientation`t:"
													$index = $element.SubString(0).IndexOf('=')
													If($index -ge 0)
													{
														$tmp2 = $element.SubString($index + 1)
														WriteWordLine 0 3 "$txt " -nonewline
														Switch ($tmp2)
														{
															"portrait"  {WriteWordLine 0 0 "Portrait"}
															"landscape" {WriteWordLine 0 0 "Landscape"}
															Default {WriteWordLine 0 3 "Orientation could not be determined: $($Element)"}
														}
													}
												}
												"pape"
												{
													$txt = "Paper Size`t:"
													$index = $element.SubString(0).IndexOf('=')
													If($index -ge 0)
													{
														$tmp2 = $element.SubString($index + 1)
														WriteWordLine 0 3 "$txt " -nonewline
														Switch ($tmp2)
														{
															1   {WriteWordLine 0 0 "Letter"}
															5   {WriteWordLine 0 0 "Legal"}
															7   {WriteWordLine 0 0 "Executive"}
															9   {WriteWordLine 0 0 "A4"}
															11  {WriteWordLine 0 0 "A5"}
															13  {WriteWordLine 0 0 "B5 (JIS)"}
															14  {WriteWordLine 0 0 "Folio"}
															20  {WriteWordLine 0 0 "Envelope #10"}
															27  {WriteWordLine 0 0 "Envelope DL"}
															28  {WriteWordLine 0 0 "Envelope C5"}
															34  {WriteWordLine 0 0 "Envelope B5"}
															37  {WriteWordLine 0 0 "Envelope Monarch"}
															43  {WriteWordLine 0 0 "Japanese Postcard"}
															70  {WriteWordLine 0 0 "A6"}
															#Default {WriteWordLine 0 3 "Paper Size could not be determined: $($element)"}
															Default 
															{
																WriteWordLine 0 0 "Custom Paper Size"
															}
														}
													}
												}
												"widt"
												{
													$index = $element.SubString(0).IndexOf('=')
													If($index -ge 0)
													{
														$tmp2 = $element.SubString($index + 1)
														If($tmp2.length -gt 0)
														{
															WriteWordLine 0 3 "Custom Paper Width`t: $($tmp2) (Millimeters)" 
														}
													}
												}
												"heig"
												{
													$index = $element.SubString(0).IndexOf('=')
													If($index -ge 0)
													{
														$tmp2 = $element.SubString($index + 1)
														If($tmp2.length -gt 0)
														{
															WriteWordLine 0 3 "Custom Paper Height`t: $($tmp2) (Millimeters)" 
														}
													}
												}
												"form"
												{
													$txt = "Form Name:"
													$index = $element.SubString(0).IndexOf('=')
													If($index -ge 0)
													{
														$tmp2 = $element.SubString($index + 1)
														If($tmp2.length -gt 0)
														{
															WriteWordLine 0 3 "$txt $tmp2"
														}
													}
												}
												"mode" 
												{
													$txt = "Printer Model`t:"
													$index = $element.SubString(0).IndexOf('=')
													If($index -ge 0)
													{
														$tmp2 = $element.SubString($index + 1)
														WriteWordLine 0 3 "$txt $tmp2"
													}
												}
												"loca" 
												{
													$txt = "Location`t:"
													$index = $element.SubString(0).IndexOf('=')
													If($index -ge 0)
													{
														$tmp2 = $element.SubString($index + 1)
														If($tmp2.length -gt 0)
														{
															WriteWordLine 0 3 "$txt $tmp2"
														}
													}
												}
												"appl"
												{
													WriteWordLine 0 3 "Apply customized settings at every logon"
												}
												Default {WriteWordLine 0 3 "Session printer setting could not be determined: $($Element)"}
											}
										}
									}
									WriteWordLine 0 0 ""
									$ValArray = $Null
									$prarray = $Null
									$txt = $Null
									$index = $Null
									$tmp2 = $Null
									$Server = $Null
									$Share = $Null
								}
							}
							If($Setting.WaitForPrintersToBeCreated.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Wait for printers to be created (desktop): " $Setting.WaitForPrintersToBeCreated.State
							}
							If($Setting.ClientPrinterAutoCreation.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Client Printers\Auto-create client printers: "
								Switch ($Setting.ClientPrinterAutoCreation.Value)
								{
									"DoNotAutoCreate"    {WriteWordLine 0 3 "Do not auto-create client printers"}
									"DefaultPrinterOnly" {WriteWordLine 0 3 "Auto-create the client's Default printer only"}
									"LocalPrintersOnly"  {WriteWordLine 0 3 "Auto-create local (non-network) client printers only"}
									"AllPrinters"        {WriteWordLine 0 3 "Auto-create all client printers"}
									Default {WriteWordLine 0 3 "Auto-create client printers could not be determined: $($Setting.ClientPrinterAutoCreation.Value)"}
								}
							}
							If($Setting.ClientPrinterNames.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Client Printers\Client printer names: " 
								Switch ($Setting.ClientPrinterNames.Value)
								{
									"StandardPrinterNames" {WriteWordLine 0 3 "Standard printer names"}
									"LegacyPrinterNames"   {WriteWordLine 0 3 "Legacy printer names"}
									Default {WriteWordLine 0 3 "Client printer names could not be determined: $($Setting.ClientPrinterNames.Value)"}
								}
							}
							If($Setting.DirectConnectionsToPrintServers.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Client Printers\Direct connections to print servers: " $Setting.DirectConnectionsToPrintServers.State
							}
							If($Setting.PrinterPropertiesRetention.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Client Printers\Printer properties retention: " 
								Switch ($Setting.PrinterPropertiesRetention.Value)
								{
									"SavedOnClientDevice"   {WriteWordLine 0 3 "Saved on the client device only"}
									"RetainedInUserProfile" {WriteWordLine 0 3 "Retained in user profile only"}
									"FallbackToProfile"     {WriteWordLine 0 3 "Held in profile only if not saved on client"}
									"DoNotRetain"           {WriteWordLine 0 3 "Do not retain printer properties"}
									Default {WriteWordLine 0 3 "Printer properties retention could not be determined: $($Setting.PrinterPropertiesRetention.Value)"}
								}
							}
							If($Setting.RetainedAndRestoredClientPrinters.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Client Printers\Retained and restored client printers: " $Setting.RetainedAndRestoredClientPrinters.State
							}
							If($Setting.InboxDriverAutoInstallation.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Drivers\Automatic installation of in-box printer drivers: " $Setting.InboxDriverAutoInstallation.State
							}
							If($Setting.PrinterDriverMappings.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Drivers\Printer driver mapping and compatibility: " 
								$array = $Setting.PrinterDriverMappings.Values
								$array = $Setting.PrinterDriverMappings.Values
								Write-Verbose "$(Get-Date): `t`t`t`tCreate table for printer drive mapping"
								$TableRange = $doc.Application.Selection.Range
								[int]$Columns = 3
								[int]$Rows = $array.count + 1
								Write-Verbose "$(Get-Date): `t`t`t`t`tAdd table to doc"
								$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
								$Table.rows.first.headingformat = $wdHeadingFormatTrue
								$Table.Style = $myHash.Word_TableGrid
								$Table.Borders.InsideLineStyle = $wdLineStyleSingle
								$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
								[int]$xRow = 1
								Write-Verbose "$(Get-Date): `t`t`t`t`tFormat first row with column headings"
								$Table.Rows.First.Shading.BackgroundPatternColor = $wdColorGray15
								$Table.Cell($xRow,1).Range.Font.Bold = $True
								$Table.Cell($xRow,1).Range.Text = "Driver Name"
								$Table.Cell($xRow,2).Range.Font.Bold = $True
								$Table.Cell($xRow,2).Range.Text = "Action"
								$Table.Cell($xRow,3).Range.Font.Bold = $True
								$Table.Cell($xRow,3).Range.Text = "Server Driver"
								ForEach($element in $array)
								{
									$Items = $element.Split(',')
									$xRow++
									Write-Verbose "$(Get-Date): `t`t`t`t`t`tProcessing row for $($Items[0])"
									$DriverName = $Items[0]
									$Action = $Items[1]
									If($Action -match 'Replace=')
									{
										$ServerDriver = $Action.substring($Action.indexof("=")+1)
										$Action = "Replace with"
									}
									Else
									{
										$ServerDriver = ""
										If($Action -eq "Allow")
										{
											$Action = "Allow"
										}
										ElseIf($Action -eq "Deny")
										{
											$Action = "Do not create"
										}
										ElseIf($Action -eq "UPD_Only")
										{
											$Action = "Create with universal driver"
										}
									}
									$Table.Cell($xRow,1).Range.Text = $DriverName
									$Table.Cell($xRow,2).Range.Text = $Action
									$Table.Cell($xRow,3).Range.Text = $ServerDriver
								}
								$array = $Null
								$Table.Rows.SetLeftIndent($Indent3TabStops,$wdAdjustProportional)
								$Table.AutoFitBehavior($wdAutoFitContent)

								FindWordDocumentEnd
								WriteWordLine 0 0 ""
							}
							If($Setting.GenericUniversalPrinterAutoCreation.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Universal Printing\Auto-create generic universal printer: " $Setting.GenericUniversalPrinterAutoCreation.State
							}
							If($Setting.UniversalDriverPriority.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Universal Printing\Universal driver priority: " 
								$TmpArray = $Setting.UniversalDriverPriority.Value.Split(';')
								ForEach($Thing in $TmpArray)
								{
									WriteWordLine 0 3 $Thing
								}
								$TmpArray = $Null
							}
							If($Setting.UniversalPrinting.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Universal Printing\Universal printing: " 
								Switch ($Setting.UniversalPrinting.Value)
								{
									"SpecificOnly"       {WriteWordLine 0 3 "Use only printer model specific drivers"}
									"UpdOnly"            {WriteWordLine 0 3 "Use universal printing only"}
									"FallbackToUpd"      {WriteWordLine 0 3 "Use universal printing only if requested driver is unavailable"}
									"FallbackToSpecific" {WriteWordLine 0 3 "Use printer model specific drivers only if universal printing is unavailable"}
									Default {WriteWordLine 0 3 "Universal print driver usage could not be determined: $($Setting.UniversalPrinting.Value)"}
								}
							}
							If($Setting.UniversalPrintingPreviewPreference.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Universal Printing\Universal printing preview preference: " 
								Switch ($Setting.UniversalPrintingPreviewPreference.Value)
								{
									"NoPrintPreview"        {WriteWordLine 0 3 "Do not use print preview for auto-created or generic universal printers"}
									"AutoCreatedOnly"       {WriteWordLine 0 3 "Use print preview for auto-created printers only"}
									"GenericOnly"           {WriteWordLine 0 3 "Use print preview for generic universal printers only"}
									"AutoCreatedAndGeneric" {WriteWordLine 0 3 "Use print preview for both auto-created and generic universal printers"}
									Default {WriteWordLine 0 3 "Universal printing preview preference could not be determined: $($Setting.UniversalPrintingPreviewPreference.Value)"}
								}
							}
							Write-Verbose "$(Get-Date): `t`t`tICA\Security"
							If($Setting.MinimumEncryptionLevel.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Security\SecureICA minimum encryption level: " 
								Switch ($Setting.MinimumEncryptionLevel.Value)
								{
									"Unknown" {WriteWordLine 0 3 "Unknown encryption"}
									"Basic"   {WriteWordLine 0 3 "Basic"}
									"LogOn"   {WriteWordLine 0 3 "RC5 (128 bit) logon only"}
									"Bits40"  {WriteWordLine 0 3 "RC5 (40 bit)"}
									"Bits56"  {WriteWordLine 0 3 "RC5 (56 bit)"}
									"Bits128" {WriteWordLine 0 3 "RC5 (128 bit)"}
									Default {WriteWordLine 0 3 "SecureICA minimum encryption level could not be determined: $($Setting.MinimumEncryptionLevel.Value)"}
								}
							}
							Write-Verbose "$(Get-Date): `t`t`tICA\Session Limits"
							If($Setting.ConcurrentLogOnLimit.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Session limits\Concurrent logon limit: " $Setting.ConcurrentLogOnLimit.Value
							}
							If($Setting.SessionDisconnectTimer.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Session Limits\Disconnected session timer: " $Setting.SessionDisconnectTimer.State
							}
							If($Setting.SessionDisconnectTimerInterval.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Session Limits\Disconnected session timer interval (minutes): " $Setting.SessionDisconnectTimerInterval.Value
							}
							If($Setting.SessionConnectionTimer.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Session Limits\Session connection timer: " $Setting.SessionConnectionTimer.State
							}
							If($Setting.SessionConnectionTimerInterval.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Session Limits\Session connection timer interval - Value (minutes): " $Setting.SessionConnectionTimerInterval.Value
							}
							If($Setting.SessionIdleTimer.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Session Limits\Session idle timer: " $Setting.SessionIdleTimer.State
							}
							If($Setting.SessionIdleTimerInterval.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Session Limits\Session idle timer interval - Value (minutes): " $Setting.SessionIdleTimerInterval.Value
							}
							Write-Verbose "$(Get-Date): `t`t`tICA Shadowing"
							If($Setting.ShadowInput.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Shadowing\Input from shadow connections: " $Setting.ShadowInput.State
							}
							If($Setting.ShadowLogging.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Shadowing\Log shadow attempts: " $Setting.ShadowLogging.State
							}
							If($Setting.ShadowUserNotification.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Shadowing\Notify user of pending shadow connections: " $Setting.ShadowUserNotification.State
							}
							If($Setting.ShadowAllowList.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Shadowing\Users who can shadow other users: " 
								$array = $Setting.ShadowAllowList.Values
								#gui only shows computer\account or domain\account
								#what is stored is:
								#0x05/NT/XA65\ANON000/S-1-5-21-1307341077-4083623718-4268213518-1028 (workgroup/local)
								#0x05/NT/XA651\CTX_CPUUSER/S-1-5-21-1200344839-3835835227-1016768578-1002 (domain/local)
								#0x05/NT/WEBSTERSLAB\ADMINISTRATOR/S-1-5-21-3679396586-1061193519-2853834051-500 (domain user)
								#0x05/NT/WEBSTERSLAB\DOMAIN ADMINS/S-1-5-21-3679396586-1061193519-2853834051-512 (domain group)
								#we only need the computer\account or domain\account
								#first 9 characters are 0x05/NT/ for all account types
								#since PoSH starts counting at 0 we don't need the first 9 characters
								#Then we need the position of the first / after the computer\account
								#what is left between the two is what we need
								ForEach($element in $array)
								{
									$x = $element.indexof("/",8)
									$tmp = $element.substring(8,$x-8)
									WriteWordLine 0 3 $tmp
								}
								$x = $Null
								$tmp = $Null
							}
							If($Setting.ShadowDenyList.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Shadowing\Users who cannot shadow other users: " 
								$array = $Setting.ShadowDenyList.Values
								ForEach($element in $array)
								{
									$x = $element.indexof("/",8)
									$tmp = $element.substring(8,$x-8)
									WriteWordLine 0 3 $tmp
								}
								$x = $Null
								$tmp = $Null
							}
							Write-Verbose "$(Get-Date): `t`t`tICA\Time Zone Control"
							If($Setting.LocalTimeEstimation.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Time Zone Control\Local Time Estimation: " $Setting.LocalTimeEstimation.State
							}
							If($Setting.SessionTimeZone.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Time Zone Control\Use local time of client: " 
								Switch ($Setting.SessionTimeZone.Value)
								{
									"UseServerTimeZone" {WriteWordLine 0 3 "Use server time zone"}
									"UseClientTimeZone" {WriteWordLine 0 3 "Use client time zone"}
									Default {WriteWordLine 0 3 "Use local time of client could not be determined: $($Setting.SessionTimeZone.Value)"}
								}
							}
							Write-Verbose "$(Get-Date): `t`t`tICA\TWAIN Devices"
							If($Setting.TwainRedirection.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\TWAIN devices\Client TWAIN device redirection: " $Setting.TwainRedirection.State
							}
							If($Setting.TwainCompressionLevel.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\TWAIN devices\TWAIN compression level: " -nonewline
								Switch ($Setting.TwainCompressionLevel.Value)
								{
									"None"   {WriteWordLine 0 0 "None"}
									"Low"    {WriteWordLine 0 0 "Low"}
									"Medium" {WriteWordLine 0 0 "Medium"}
									"High"   {WriteWordLine 0 0 "High"}
									Default {WriteWordLine 0 0 "TWAIN compression level could not be determined: $($Setting.TwainCompressionLevel.Value)"}
								}
							}
							Write-Verbose "$(Get-Date): `t`t`tICA\USB devices"
							If($Setting.UsbDeviceRedirection.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\USB devices\Client USB device redirection: " $Setting.UsbDeviceRedirection.State
							}
							If($Setting.UsbDeviceRedirectionRules.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\USB devices\Client USB device redirection rules: " 
								$array = $Setting.UsbDeviceRedirectionRules.Values
								ForEach($element in $array)
								{
									WriteWordLine 0 3 $element
								}
								$array = $Null
							}
							If($Setting.UsbPlugAndPlayRedirection.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\USB devices\Client USB Plug and Play device redirection: " $Setting.UsbPlugAndPlayRedirection.State
							}
							If($Setting.SessionImportance.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Session Settings\Session importance: " -nonewline
								Switch ($Setting.SessionImportance.Value)
								{
									"Low"    {WriteWordLine 0 0 "Low"}
									"Normal" {WriteWordLine 0 0 "Normal"}
									"High"   {WriteWordLine 0 0 "High"}
									Default {WriteWordLine 0 0 "Session importance could not be determined: $($Setting.SessionImportance.Value)"}
								}
							}
							Write-Verbose "$(Get-Date): `t`t`tServer Session Settings"
							If($Setting.SingleSignOn.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Session Settings\Single Sign-On: " $Setting.SingleSignOn.State
							}
							If($Setting.SingleSignOnCentralStore.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Session Settings\Single Sign-On central store: " $Setting.SingleSignOnCentralStore.Value
							}
						}
					}
					WriteWordLine 0 0 ""
				}
				Else
				{
					WriteWordLine 0 1 "Unable to retrieve settings"
				}
				$Filter = $Null
				$Settings = $Null
				Write-Verbose "$(Get-Date): `t`tFinished $($Policy.PolicyName)`t$($Policy.Type)"
				Write-Verbose "$(Get-Date): "
			}
			Else
			{
				WriteWordLine 0 1 "$($Policy.Type) Policy - $($Policy.PolicyName)"
				If($xDriveName -eq "")
				{
					$Global:TotalIMAPolicies++
				}
				Else
				{
					$Global:TotalADPolicies++
				}
				$Global:TotalPolicies++
			}
		}
	}
	ElseIf(!$?)
	{
		Write-Warning "Citrix Policy information could not be retrieved."
	}
	Else
	{
		Write-Warning "No results returned for Citrix Policy information"
	}

	$Policies = $Null
	If($xDriveName -ne "")
	{
		Write-Verbose "$(Get-Date): `tRemoving ADGpoDrv PSDrive"
		Remove-PSDrive ADGpoDrv -EA 0 4>$Null
		Write-Verbose "$(Get-Date): "
	}
}

Function GetCtxGPOsInAD
{
	#thanks to the Citrix Engineering Team for pointers and for Michael B. Smith for creating the function
	#updated 07-Nov-13 to work in a Windows Workgroup environment
	Write-Verbose "$(Get-Date): Testing for an Active Directory environment"
	$root = [ADSI]"LDAP://RootDSE"
	If([String]::IsNullOrEmpty($root.PSBase.Name))
	{
		Write-Verbose "$(Get-Date): Not in an Active Directory environment"
		$root = $Null
		$xArray = @()
	}
	Else
	{
		Write-Verbose "$(Get-Date): In an Active Directory environment"
		$domainNC = $root.defaultNamingContext.ToString()
		$root = $Null
		$xArray = @()

		$domain = $domainNC.Replace( 'DC=', '' ).Replace( ',', '.' )
		Write-Verbose "$(Get-Date): Searching \\$($domain)\sysvol\$($domain)\Policies"
		$sysvolFiles = @()
		$sysvolFiles = dir -Recurse ( '\\' + $domain  + '\sysvol\' + $domain + '\Policies' ) -EA 0
		If($sysvolFiles.Count -eq 0)
		{
			Write-Verbose "$(Get-Date): Search timed out.  Retrying.  Searching \\ + $($domain)\sysvol\$($domain)\Policies a second time."
			$sysvolFiles = dir -Recurse ( '\\' + $domain  + '\sysvol\' + $domain + '\Policies' ) -EA 0
		}
		foreach( $file in $sysvolFiles )
		{
			If( -not $file.PSIsContainer )
			{
				#$file.FullName  ### name of the policy file
				If( $file.FullName -like "*\Citrix\GroupPolicy\Policies.gpf" )
				{
					#"have match " + $file.FullName ### name of the Citrix policies file
					$array = $file.FullName.Split( '\' )
					If( $array.Length -gt 7 )
					{
						$gp = $array[ 6 ].ToString()
						$gpObject = [ADSI]( "LDAP://" + "CN=" + $gp + ",CN=Policies,CN=System," + $domainNC )
						$xArray += $gpObject.DisplayName	### name of the group policy object
					}
				}
			}
		}
	}
	Return ,$xArray
}

Function BuildTableForServerOrWG
{
	Param([Array]$xArray, [String]$xType)
	
	#divide by 0 bug reported 9-Apr-2014 by Lee Dehmer 
	#if security group name or OU name was longer than 60 characters it caused a divide by 0 error
	
	#added a second parameter to the function so the verbose message would say whether 
	#the function is processing servers, security groups or OUs.
	
	If(-not ($xArray -is [Array]))
	{
		$xArray = (,$xArray)
	}
	[int]$MaxLength = 0
	[int]$TmpLength = 0
	#remove 60 as a hard-coded value
	#60 is the max width the table can be when indented 36 points
	[int]$MaxTableWidth = 60
	ForEach($xName in $xArray)
	{
		$TmpLength = $xName.Length
		If($TmpLength -gt $MaxLength)
		{
			$MaxLength = $TmpLength
		}
	}
	Write-Verbose "$(Get-Date): `t`tMax length of $xType name is $($MaxLength)"
	$TableRange = $doc.Application.Selection.Range
	#removed hard-coded value of 60 and replace with MaxTableWidth variable
	[int]$Columns = [Math]::Floor($MaxTableWidth / $MaxLength)
	If($xArray.count -lt $Columns)
	{
		[int]$Rows = 1
		#not enough array items to fill columns so use array count
		$MaxCells  = $xArray.Count
		#reset column count so there are no empty columns
		$Columns   = $xArray.Count 
	}
	ElseIf($Columns -eq 0)
	{
		#divide by 0 bug if this condition is not handled
		#number was larger than $MaxTableWidth so there can only be one column
		#with one cell per row
		[int]$Rows = $xArray.count
		$Columns   = 1
		$MaxCells  = 1
	}
	Else
	{
		[int]$Rows = [Math]::Floor( ( $xArray.count + $Columns - 1 ) / $Columns)
		#more array items than columns so don't go past last column
		$MaxCells  = $Columns
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
	$Script:Word = New-Object -comobject "Word.Application" -EA 0 4>$Null
	
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
		Remove-Item $Script:FileName1 4>$Null
	}
	Write-Verbose "$(Get-Date): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
	If(Test-Path variable:global:word)
	{
		Remove-Variable -Name word -Scope Global 4>$Null
	}
	$SaveFormat = $Null
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
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
	If($AddDateTime)
	{
		$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).html"
	}
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
	ElseIf($Text)
	{
		If(!$AddDateTime)
		{
			[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName).txt"
		}
	}
	ElseIf($HTML)
	{
		If(!$AddDateTime)
		{
			[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName).html"
		}
	}
}

#Script begins

$script:startTime = Get-date

If($TEXT)
{
	$global:output = ""
}

If(!(Check-NeededPSSnapins "Citrix.Common.Commands","Citrix.XenApp.Commands"))
{
	#We're missing Citrix Snapins that we need
	$ErrorActionPreference = $SaveEAPreference
	Write-Error "Missing Citrix PowerShell Snap-ins Detected, check the console above for more information. Are you sure you are running this script on a XenApp 6 Server? Script will now close."
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

# Get farm information
Write-Verbose "$(Get-Date): Getting initial Farm data"
$farm = Get-XAFarm -EA 0

If($? -and $Farm -ne $Null)
{
	Write-Verbose "$(Get-Date): Verify farm version"
	#first check to make sure this is a XenApp 6 farm
	If($Farm.ServerVersion.ToString().SubString(0,1) -eq "6")
	{
		#this is a XenApp 6 farm, script can proceed
	}
	Else
	{
		#this is not a XenApp 6 farm, script cannot proceed
		Write-Warning "This script is designed for XenApp 6 and should not be run on XenApp 5"
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

If(!$Summary -and ($Section -eq "All" -or $Section -eq "ConfigLog"))
{
	Write-Verbose "$(Get-Date): Processing Configuration Logging"
	[bool]$ConfigLog = $False
	$ConfigurationLogging = Get-XAConfigurationLog -EA 0

	If($? -and $ConfigurationLogging -ne $Null)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Configuration Logging"
		If($ConfigurationLogging.LoggingEnabled) 
		{
			$ConfigLog = $True
			WriteWordLine 0 1 "Configuration Logging is enabled."
			WriteWordLine 0 1 "Allow changes to the farm when logging database is disconnected: " $ConfigurationLogging.ChangesWhileDisconnectedAllowed
			WriteWordLine 0 1 "Require administrator to enter credentials before clearing the log: " $ConfigurationLogging.CredentialsOnClearLogRequired
			WriteWordLine 0 1 "Database type: " $ConfigurationLogging.DatabaseType
			WriteWordLine 0 1 "Authentication mode: " $ConfigurationLogging.AuthenticationMode
			WriteWordLine 0 1 "Connection string: " 
			$Tmp = "`t`t" + $ConfigurationLogging.ConnectionString.replace(";","`n`t`t`t")
			WriteWordLine 0 1 $Tmp -NoNewline
			WriteWordLine 0 0 ""
			WriteWordLine 0 1 "User name: " $ConfigurationLogging.UserName
			$Tmp = $Null
		}
		Else 
		{
			WriteWordLine 0 1 "Configuration Logging is disabled."
		}
	}
	ElseIf(!$?)
	{
		Write-Warning  "Configuration Logging could not be retrieved"
	}
	Else
	{
		Write-Warning  "No results returned for Configuration Logging"
	}
	$ConfigurationLogging = $Null
	Write-Verbose "$(Get-Date): Finished Configuration Logging"
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
				WriteWordLine 0 1 "Administrator account is " -NoNewLine
				If($Administrator.Enabled)
				{
					WriteWordLine 0 0 "Enabled" 
				} 
				Else
				{
					WriteWordLine 0 0 "Disabled" 
				}
				If($Administrator.AdministratorType -eq "Custom") 
				{
					Write-Verbose "$(Get-Date): `t`t`tProcessing farm privileges"
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
							Default {WriteWordLine 0 2 "Farm privileges could not be determined: $($farmprivilege)"}
						}
					}
			
					WriteWordLine 0 1 "Folder Privileges:"
					ForEach($folderprivilege in $Administrator.FolderPrivileges) 
					{
						#The Citrix PoSH cmdlet only returns data for three folders:
						#Servers
						#WorkerGroups
						#Applications
						
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
								Default {WriteWordLine 0 3 "Folder permission could not be determined: $($folderpermissions)"}
							}
						}
					}
				}		
				#WriteWordLine 0 0 ""
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
				If($? -and $AppServerInfo -ne $Null)
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
					#if streamed, OffWriteWordLine 0 access properties
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
							WriteWordLine 0 1 "Servers:"
							$TempArray = $AppServerInfo.ServerNames | Sort
							BuildTableForServerOrWG $TempArray "Server"
							$TempArray = $Null
						}
						If(![String]::IsNullOrEmpty($AppServerInfo.WorkerGroupNames))
						{
							WriteWordLine 0 1 "Worker Groups:"
							$TempArray = $AppServerInfo.WorkerGroupNames | Sort
							BuildTableForServerOrWG $TempArray "Server"
							$TempArray = $Null
						}
					}
					Else
					{
						WriteWordLine 0 2 "Unable to retrieve a list of Servers or Worker Groups for this application"
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
							WriteWordLine 0 2 $filetype
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
						WriteWordLine 0 1 "Run app as a least-privileged user account`t: " $Application.RunAsLeastPrivilegedUser
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
					WriteWordLine 0 0 "Disabled"
				} 
				Else
				{
					WriteWordLine 0 0 "Enabled"
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
						"Unknown"     {WriteWordLine 0 0 "Unknown color depth"}
						"Colors8Bit"  {WriteWordLine 0 0 "256-color (8-bit)"}
						"Colors16Bit" {WriteWordLine 0 0 "Better Speed (16-bit)"}
						"Colors32Bit" {WriteWordLine 0 0 "Better Appearance (32-bit)"}
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
		Write-Warning "Application information could not be retrieved"
	}
	$Applications = $Null
	Write-Verbose "$(Get-Date): Finished Processing Applications"
	Write-Verbose "$(Get-Date): "
}

If(!$Summary -and ($Section -eq "All" -or $Section -eq "ConfigLog"))
{
	Write-Verbose "$(Get-Date): Setting summary variables"
	[int]$TotalConfigLogItems = 0

	Write-Verbose "$(Get-Date): Processing Configuration Logging/History Report"
	If($ConfigLog)
	{
		#history AKA Configuration Logging report
		#only process if $ConfigLog = $True and XA6ConfigLog.udl file exists
		#build connection string
		#User ID is account that has access permission for the configuration logging database
		#Initial Catalog is the name of the Configuration Logging SQL Database
		#bug fixed by Esther Barthel
		If(Test-Path "$($pwd.path)\XA6ConfigLog.udl")
		{
			Write-Verbose "$(Get-Date): `tRetrieving logging data for date range $($StartDate) through $($EndDate)"
			$ConnectionString = Get-Content "$($pwd.path)\XA6ConfigLog.udl" 4>$Null| select-object -last 1
			$ConfigLogReport = Get-CtxConfigurationLogReport -connectionstring $ConnectionString -TimePeriodFrom $StartDate -TimePeriodTo $EndDate -EA 0 4>$Null

			If($? -and $ConfigLogReport -ne $Null)
			{
				Write-Verbose "$(Get-Date): `tProcessing $($ConfigLogReport.Count) history items"
				$selection.InsertNewPage()
				WriteWordLine 1 0 "History:"
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
			ElseIf($? -and $ConfigLogReport -eq $Null)
			{
				Write-Verbose "$(Get-Date): There was no configuration logging data returned"
				WriteWordLine 0 0 "There was no configuration logging data returned"
			}
			ElseIf(!$?)
			{
				Write-Warning "Error retrieving configuration logging data"
				WriteWordLine 0 0 "Error retrieving configuration logging data"
			}
		}
		Else 
		{
			WriteWordLine 1 0 "Configuration Logging is enabled but the XA6ConfigLog.udl file was not found"
		}
		$ConnectionString = $Null
		$ConfigLogReport = $Null
	}
	Write-Verbose "$(Get-Date): Finished Processing Configuration Logging/History Report"
	Write-Verbose "$(Get-Date): "
}

If($Section -eq "All" -or $Section -eq "LBPolicies")
{
	#load balancing policies
	Write-Verbose "$(Get-Date): Processing Load Balancing Policies"
	Write-Verbose "$(Get-Date): `tSetting summary variables"
	[int]$TotalLBPolicies = 0

	Write-Verbose "$(Get-Date): `tRetrieving Load Balancing Policies"
	$LoadBalancingPolicies = Get-XALoadBalancingPolicy -EA 0 | Sort PolicyName

	If($? -and $LoadBalancingPolicies -ne $Null)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Load Balancing Policies:"
		ForEach($LoadBalancingPolicy in $LoadBalancingPolicies)
		{
			$TotalLBPolicies++
			Write-Verbose "$(Get-Date): `t`tProcessing Load Balancing Policy $($LoadBalancingPolicy.PolicyName)"
			$LoadBalancingPolicyConfiguration = Get-XALoadBalancingPolicyConfiguration -PolicyName $LoadBalancingPolicy.PolicyName -EA 0
			$LoadBalancingPolicyFilter = Get-XALoadBalancingPolicyFilter -PolicyName $LoadBalancingPolicy.PolicyName -EA 0
		
			If(!$Summary)
			{
				WriteWordLine 2 0 $LoadBalancingPolicy.PolicyName
				If(![String]::IsNullOrEmpty($LoadBalancingPolicy.Description))
				{
					WriteWordLine 0 1 "Description`t: " $LoadBalancingPolicy.Description
				}
				WriteWordLine 0 1 "Enabled`t`t: " -nonewline
				If($LoadBalancingPolicy.Enabled)
				{
					WriteWordLine 0 0 "Yes"
				}
				Else
				{
					WriteWordLine 0 0 "No"
				}
				WriteWordLine 0 1 "Priority`t`t: " $LoadBalancingPolicy.Priority
			
				WriteWordLine 0 1 "Filter based on Access Control: " -nonewline
				If($LoadBalancingPolicyFilter.AccessControlEnabled)
				{
					WriteWordLine 0 0 "Yes"
				}
				Else
				{
					WriteWordLine 0 0 "No"
				}
				If($LoadBalancingPolicyFilter.AccessControlEnabled)
				{
					WriteWordLine 0 1 "Apply to connections made through Access Gateway: " -nonewline
					If($LoadBalancingPolicyFilter.AllowConnectionsThroughAccessGateway)
					{
						WriteWordLine 0 0 "Yes"
					}
					Else
					{
						WriteWordLine 0 0 "No"
					}
					If($LoadBalancingPolicyFilter.AllowConnectionsThroughAccessGateway)
					{
						If($LoadBalancingPolicyFilter.AllowOtherConnections)
						{
							WriteWordLine 0 2 "Any connection"
						} 
						Else
						{
							WriteWordLine 0 2 "Any connection that meets any of the following filters"
							If($LoadBalancingPolicyFilter.AccessSessionConditions)
							{
								Write-Verbose "$(Get-Date): `t`t`tCreate table for Load Balancing Policy Access Session Condition"
								$TableRange = $doc.Application.Selection.Range
								[int]$Columns = 2
								[int]$Rows = $LoadBalancingPolicyFilter.AccessSessionConditions.count + 1
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
								ForEach($AccessSessionCondition in $LoadBalancingPolicyFilter.AccessSessionConditions)
								{
									[string]$Tmp = $AccessSessionCondition
									[string]$AGFarm = $Tmp.substring(0, $Tmp.indexof(":"))
									[string]$AGFilter = $Tmp.substring($Tmp.indexof(":")+1)
									$xRow++
									Write-Verbose "$(Get-Date): `t`t`t`t`tProcessing row for Access Session Condition $($Tmp)"
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
						}
					}
				}
			
				If($LoadBalancingPolicyFilter.ClientIPAddressEnabled)
				{
					WriteWordLine 0 1 "Filter based on client IP address"
					If($LoadBalancingPolicyFilter.ApplyToAllClientIPAddresses)
					{
						WriteWordLine 0 2 "Apply to all client IP addresses"
					} 
					Else
					{
						If($LoadBalancingPolicyFilter.AllowedIPAddresses)
						{
							ForEach($AllowedIPAddress in $LoadBalancingPolicyFilter.AllowedIPAddresses)
							{
								WriteWordLine 0 2 "Client IP Address Matched: " $AllowedIPAddress
							}
						}
						If($LoadBalancingPolicyFilter.DeniedIPAddresses)
						{
							ForEach($DeniedIPAddress in $LoadBalancingPolicyFilter.DeniedIPAddresses)
							{
								WriteWordLine 0 2 "Client IP Address Ignored: " $DeniedIPAddress
							}
						}
					}
				}
				If($LoadBalancingPolicyFilter.ClientNameEnabled)
				{
					WriteWordLine 0 1 "Filter based on client name"
					If($LoadBalancingPolicyFilter.ApplyToAllClientNames)
					{
						WriteWordLine 0 2 "Apply to all client names"
					} 
					Else
					{
						If($LoadBalancingPolicyFilter.AllowedClientNames)
						{
							ForEach($AllowedClientName in $LoadBalancingPolicyFilter.AllowedClientNames)
							{
								WriteWordLine 0 2 "Client Name Matched: " $AllowedClientName
							}
						}
						If($LoadBalancingPolicyFilter.DeniedClientNames)
						{
							ForEach($DeniedClientName in $LoadBalancingPolicyFilter.DeniedClientNames)
							{
								WriteWordLine 0 2 "Client Name Ignored: " $DeniedClientName
							}
						}
					}
				}
				If($LoadBalancingPolicyFilter.AccountEnabled)
				{
					WriteWordLine 0 1 "Filter based on user"
					WriteWordLine 0 2 "Apply to anonymous users: " -nonewline
					If($LoadBalancingPolicyFilter.ApplyToAnonymousAccounts)
					{
						WriteWordLine 0 0 "Yes"
					}
					Else
					{
						WriteWordLine 0 0 "No"
					}
					If($LoadBalancingPolicyFilter.ApplyToAllExplicitAccounts)
					{
						WriteWordLine 0 2 "Apply to all explicit (non-anonymous) users"
					} 
					Else
					{
						If($LoadBalancingPolicyFilter.AllowedAccounts)
						{
							ForEach($AllowedAccount in $LoadBalancingPolicyFilter.AllowedAccounts)
							{
								WriteWordLine 0 2 "User Matched: " $AllowedAccount
							}
						}
						If($LoadBalancingPolicyFilter.DeniedAccounts)
						{
							ForEach($DeniedAccount in $LoadBalancingPolicyFilter.DeniedAccounts)
							{
								WriteWordLine 0 2 "User Ignored: " $DeniedAccount
							}
						}
					}
				}
				If($LoadBalancingPolicyConfiguration.WorkerGroupPreferenceAndFailoverState)
				{
					WriteWordLine 0 1 "Configure application connection preference based on worker group"
					If($LoadBalancingPolicyConfiguration.WorkerGroupPreferences)
					{
						Write-Verbose "$(Get-Date): `t`t`tCreate table for Load Balancing Policy Worker Group Filter"
						$TableRange = $doc.Application.Selection.Range
						[int]$Columns = 2
						[int]$Rows = $LoadBalancingPolicyConfiguration.WorkerGroupPreferences.count + 1
						$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
						$Table.rows.first.headingformat = $wdHeadingFormatTrue
						$Table.Style = $myHash.Word_TableGrid
						$Table.Borders.InsideLineStyle = $wdLineStyleSingle
						$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
						[int]$xRow = 1
						Write-Verbose "$(Get-Date): `t`t`t`tFormat first row with column headings"
						$Table.Rows.First.Shading.BackgroundPatternColor = $wdColorGray15
						$Table.Cell($xRow,1).Range.Font.Bold = $True
						$Table.Cell($xRow,1).Range.Text = "Worker Group"
						$Table.Cell($xRow,2).Range.Font.Bold = $True
						$Table.Cell($xRow,2).Range.Text = "Priority"
						ForEach($WorkerGroupPreference in $LoadBalancingPolicyConfiguration.WorkerGroupPreferences)
						{
							[string]$Tmp = $WorkerGroupPreference
							[string]$WGName = $Tmp.substring($Tmp.indexof("=")+1)
							[string]$WGPriority = $Tmp.substring($Tmp.indexof(":")+1, (($Tmp.indexof("=")-1)-$Tmp.indexof(":")))
							$xRow++
							Write-Verbose "$(Get-Date): `t`t`tProcessing row for Worker Group Filter $($Tmp)"
							$Table.Cell($xRow,1).Range.Text = $WGName
							$Table.Cell($xRow,2).Range.Text = $WGPriority
						}

						$Table.Rows.SetLeftIndent($Indent2TabStops,$wdAdjustProportional)
						$Table.AutoFitBehavior($wdAutoFitContent)

						FindWordDocumentEnd
						$tmp = $Null
						$WGName = $Null
						$WGPriority = $Null
					}
				}
				If($LoadBalancingPolicyConfiguration.StreamingDeliveryProtocolState -eq "Enabled")
				{
					WriteWordLine 0 1 "Set the delivery protocols for applications streamed to client"
					WriteWordLine 0 2 "" -nonewline
					Switch ($LoadBalancingPolicyConfiguration.StreamingDeliveryOption)
					{
						"Unknown"                {WriteWordLine 0 0 "Unknown"}
						"ForceServerAccess"      {WriteWordLine 0 0 "Do not allow applications to stream to the client"}
						"ForcedStreamedDelivery" {WriteWordLine 0 0 "Force applications to stream to the client"}
						Default {WriteWordLine 0 0 "Delivery protocol could not be determined: $($LoadBalancingPolicyConfiguration.StreamingDeliveryOption)"}
					}
				}
				ElseIf($LoadBalancingPolicyConfiguration.StreamingDeliveryProtocolState -eq "Disabled")
				{
					#In the GUI, if "Set the delivery protocols for applications streamed to client" IS selected AND 
					#"Allow applications to stream to the client or run on a Terminal Server (Default)" IS selected
					#then "Set the delivery protocols for applications streamed to client" is set to Disabled
					WriteWordLine 0 1 "Set the delivery protocols for applications streamed to client"
					WriteWordLine 0 2 "Allow applications to stream to the client or run on a Terminal Server (Default)"
				}
				Else
				{
					WriteWordLine 0 1 "Streamed App Delivery is not configured"
				}
			
				$LoadBalancingPolicyConfiguration = $Null
				$LoadBalancingPolicyFilter = $Null
			}
			Else
			{
				WriteWordLine 0 0 $LoadBalancingPolicy.PolicyName
			}
		}
	}
	ElseIf($LoadBalancingPolicies -eq $Null)
	{
		Write-Verbose "$(Get-Date): There are no Load balancing policies created"
	}
	Else 
	{
		Write-Warning "No results returned for Load balancing policy information"
	}
	$LoadBalancingPolicies = $Null
	Write-Verbose "$(Get-Date): Finished Processing Load Balancing Policies"
	Write-Verbose "$(Get-Date): "
}

If($Section -eq "All" -or $Section -eq "LoadEvals")
{
	#load evaluators
	Write-Verbose "$(Get-Date): Processing Load Evaluators"
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
				If(![String]::IsNullOrEmpty($LoadEvaluator.Description))
				{
					WriteWordLine 0 1 "Description: " $LoadEvaluator.Description
				}
				
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
					WriteWordLine 0 2 "Report full load when the # of users for this application equals: " $LoadEvaluator.ApplicationUserLoad
					WriteWordLine 0 2 "Application: " $LoadEvaluator.ApplicationBrowserName
				}
			
				If($LoadEvaluator.ContextSwitchesEnabled)
				{
					WriteWordLine 0 1 "Context Switches Settings"
					WriteWordLine 0 2 "Report full load when the # of context Switches per second is > than: " $LoadEvaluator.ContextSwitches[1]
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
					WriteWordLine 0 2 "Report full load when the total disk I/O in kbps is > than: " $LoadEvaluator.DiskDataIO[1]
					WriteWordLine 0 2 "Report no load when the total disk I/O in kbps per second is <= to: " $LoadEvaluator.DiskDataIO[0]
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
					WriteWordLine 0 0 "client connections from the listed IP Ranges"
					ForEach($IPRange in $LoadEvaluator.IPRanges)
					{
						WriteWordLine 0 3 "IP Address Ranges: " $IPRange
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

If($Section -eq "All" -or $Section -eq "Servers")
{
	#servers
	Write-Verbose "$(Get-Date): Processing Servers"
	Write-Verbose "$(Get-Date): `tSetting summary variables"
	[int]$TotalControllers = 0
	[int]$TotalWorkers = 0
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
						Write-Verbose "$(Get-Date): `t`t`t`t$($server.ServerName) is online."
						Write-Verbose "$(Get-Date): `t`t`t`tHardware and Software Inventory, Citrix Services and Hotfix areas will be processed."
					}
					ElseIf($Hardware -and !($Software))
					{
						Write-Verbose "$(Get-Date): `t`t`t`t$($server.ServerName) is online."
						Write-Verbose "$(Get-Date): `t`t`t`tHardware inventory, Citrix Services and Hotfix areas will be processed."
					}
					ElseIf(!($Hardware) -and $Software)
					{
						Write-Verbose "$(Get-Date): `t`t`t`t$($server.ServerName) is online."
						Write-Verbose "$(Get-Date): `t`t`t`tSoftware Inventory, Citrix Services and Hotfix areas will be processed."
					}
					Else
					{
						Write-Verbose "$(Get-Date): `t`t`t`t$($server.ServerName) is online."
						Write-Verbose "$(Get-Date): `t`t`t`tCitrix Services and Hotfix areas will be processed."
					}
				}
				
				#create array for appendix B
				Write-Verbose "$(Get-Date): `t`t`tGather server info for Appendix B"
				$obj = New-Object -TypeName PSObject
				$obj | Add-Member -MemberType NoteProperty -Name ServerName -Value $server.ServerName
				$obj | Add-Member -MemberType NoteProperty -Name ZoneName -Value $server.ZoneName
				$obj | Add-Member -MemberType NoteProperty -Name OSVersion -Value $server.OSVersion
				$obj | Add-Member -MemberType NoteProperty -Name CitrixVersion -Value $server.CitrixVersion
				$obj | Add-Member -MemberType NoteProperty -Name ProductEdition -Value $server.CitrixEdition
				$obj | Add-Member -MemberType NoteProperty -Name LicenseServer -Value $Server.LicenseServerName			

				$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $server.ServerName)
				$RegKey= $Reg.OpenSubKey("SYSTEM\\CurrentControlSet\\Control\\Citrix\\Wfshell\\TWI")
				$SSDisabled = $RegKey.GetValue("SeamlessFlags")
				
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

				WriteWordLine 2 0 $server.ServerName
				WriteWordLine 0 1 "Product`t`t`t`t: " $server.CitrixProductName
				WriteWordLine 0 1 "Edition`t`t`t`t: " $server.CitrixEdition
				WriteWordLine 0 1 "Version`t`t`t`t: " $server.CitrixVersion
				WriteWordLine 0 1 "Service Pack`t`t`t: " $server.CitrixServicePack
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
				WriteWordLine 0 0 " " $server.OSServicePack
				WriteWordLine 0 1 "Zone`t`t`t`t: " $server.ZoneName
				WriteWordLine 0 1 "Election Preference`t`t: " -nonewline
				Switch ($server.ElectionPreference)
				{
					"Unknown"           {WriteWordLine 0 0 "Unknown"}
					"MostPreferred"     {WriteWordLine 0 0 "Most Preferred"; $TotalControllers++}
					"Preferred"         {WriteWordLine 0 0 "Preferred"; $TotalControllers++}
					"DefaultPreference" {WriteWordLine 0 0 "Default Preference"; $TotalControllers++}
					"NotPreferred"      {WriteWordLine 0 0 "Not Preferred"; $TotalControllers++}
					"WorkerMode"        {WriteWordLine 0 0 "Worker Mode"; $TotalWorkers++}
					Default {WriteWordLine 0 0 "Server election preference could not be determined: $($server.ElectionPreference)"}
				}
				WriteWordLine 0 1 "Folder`t`t`t`t: " $server.FolderPath
				WriteWordLine 0 1 "Product Installation Path`t: " $server.CitrixInstallPath
				If($server.LicenseServerName)
				{
					WriteWordLine 0 1 "License Server Name`t`t: " $server.LicenseServerName
					WriteWordLine 0 1 "License Server Port`t`t: " $server.LicenseServerPortNumber
				}
				If($server.ICAPortNumber -gt 0)
				{
					WriteWordLine 0 1 "ICA Port Number`t`t: " $server.ICAPortNumber
				}
				
				WriteWordLine 0 0 ""

				If($SvrOnline -and $Hardware)
				{
					GetComputerWMIInfo $server.ServerName
				}

				#applications published to server
				$Applications = Get-XAApplication -ServerName $server.ServerName -EA 0 | Sort FolderPath, DisplayName
				If($? -and $Applications -ne $Null)
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
						foreach ($key in $subkeys1) 
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

						foreach ($key in $subkeys2) 
						{
							$thisKey=$UninstallKey2+"\\"+$key 
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
					$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)
					$Table.AutoFitBehavior($wdAutoFitContent)

					FindWordDocumentEnd
					WriteWordLine 0 0 ""
				}
				
				#list citrix services
				If($SvrOnline)
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
						ElseIf($Text)
						{
							Line 0 " ($NumServices Services found)"
						}
						ElseIf($HTML)
						{
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
							ElseIf($Text)
							{
								Line 0 "Display Name`t: " $Service.DisplayName
								Line 0 "Status`t`t: " $Service.State
								Line 0 "Start Mode`t: " $Service.StartMode
								Line 0 ""
							}
							ElseIf($HTML)
							{
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
						ElseIf($Text)
						{
						}
						ElseIf($HTML)
						{
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
						ElseIf($Text)
						{
							Line 0 "Warning: No Services were retrieved"
							Line 1 "If this is a trusted Forest, you may need to rerun the"
							Line 1 "script with Admin credentials from the trusted Forest."
						}
						ElseIf($HTML)
						{
						}
					}
					Else
					{
						Write-Warning "Services retrieval was successful but no services were returned."
						WriteWordLine 0 0 "Services retrieval was successful but no services were returned." "" $Null 0 $False $True
					}

					#Citrix hotfixes installed
					Write-Verbose "$(Get-Date): `t`tGet list of Citrix hotfixes installed"
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
						[bool]$HRP2Installed = $False
						
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
						ElseIf($Text)
						{
						}
						ElseIf($HTML)
						{
						}

						ForEach($hotfix in $hotfixes)
						{
							$HotfixArray += $hotfix.HotfixName
							If($hotfix.HotfixName -eq "XA600W2K8R2X64R02")
							{
								$HRP2Installed = $True
							}
							$InstallDate = $hotfix.InstalledOn.ToString()
							
							If($MSWord -or $PDF)
							{
								## Add the required key/values to the hashtable
								$WordTableRowHash = @{ HotfixName = $hotfix.HotfixName; InstalledBy = $hotfix.InstalledBy; InstallDate = $InstallDate.SubString(0,$InstallDate.IndexOf(" ")); HotfixType = $hotfix.HotfixType}

								## Add the hash to the array
								$HotfixesWordTable += $WordTableRowHash;

								$CurrentServiceIndex++;
							}
							ElseIf($Text)
							{
								Line 1 "Hotfix: " $hotfix.HotfixName
								Line 1 "Installed By: " $hotfix.InstalledBy
								Line 1 "Install Date: " $InstallDate.SubString(0,$InstallDate.IndexOf(" "))
								Line 1 "Type: " $hotfix.HotfixType
								Line 0 ""
							}
							ElseIf($HTML)
							{
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
						ElseIf($Text)
						{
						}
						ElseIf($HTML)
						{
						}

						#compare Citrix hotfixes to recommended Citrix hotfixes from CTX129229
						#hotfix lists are from CTX129229 dated 18-DEC-2014
						Write-Verbose "$(Get-Date): `t`tCompare Citrix hotfixes to recommended Citrix hotfixes from CTX129229"
						If(!$HRP2Installed)
						{
							Write-Verbose "$(Get-Date): `t`tProcessing pre HRP02 hotfix list for server $($server.ServerName)"
							$RecommendedList = @("XA600W2K8R2X64R01","XA600W2K8R2X64012","XA600W2K8R2X64017","XA600W2K8R2X64021",
										"XA600W2K8R2X64029", "XA600W2K8R2X64046", "XA600W2K8R2X64058", "XA600W2K8R2X64060",
										"XA600W2K8R2X64062", "XA600W2K8R2X64063", "XA600W2K8R2X64068", "XA600W2K8R2X64077", 
										"XA600W2K8R2X64079", "XA600W2K8R2X64089")
						}
						Else #HRP2 installed
						{
							Write-Verbose "$(Get-Date): `t`tProcessing HRP02 hotfix list for server $($server.ServerName)"
							$RecommendedList = @("XA600R02W2K8R2X64015", "XA600R02W2K8R2X64051")
						}
						
						If($RecommendedList.count -gt 0)
						{
							Write-Verbose "$(Get-Date): `t`tCreate Word Table for Citrix Hotfixes"
							If($MSWord -or $PDF)
							{
								WriteWordLine 0 1 "Citrix Recommended Hotfixes:"
								## Create an array of hashtables to store our hotfixes
								[System.Collections.Hashtable[]] $HotfixesWordTable = @();
								## Create an array of hashtables to store references of cells that we wish to highlight after the table has been added
								[System.Collections.Hashtable[]] $HighlightedCells = @();
								## Seed the row index from the second row
								[int] $CurrentServiceIndex = 2;
							}
							ElseIf($Text)
							{
								Line 1 "Citrix Recommended Hotfixes:"
							}
							ElseIf($HTML)
							{
							}
							
							ForEach($element in $RecommendedList)
							{
								$Tmp = $Null
								If(!($HotfixArray -contains $element))
								{
									#missing a recommended Citrix hotfix
									$Tmp = "Not Installed"
								}
								Else
								{
									$Tmp = "Installed"
								}
								If($MSWord -or $PDF)
								{
									## Add the required key/values to the hashtable
									$WordTableRowHash = @{ CitrixHotfix = $element; Status = $Tmp}

									## Add the hash to the array
									$HotfixesWordTable += $WordTableRowHash;

									If($Tmp -eq "Not Installed")
									{
										## Store "to highlight" cell references
										$HighlightedCells += @{ Row = $CurrentServiceIndex; Column = 2; }
									}
									$CurrentServiceIndex++;
								}
								ElseIf($Text)
								{
									Line 0 "Citrix Hotfix: " $element
									Line 0 "Status: " $Tmp
									Line 0 ""
								}
								ElseIf($HTML)
								{
								}
							}
							
							If($MSWord -or $PDF)
							{
								## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
								$Table = AddWordTable -Hashtable $HotfixesWordTable `
								-Columns CitrixHotfix, Status `
								-Headers "Citrix Hotfix", "Status" `
								-AutoFit $wdAutoFitContent;

								## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
								SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
								## IB - Set the required highlighted cells
								SetWordCellFormat -Coordinates $HighlightedCells -Table $Table -Bold -BackgroundColor $wdColorRed -Solid;

								#indent the entire table 1 tab stop
								$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)

								FindWordDocumentEnd
								$Table = $Null
								WriteWordLine 0 0 ""
							}
							ElseIf($Text)
							{
							}
							ElseIf($HTML)
							{
							}
						}
						#build list of installed Microsoft hotfixes
						Write-Verbose "$(Get-Date): `t`tProcessing Microsoft hotfixes for server $($server.ServerName)"
						[bool]$GotMSHotfixes = $True
						
						Try
						{
							$results = Get-HotFix -computername $Server.ServerName 
							$MSInstalledHotfixes = $results | select-object -Expand HotFixID | Sort HotFixID
							$results = $Null
						}
						
						Catch
						{
							$GotMSHotfixes = $False
						}
						
						If($GotMSHotfixes)
						{
							If($server.OSServicePack.IndexOf('1') -gt 0)
							{
								#Server 2008 R2 SP1 installed
								$RecommendedList = @("KB2620656", "KB2647753", "KB2728738", "KB2748302", 
												"KB2775511", "KB2778831", "KB2896256", "KB2908190", 
												"KB2920289", "KB917607")
							}
							Else
							{
								#Server 2008 R2 without SP1 installed
								$RecommendedList = @("KB2265716", "KB2383928", "KB2647753", "KB2728738", 
												"KB2748302", "KB2775511", "KB2778831", "KB2896256", 
												"KB3014783", "KB917607", "KB975777", "KB979530", 
												"KB980663", "KB983460")
							}
							
							If($RecommendedList.count -gt 0)
							{
								Write-Verbose "$(Get-Date): `t`tCreate Word Table for Microsoft Hotfixes"
								Write-Verbose "$(Get-Date): `t`tAdd Microsoft hotfix table to doc"
								If($MSWord -or $PDF)
								{
									WriteWordLine 0 1 "Microsoft Recommended Hotfixes (from CTX129229):"
									## Create an array of hashtables to store our hotfixes
									[System.Collections.Hashtable[]] $HotfixesWordTable = @();
									## Create an array of hashtables to store references of cells that we wish to highlight after the table has been added
									[System.Collections.Hashtable[]] $HighlightedCells = @();
									## Seed the row index from the second row
									[int] $CurrentServiceIndex = 2;
								}
								ElseIf($Text)
								{
									Line 1 "Microsoft Recommended Hotfixes (from CTX129229):"
								}
								ElseIf($HTML)
								{
								}

								ForEach($hotfix in $RecommendedList)
								{
									$Tmp = $Null
									If(!($MSInstalledHotfixes -contains $hotfix))
									{
										$Tmp = "Not Installed"
									}
									Else
									{
										$Tmp = "Installed"
									}
									If($MSWord -or $PDF)
									{
										## Add the required key/values to the hashtable
										$WordTableRowHash = @{ MicrosoftHotfix = $hotfix; Status = $Tmp}

										## Add the hash to the array
										$HotfixesWordTable += $WordTableRowHash;

										If($Tmp -eq "Not Installed")
										{
											## Store "to highlight" cell references
											$HighlightedCells += @{ Row = $CurrentServiceIndex; Column = 2; }
										}
										$CurrentServiceIndex++;
									}
									ElseIf($Text)
									{
										Line 0 "Microsoft Hotfix: " $hotfix
										Line 0 "Status: " $Tmp
										Line 0 ""
									}
									ElseIf($HTML)
									{
									}
								}
								
								If($MSWord -or $PDF)
								{
									## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
									$Table = AddWordTable -Hashtable $HotfixesWordTable `
									-Columns MicrosoftHotfix, Status `
									-Headers "Microsoft Hotfix", "Status" `
									-AutoFit $wdAutoFitContent;

									## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
									SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
									## IB - Set the required highlighted cells
									SetWordCellFormat -Coordinates $HighlightedCells -Table $Table -Bold -BackgroundColor $wdColorRed -Solid;

									#indent the entire table 1 tab stop
									$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)

									FindWordDocumentEnd
									$Table = $Null
									WriteWordLine 0 1 "Not all missing Microsoft hotfixes may be needed for this server `n`tor might already be replaced and not recorded in CTX129229."
									WriteWordLine 0 0 ""
								}
								ElseIf($Text)
								{
								}
								ElseIf($HTML)
								{
								}
							}
						}
						Else
						{
							Write-Verbose "$(Get-Date): Get-HotFix failed for $($server.ServerName)"
							Write-Warning "Get-HotFix failed for $($server.ServerName)"
							If($MSWord -or $PDF)
							{
								WriteWordLine 0 0 "Get-HotFix failed for $($server.ServerName)" "" $Null 0 $False $True
								WriteWordLine 0 0 "On $($server.ServerName) you may need to run winmgmt /verifyrepository and winmgmt /salvagerepository"
							}
							ElseIf($Text)
							{
								Line 0 "Get-HotFix failed for $($server.ServerName)"
								Line 0 "On $($server.ServerName) you may need to run winmgmt /verifyrepository and winmgmt /salvagerepository"
							}
							ElseIf($HTML)
							{
							}
						}
					}
					ElseIf(!$?)
					{
						Write-Warning "No Citrix hotfixes were retrieved"
						If($MSWORD -or $PDF)
						{
							WriteWordLine 0 0 "Warning: No Citrix hotfixes were retrieved" "" $Null 0 $False $True
						}
						ElseIf($Text)
						{
							Line 0 "Warning: No Citrix hotfixes were retrieved"
						}
						ElseIf($HTML)
						{
						}
					}
					Else
					{
						Write-Warning "Citrix hotfix retrieval was successful but no Citrix hotfixes were returned."
						If($MSWORD -or $PDF)
						{
							WriteWordLine 0 0 "Citrix hotfix retrieval was successful but no Citrix hotfixes were returned." "" $Null 0 $False $True
						}
						ElseIf($Text)
						{
							Line 0 "Citrix hotfix retrieval was successful but no Citrix hotfixes were returned."
						}
						ElseIf($HTML)
						{
						}
					}
				}
				Else
				{
					Write-Verbose "$(Get-Date): `t`t$($server.ServerName) is offline or unreachable.  Citrix Services and Hotfix areas skipped."
					WriteWordLine 0 0 "Server $($server.ServerName) was offline or unreachable at "(Get-date).ToString()
					WriteWordLine 0 0 "The Citrix Services and Hotfix areas were skipped."
				}
				WriteWordLine 0 0 "" 
				Write-Verbose "$(Get-Date): `tFinished Processing server $($server.ServerName)"
				Write-Verbose "$(Get-Date): "
			}
			Else
			{
				WriteWordLine 0 0 $server.ServerName
				$TotalServers++
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

If($Section -eq "All" -or $Section -eq "WGs")
{
	#worker groups
	Write-Verbose "$(Get-Date): Processing Worker Groups"
	Write-Verbose "$(Get-Date): `tSetting summary variables"
	[int]$TotalWGByServerName = 0
	[int]$TotalWGByServerGroup = 0
	[int]$TotalWGByOU = 0
	[int]$TotalWGs = 0

	Write-Verbose "$(Get-Date): `tRetrieving Worker Groups"
	$WorkerGroups = Get-XAWorkerGroup -EA 0 | Sort WorkerGroupName

	If($? -and $WorkerGroups -ne $Null)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Worker Groups:"
		ForEach($WorkerGroup in $WorkerGroups)
		{
			Write-Verbose "$(Get-Date): `tProcessing worker group $($WorkerGroup.WorkerGroupName)"
			If(!$Summary)
			{
				WriteWordLine 2 0 $WorkerGroup.WorkerGroupName
				If(![String]::IsNullOrEmpty($WorkerGroup.Description))
				{
					WriteWordLine 0 1 "Description: " $WorkerGroup.Description
				}
				WriteWordLine 0 1 "Folder Path: " $WorkerGroup.FolderPath
				If($WorkerGroup.ServerNames)
				{
					$TotalWGByServerName++
					WriteWordLine 0 1 "Farm Servers:"
					Write-Verbose "$(Get-Date): `t`tProcessing Worker Group by Farm Servers"
					Write-Verbose "$(Get-Date): `t`tCreate Word Table for Worker Group by Farm Server"
					$TempArray = $WorkerGroup.ServerNames | Sort
					BuildTableForServerOrWG $TempArray "Server"
					$TempArray = $Null
				}
				If($WorkerGroup.ServerGroups)
				{
					$TotalWGByServerGroup++
					WriteWordLine 0 1 "Server Group Accounts:"
					Write-Verbose "$(Get-Date): `t`tProcessing Worker Group by Server Groups"
					Write-Verbose "$(Get-Date): `t`tCreate Word Table for Worker Group by Server Groups"
					$TempArray = $WorkerGroup.ServerGroups | Sort
					BuildTableForServerOrWG $TempArray "Security Group"
					$TempArray = $Null
				}
				If($WorkerGroup.OUs)
				{
					$TotalWGByOU++
					WriteWordLine 0 1 "Organizational Units:"
					Write-Verbose "$(Get-Date): `t`tProcessing Worker Group by OUs"
					Write-Verbose "$(Get-Date): `t`tCreate Word Table for Worker Group by OUs"
					$TempArray = $WorkerGroup.OUs | Sort {$_.Length}
					BuildTableForServerOrWG $TempArray "OU"
					$TempArray = $Null
				}
				#applications published to worker group
				$Applications = Get-XAApplication -WorkerGroup $WorkerGroup.WorkerGroupName -EA 0 | Sort FolderPath, DisplayName
				If($? -and $Applications -ne $Null)
				{
					WriteWordLine 0 0 ""
					WriteWordLine 0 1 "Published applications:"
					Write-Verbose "$(Get-Date): `t`tProcessing published applications for Worker Group $($WorkerGroup.WorkerGroupName)"
					Write-Verbose "$(Get-Date): `t`tCreate Word Table for Worker Group's published applications"
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
					$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)
					$Table.AutoFitBehavior($wdAutoFitContent)

					FindWordDocumentEnd
				}
				WriteWordLine 0 0 ""
			}
			Else
			{
				WriteWordLine 0 0 $WorkerGroup.WorkerGroupName
				$TotalWGs++
			}
		}
	}
	ElseIf($WorkerGroups -eq $Null)
	{

		Write-Verbose "$(Get-Date): There are no Worker Groups created"
	}
	Else 
	{
		Write-Warning "No results returned for Worker Group information"
	}
	$WorkerGroups = $Null
	Write-Verbose "$(Get-Date): Finished Processing Worker Groups"
	Write-Verbose "$(Get-Date): "
}

If($Section -eq "All" -or $Section -eq "Zones")
{
	#zones
	Write-Verbose "$(Get-Date): Processing Zones"
	Write-Verbose "$(Get-Date): `tSetting summary variables"
	[int]$TotalZones = 0

	Write-Verbose "$(Get-Date): `tRetrieving Zones"
	$Zones = Get-XAZone -EA 0 | Sort ZoneName
	If($? -and $Zones -ne $Null)
	{
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
	$Zones = $Null
	Write-Verbose "$(Get-Date): Finished Processing Zones"
	Write-Verbose "$(Get-Date): "
}

If($Section -eq "All" -or $Section -eq "Policies")
{
	[int]$Global:TotalComputerPolicies = 0
	[int]$Global:TotalUserPolicies = 0
	[int]$Global:TotalIMAPolicies = 0
	[int]$Global:TotalADPolicies = 0
	[int]$Global:TotalADPoliciesNotProcessed = 0
	[int]$Global:TotalPolicies = 0
	$ADPoliciesNotProcessed = @()

	#make sure Citrix.GroupPolicy.Commands module is loaded
	If(!(Check-LoadedModule "Citrix.GroupPolicy.Commands"))
	{
		Write-Warning "The Citrix Group Policy module Citrix.GroupPolicy.Commands.psm1 could not be loaded `nPlease see the Prerequisites section in the ReadMe file (https://www.dropbox.com/s/glq4u2p5xte8s6g/XA6_Inventory_V42_ReadMe.rtf). `nCitrix Policy documentation will not take place"
		Write-Verbose "$(Get-Date): "
	}
	Else
	{

		$selection.InsertNewPage()
		WriteWordLine 1 0 "Policies:"
		Write-Verbose "$(Get-Date): Processing Citrix IMA Policies"
		Write-Verbose "$(Get-Date): `tRetrieving IMA Farm Policies"
		ProcessCitrixPolicies	
		Write-Verbose "$(Get-Date): Finished Processing Citrix IMA Policies"
		Write-Verbose "$(Get-Date): "
		
		#thanks to the Citrix Engineering Team for helping me solve processing Citrix AD based Policies
		Write-Verbose "$(Get-Date): See if there are any Citrix AD based policies to process"
		$CtxGPOArray = @()
		$CtxGPOArray = GetCtxGPOsInAD
		If($CtxGPOArray -is [Array] -and $CtxGPOArray.Count -gt 0)
		{
			Write-Verbose "$(Get-Date): There are $($CtxGPOArray.Count) Citrix AD based policies to process"

			$CtxGPOArray = $CtxGPOArray | Sort
			
			ForEach($CtxGPO in $CtxGPOArray)
			{
				Write-Verbose "$(Get-Date): Creating ADGpoDrv PSDrive"
				New-PSDrive -Name ADGpoDrv -PSProvider CitrixGroupPolicy -Root \ -DomainGpo $($CtxGPO) -Scope "Global" -EA 0 4>$Null
				If(Get-PSDrive ADGpoDrv -EA 0)
				{
					Write-Verbose "$(Get-Date): Processing Citrix AD Policy $($CtxGPO)"
				
					Write-Verbose "$(Get-Date): `tRetrieving AD Policy $($CtxGPO)"
					ProcessCitrixPolicies "ADGpoDrv"
					Write-Verbose "$(Get-Date): Finished Processing Citrix AD Policy $($CtxGPO)"
					Write-Verbose "$(Get-Date): "
				}
				Else
				{
					$ADPoliciesNotProcessed += $CtxGPO
					$Global:TotalADPoliciesNotProcessed++
					Write-Warning "$($CtxGPO) is not readable by this XenApp 6.0 server"
					Write-Warning "$($CtxGPO)  was probably created by an updated Citrix Group Policy Provider"
				}
			}
			
			If(!$Summary)
			{
				If($Global:TotalADPoliciesNotProcessed -gt 0)
				{
					Write-Verbose "$(Get-Date): Processing list of Citrix AD Policies not processed"
					$ADPoliciesNotProcessed = $ADPoliciesNotProcessed | Sort -unique
					WriteWordLine 0 0 ""
					WriteWordLine 2 0 "Active Directory Citrix policies that could not be processed:"
					ForEach($Policy in $ADPoliciesNotProcessed)
					{
						Write-Verbose "$(Get-Date): `t Processing skipped Citrix AD policy $($Policy)"
						WriteWordLine 0 1 $Policy
					}
					Write-Verbose "$(Get-Date): Finished processing list of Citrix AD Policies not processed"
					WriteWordLine 0 0 ""
				}
			}
			Write-Verbose "$(Get-Date): Finished Processing Citrix AD Policies"
			Write-Verbose "$(Get-Date): "
		}
		Else
		{
			Write-Verbose "$(Get-Date): There are no Citrix AD based policies to process"
			Write-Verbose "$(Get-Date): "
		}

		$Policies = $Null
		Write-Verbose "$(Get-Date): Finished Processing Citrix Policies"
		Write-Verbose "$(Get-Date): "
	}
}

If(!$Summary -and ($Section -eq "All"))
{
	#	The Session Sharing Key is generated by the XML Broker in XenApp 6.5.  
	#	Web Interface or StoreFront send the following information to the XML Broker:"
	#	Audio Quality (Policy Setting)"
	#	Client Printer Port Mapping (Policy Setting)"
	#	Client Printer Spooling (Policy Setting)"
	#	Color Depth (Application Setting)"
	#	COM Port Mapping (Policy Setting)"
	#	Display Size (Application Setting)"
	#	Domain Name (Logon)"
	#	EnableSessionSharing (ICA file or Client Registry Setting)"
	#	Encryption Level (Application Setting and Policy Setting.  Policy wins.)"
	#	Farm Name (Web Interface/StoreFront)"
	#	Special Folder Redirection (Policy Setting)"
	#	TWIDisableSessionSharing(ICA file or Client Registry Setting)"
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
		## Seed the $Services row index from the second row
		[int] $CurrentServiceIndex = 2;
	}
	ElseIf($Text)
	{
	}
	ElseIf($HTML)
	{
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
		ElseIf($Text)
		{
		}
		ElseIf($HTML)
		{
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
	ElseIf($Text)
	{
	}
	ElseIf($HTML)
	{
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
		## Seed the $Services row index from the second row
		[int] $CurrentServiceIndex = 2;
	}
	ElseIf($Text)
	{
	}
	ElseIf($HTML)
	{
	}

	ForEach($Item in $ServerItems)
	{
		If($MSWord -or $PDF)
		{
			## Add the required key/values to the hashtable
			$Tmp = $Null
			If([String]::IsNullOrEmpty($Item.LicenseServer))
			{
				$Tmp = "Set by policy"
			}
			Else
			{
				$Tmp = $Item.LicenseServer
			}
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
		ElseIf($Text)
		{
		}
		ElseIf($HTML)
		{
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
	ElseIf($Text)
	{
	}
	ElseIf($HTML)
	{
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
		Write-Verbose "$(Get-Date): `tAdd administrator summary info"
		WriteWordLine 0 0 "Administrators"
		WriteWordLine 0 1 "Total Full Administrators`t: " $TotalFullAdmins
		WriteWordLine 0 1 "Total View Administrators`t: " $TotalViewAdmins
		WriteWordLine 0 1 "Total Custom Administrators`t: " $TotalCustomAdmins
		WriteWordLine 0 2 "Total Administrators`t: " ($TotalFullAdmins + $TotalViewAdmins + $TotalCustomAdmins)
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): `tAdd application summary info"
		WriteWordLine 0 0 "Applications"
		WriteWordLine 0 1 "Total Published Applications`t: " $TotalPublishedApps
		WriteWordLine 0 1 "Total Published Content`t`t: " $TotalPublishedContent
		WriteWordLine 0 1 "Total Published Desktops`t: " $TotalPublishedDesktops
		WriteWordLine 0 1 "Total Streamed Applications`t: " $TotalStreamedApps
		WriteWordLine 0 2 "Total Applications`t: " ($TotalPublishedApps + $TotalPublishedContent + $TotalPublishedDesktops + $TotalStreamedApps)
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): `tAdd configuration logging summary info"
		WriteWordLine 0 0 "Configuration Logging"
		WriteWordLine 0 1 "Total Config Log Items`t`t: " $TotalConfigLogItems 
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): `tAdd load balancing policies summary info"
		WriteWordLine 0 0 "Load Balancing Policies"
		WriteWordLine 0 1 "Total Load Balancing Policies`t: " $TotalLBPolicies
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): `tAdd load evaluator summary info"
		WriteWordLine 0 0 "Load Evaluators"
		WriteWordLine 0 1 "Total Load Evaluators`t`t: " $TotalLoadEvaluators
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): `tAdd server summary info"
		WriteWordLine 0 0 "Servers"
		WriteWordLine 0 1 "Total Controllers`t`t: " $TotalControllers
		WriteWordLine 0 1 "Total Workers`t`t`t: " $TotalWorkers
		WriteWordLine 0 2 "Total Servers`t`t: " ($TotalControllers + $TotalWorkers)
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): `tAdd worker group summary info"
		WriteWordLine 0 0 "Worker Groups"
		WriteWordLine 0 1 "Total WGs by Server Name`t: " $TotalWGByServerName
		WriteWordLine 0 1 "Total WGs by Server Group`t: " $TotalWGByServerGroup
		WriteWordLine 0 1 "Total WGs by AD Container`t: " $TotalWGByOU
		WriteWordLine 0 2 "Total Worker Groups`t: " ($TotalWGByServerName + $TotalWGByServerGroup + $TotalWGByOU)
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): `tAdd zone summary info"
		WriteWordLine 0 0 "Zones"
		WriteWordLine 0 1 "Total Zones`t`t`t: " $TotalZones
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): `tAdd policy summary info"
		WriteWordLine 0 0 "Policies"
		WriteWordLine 0 1 "Total Computer Policies`t`t: " $Global:TotalComputerPolicies
		WriteWordLine 0 1 "Total User Policies`t`t: " $Global:TotalUserPolicies
		WriteWordLine 0 2 "Total Policies`t`t: " ($Global:TotalComputerPolicies + $Global:TotalUserPolicies)
		WriteWordLine 0 0 ""
		WriteWordLine 0 1 "IMA Policies`t`t`t: " $Global:TotalIMAPolicies
		WriteWordLine 0 1 "Citrix AD Policies Processed`t: $($Global:TotalADPolicies)`t(AD Policies can contain multiple Citrix policies)"
		WriteWordLine 0 1 "Citrix AD Policies not Processed`t: " $Global:TotalADPoliciesNotProcessed
	}
	Else
	{
		Write-Verbose "$(Get-Date): `tAdd administrator summary info"
		WriteWordLine 0 0 "Administrators"
		WriteWordLine 0 1 "Total Administrators`t`t: " $TotalAdmins
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): `tAdd application summary info"
		WriteWordLine 0 0 "Applications"
		WriteWordLine 0 1 "Total Applications`t`t: " $TotalApps
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): `tAdd load balancing policies summary info"
		WriteWordLine 0 0 "Load Balancing Policies"
		WriteWordLine 0 1 "Total Load Balancing Policies`t: " $TotalLBPolicies
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): `tAdd load evaluator summary info"
		WriteWordLine 0 0 "Load Evaluators"
		WriteWordLine 0 1 "Total Load Evaluators`t`t: " $TotalLoadEvaluators
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): `tAdd server summary info"
		WriteWordLine 0 0 "Servers"
		WriteWordLine 0 1 "Total Servers`t`t`t: " $TotalServers
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): `tAdd worker group summary info"
		WriteWordLine 0 0 "Worker Groups"
		WriteWordLine 0 1 "Total Worker Groups`t`t: " $TotalWGs
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): `tAdd zone summary info"
		WriteWordLine 0 0 "Zones"
		WriteWordLine 0 1 "Total Zones`t`t`t: " $TotalZones
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): `tAdd policy summary info"
		WriteWordLine 0 0 "Policies"
		WriteWordLine 0 1 "IMA Policies`t`t`t: " $Global:TotalIMAPolicies
		WriteWordLine 0 1 "Citrix AD Policies Processed`t: $($Global:TotalADPolicies)`t(AD Policies can contain multiple Citrix policies)"
		WriteWordLine 0 1 "Total Policies`t`t`t: " $Global:TotalPolicies
	}
}

Write-Verbose "$(Get-Date): Finishing up document"
#end of document processing

$AbstractTitle = "Citrix XenApp 6 Inventory"
$SubjectTitle = "XenApp 6 Farm Inventory"
UpdateDocumentProperties $AbstractTitle $SubjectTitle

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

# SIG # Begin signature block
# MIIgAAYJKoZIhvcNAQcCoIIf8TCCH+0CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU3J2j6LqdAJDzoME1EyHC3pfP
# otegghtnMIIDtzCCAp+gAwIBAgIQDOfg5RfYRv6P5WD8G/AwOTANBgkqhkiG9w0B
# AQUFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVk
# IElEIFJvb3QgQ0EwHhcNMDYxMTEwMDAwMDAwWhcNMzExMTEwMDAwMDAwWjBlMQsw
# CQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cu
# ZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVkIElEIFJvb3Qg
# Q0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCtDhXO5EOAXLGH87dg
# +XESpa7cJpSIqvTO9SA5KFhgDPiA2qkVlTJhPLWxKISKityfCgyDF3qPkKyK53lT
# XDGEKvYPmDI2dsze3Tyoou9q+yHyUmHfnyDXH+Kx2f4YZNISW1/5WBg1vEfNoTb5
# a3/UsDg+wRvDjDPZ2C8Y/igPs6eD1sNuRMBhNZYW/lmci3Zt1/GiSw0r/wty2p5g
# 0I6QNcZ4VYcgoc/lbQrISXwxmDNsIumH0DJaoroTghHtORedmTpyoeb6pNnVFzF1
# roV9Iq4/AUaG9ih5yLHa5FcXxH4cDrC0kqZWs72yl+2qp/C3xag/lRbQ/6GW6whf
# GHdPAgMBAAGjYzBhMA4GA1UdDwEB/wQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB0G
# A1UdDgQWBBRF66Kv9JLLgjEtUYunpyGd823IDzAfBgNVHSMEGDAWgBRF66Kv9JLL
# gjEtUYunpyGd823IDzANBgkqhkiG9w0BAQUFAAOCAQEAog683+Lt8ONyc3pklL/3
# cmbYMuRCdWKuh+vy1dneVrOfzM4UKLkNl2BcEkxY5NM9g0lFWJc1aRqoR+pWxnmr
# EthngYTffwk8lOa4JiwgvT2zKIn3X/8i4peEH+ll74fg38FnSbNd67IJKusm7Xi+
# fT8r87cmNW1fiQG2SVufAQWbqz0lwcy2f8Lxb4bG+mRo64EtlOtCt/qMHt1i8b5Q
# Z7dsvfPxH2sMNgcWfzd8qVttevESRmCD1ycEvkvOl77DZypoEd+A5wwzZr8TDRRu
# 838fYxAe+o0bJW1sj6W3YQGx0qMmoRBxna3iw/nDmVG3KwcIzi7mULKn+gpFL6Lw
# 8jCCBTAwggQYoAMCAQICEAQJGBtf1btmdVNDtW+VUAgwDQYJKoZIhvcNAQELBQAw
# ZTELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQ
# d3d3LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBS
# b290IENBMB4XDTEzMTAyMjEyMDAwMFoXDTI4MTAyMjEyMDAwMFowcjELMAkGA1UE
# BhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2lj
# ZXJ0LmNvbTExMC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIENvZGUg
# U2lnbmluZyBDQTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAPjTsxx/
# DhGvZ3cH0wsxSRnP0PtFmbE620T1f+Wondsy13Hqdp0FLreP+pJDwKX5idQ3Gde2
# qvCchqXYJawOeSg6funRZ9PG+yknx9N7I5TkkSOWkHeC+aGEI2YSVDNQdLEoJrsk
# acLCUvIUZ4qJRdQtoaPpiCwgla4cSocI3wz14k1gGL6qxLKucDFmM3E+rHCiq85/
# 6XzLkqHlOzEcz+ryCuRXu0q16XTmK/5sy350OTYNkO/ktU6kqepqCquE86xnTrXE
# 94zRICUj6whkPlKWwfIPEvTFjg/BougsUfdzvL2FsWKDc0GCB+Q4i2pzINAPZHM8
# np+mM6n9Gd8lk9ECAwEAAaOCAc0wggHJMBIGA1UdEwEB/wQIMAYBAf8CAQAwDgYD
# VR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHkGCCsGAQUFBwEBBG0w
# azAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEMGCCsGAQUF
# BzAChjdodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVk
# SURSb290Q0EuY3J0MIGBBgNVHR8EejB4MDqgOKA2hjRodHRwOi8vY3JsNC5kaWdp
# Y2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMDqgOKA2hjRodHRw
# Oi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3Js
# ME8GA1UdIARIMEYwOAYKYIZIAYb9bAACBDAqMCgGCCsGAQUFBwIBFhxodHRwczov
# L3d3dy5kaWdpY2VydC5jb20vQ1BTMAoGCGCGSAGG/WwDMB0GA1UdDgQWBBRaxLl7
# KgqjpepxA8Bg+S32ZXUOWDAfBgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823I
# DzANBgkqhkiG9w0BAQsFAAOCAQEAPuwNWiSz8yLRFcgsfCUpdqgdXRwtOhrE7zBh
# 134LYP3DPQ/Er4v97yrfIFU3sOH20ZJ1D1G0bqWOWuJeJIFOEKTuP3GOYw4TS63X
# X0R58zYUBor3nEZOXP+QsRsHDpEV+7qvtVHCjSSuJMbHJyqhKSgaOnEoAjwukaPA
# JRHinBRHoXpoaK+bp1wgXNlxsQyPu6j4xRJon89Ay0BEpRPw5mQMJQhCMrI2iiQC
# /i9yfhzXSUWW6Fkd6fp0ZGuy62ZD2rOwjNXpDd32ASDOmTFjPQgaGLOBm0/GkxAG
# /AeB+ova+YJJ92JuoVP6EpQYhS6SkepobEQysmah5xikmmRR7zCCBTUwggQdoAMC
# AQICEAS/j+KrDcLYfUmxXxwPKbAwDQYJKoZIhvcNAQELBQAwcjELMAkGA1UEBhMC
# VVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0
# LmNvbTExMC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIENvZGUgU2ln
# bmluZyBDQTAeFw0xNDEwMTQwMDAwMDBaFw0xNTEwMTkxMjAwMDBaMHwxCzAJBgNV
# BAYTAlVTMQswCQYDVQQIEwJUTjESMBAGA1UEBxMJVHVsbGFob21hMSUwIwYDVQQK
# ExxDYXJsIFdlYnN0ZXIgQ29uc3VsdGluZywgTExDMSUwIwYDVQQDExxDYXJsIFdl
# YnN0ZXIgQ29uc3VsdGluZywgTExDMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIB
# CgKCAQEAuRjU1TLpMCFI+3Yw2ypfw5isSm3Dd8UO/PdhT9wou8k9mVgF8PEB+7Y3
# pz+wMnAywReRSeEFYkSA9NHLFn7/KFU65NyJklwU4EFWnHCg8lEvjwIlDmKQJB6G
# DJ9SMPuOVUiFFQLK3sDfi9SJEaPO7QiIb48IJRCTtJdMcN1MXEh6J6nZt2dvvEqf
# +RrsCzIg8ETAiWN+ha3iWJ5shtRf4kQo+toZmIBejP+DSu7vbh+lunckySMfVSws
# JHKcnb3QHwXDz9oV8gjxjBTuLyx9lAyEeuMhFPSJF4v1WbAW+l51x6XQUUn1lFPV
# Ru8QA/U04s/UQO3P7Fhp8CjZ0CGNhQIDAQABo4IBuzCCAbcwHwYDVR0jBBgwFoAU
# WsS5eyoKo6XqcQPAYPkt9mV1DlgwHQYDVR0OBBYEFL3QsUmxf1g5OE1z/KRjZGBy
# //TrMA4GA1UdDwEB/wQEAwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzB3BgNVHR8E
# cDBuMDWgM6Axhi9odHRwOi8vY3JsMy5kaWdpY2VydC5jb20vc2hhMi1hc3N1cmVk
# LWNzLWcxLmNybDA1oDOgMYYvaHR0cDovL2NybDQuZGlnaWNlcnQuY29tL3NoYTIt
# YXNzdXJlZC1jcy1nMS5jcmwwQgYDVR0gBDswOTA3BglghkgBhv1sAwEwKjAoBggr
# BgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzCBhAYIKwYBBQUH
# AQEEeDB2MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wTgYI
# KwYBBQUHMAKGQmh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFNI
# QTJBc3N1cmVkSURDb2RlU2lnbmluZ0NBLmNydDAMBgNVHRMBAf8EAjAAMA0GCSqG
# SIb3DQEBCwUAA4IBAQCBqMZIc7qV9RkO55hJLX5kWnsu1dnpoqy1Ww2YnhwnQbiV
# JzytM8q/tXw4RF5Hstj9m2KepztzBM4NIyxryDKpaC34rQsp6zCF8Fq6a81rQpY+
# cVy4muKpVptzuaz4aYVi+6tnoK/KheMPu5g52M17HOkwxZUlcCpGpjKted3vbRoh
# 5K16sRzjcOcXJW/dw+6UGD14lHPCIShp/KNasB2XT3pvBOBg75heJmoeBVJA5ztB
# pHJ5XVAerIwd/Ycglqu1CteKN5D6wQBo95Eb74HTesVQtSsAD2UflpVt2Y4VJyk2
# mq4z16PhD1UABXmZujQe2JgrVcyPbpM3Ub+i+1wUMIIGajCCBVKgAwIBAgIQAwGa
# Ajr/WLFr1tXq5hfwZjANBgkqhkiG9w0BAQUFADBiMQswCQYDVQQGEwJVUzEVMBMG
# A1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEw
# HwYDVQQDExhEaWdpQ2VydCBBc3N1cmVkIElEIENBLTEwHhcNMTQxMDIyMDAwMDAw
# WhcNMjQxMDIyMDAwMDAwWjBHMQswCQYDVQQGEwJVUzERMA8GA1UEChMIRGlnaUNl
# cnQxJTAjBgNVBAMTHERpZ2lDZXJ0IFRpbWVzdGFtcCBSZXNwb25kZXIwggEiMA0G
# CSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCjZF38fLPggjXg4PbGKuZJdTvMbuBT
# qZ8fZFnmfGt/a4ydVfiS457VWmNbAklQ2YPOb2bu3cuF6V+l+dSHdIhEOxnJ5fWR
# n8YUOawk6qhLLJGJzF4o9GS2ULf1ErNzlgpno75hn67z/RJ4dQ6mWxT9RSOOhkRV
# fRiGBYxVh3lIRvfKDo2n3k5f4qi2LVkCYYhhchhoubh87ubnNC8xd4EwH7s2AY3v
# J+P3mvBMMWSN4+v6GYeofs/sjAw2W3rBerh4x8kGLkYQyI3oBGDbvHN0+k7Y/qpA
# 8bLOcEaD6dpAoVk62RUJV5lWMJPzyWHM0AjMa+xiQpGsAsDvpPCJEY93AgMBAAGj
# ggM1MIIDMTAOBgNVHQ8BAf8EBAMCB4AwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8E
# DDAKBggrBgEFBQcDCDCCAb8GA1UdIASCAbYwggGyMIIBoQYJYIZIAYb9bAcBMIIB
# kjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzCCAWQG
# CCsGAQUFBwICMIIBVh6CAVIAQQBuAHkAIAB1AHMAZQAgAG8AZgAgAHQAaABpAHMA
# IABDAGUAcgB0AGkAZgBpAGMAYQB0AGUAIABjAG8AbgBzAHQAaQB0AHUAdABlAHMA
# IABhAGMAYwBlAHAAdABhAG4AYwBlACAAbwBmACAAdABoAGUAIABEAGkAZwBpAEMA
# ZQByAHQAIABDAFAALwBDAFAAUwAgAGEAbgBkACAAdABoAGUAIABSAGUAbAB5AGkA
# bgBnACAAUABhAHIAdAB5ACAAQQBnAHIAZQBlAG0AZQBuAHQAIAB3AGgAaQBjAGgA
# IABsAGkAbQBpAHQAIABsAGkAYQBiAGkAbABpAHQAeQAgAGEAbgBkACAAYQByAGUA
# IABpAG4AYwBvAHIAcABvAHIAYQB0AGUAZAAgAGgAZQByAGUAaQBuACAAYgB5ACAA
# cgBlAGYAZQByAGUAbgBjAGUALjALBglghkgBhv1sAxUwHwYDVR0jBBgwFoAUFQAS
# KxOYspkH7R7for5XDStnAs0wHQYDVR0OBBYEFGFaTSS2STKdSip5GoNL9B6Jwcp9
# MH0GA1UdHwR2MHQwOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydEFzc3VyZWRJRENBLTEuY3JsMDigNqA0hjJodHRwOi8vY3JsNC5kaWdpY2Vy
# dC5jb20vRGlnaUNlcnRBc3N1cmVkSURDQS0xLmNybDB3BggrBgEFBQcBAQRrMGkw
# JAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBBBggrBgEFBQcw
# AoY1aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElE
# Q0EtMS5jcnQwDQYJKoZIhvcNAQEFBQADggEBAJ0lfhszTbImgVybhs4jIA+Ah+WI
# //+x1GosMe06FxlxF82pG7xaFjkAneNshORaQPveBgGMN/qbsZ0kfv4gpFetW7ea
# sGAm6mlXIV00Lx9xsIOUGQVrNZAQoHuXx/Y/5+IRQaa9YtnwJz04HShvOlIJ8Oxw
# YtNiS7Dgc6aSwNOOMdgv420XEwbu5AO2FKvzj0OncZ0h3RTKFV2SQdr5D4HRmXQN
# JsQOfxu19aDxxncGKBXp2JPlVRbwuwqrHNtcSCdmyKOLChzlldquxC5ZoGHd2vNt
# omHpigtt7BIYvfdVVEADkitrwlHCCkivsNRu4PQUCjob4489yq9qjXvc2EQwggbN
# MIIFtaADAgECAhAG/fkDlgOt6gAK6z8nu7obMA0GCSqGSIb3DQEBBQUAMGUxCzAJ
# BgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5k
# aWdpY2VydC5jb20xJDAiBgNVBAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBD
# QTAeFw0wNjExMTAwMDAwMDBaFw0yMTExMTAwMDAwMDBaMGIxCzAJBgNVBAYTAlVT
# MRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5j
# b20xITAfBgNVBAMTGERpZ2lDZXJ0IEFzc3VyZWQgSUQgQ0EtMTCCASIwDQYJKoZI
# hvcNAQEBBQADggEPADCCAQoCggEBAOiCLZn5ysJClaWAc0Bw0p5WVFypxNJBBo/J
# M/xNRZFcgZ/tLJz4FlnfnrUkFcKYubR3SdyJxArar8tea+2tsHEx6886QAxGTZPs
# i3o2CAOrDDT+GEmC/sfHMUiAfB6iD5IOUMnGh+s2P9gww/+m9/uizW9zI/6sVgWQ
# 8DIhFonGcIj5BZd9o8dD3QLoOz3tsUGj7T++25VIxO4es/K8DCuZ0MZdEkKB4YNu
# gnM/JksUkK5ZZgrEjb7SzgaurYRvSISbT0C58Uzyr5j79s5AXVz2qPEvr+yJIvJr
# GGWxwXOt1/HYzx4KdFxCuGh+t9V3CidWfA9ipD8yFGCV/QcEogkCAwEAAaOCA3ow
# ggN2MA4GA1UdDwEB/wQEAwIBhjA7BgNVHSUENDAyBggrBgEFBQcDAQYIKwYBBQUH
# AwIGCCsGAQUFBwMDBggrBgEFBQcDBAYIKwYBBQUHAwgwggHSBgNVHSAEggHJMIIB
# xTCCAbQGCmCGSAGG/WwAAQQwggGkMDoGCCsGAQUFBwIBFi5odHRwOi8vd3d3LmRp
# Z2ljZXJ0LmNvbS9zc2wtY3BzLXJlcG9zaXRvcnkuaHRtMIIBZAYIKwYBBQUHAgIw
# ggFWHoIBUgBBAG4AeQAgAHUAcwBlACAAbwBmACAAdABoAGkAcwAgAEMAZQByAHQA
# aQBmAGkAYwBhAHQAZQAgAGMAbwBuAHMAdABpAHQAdQB0AGUAcwAgAGEAYwBjAGUA
# cAB0AGEAbgBjAGUAIABvAGYAIAB0AGgAZQAgAEQAaQBnAGkAQwBlAHIAdAAgAEMA
# UAAvAEMAUABTACAAYQBuAGQAIAB0AGgAZQAgAFIAZQBsAHkAaQBuAGcAIABQAGEA
# cgB0AHkAIABBAGcAcgBlAGUAbQBlAG4AdAAgAHcAaABpAGMAaAAgAGwAaQBtAGkA
# dAAgAGwAaQBhAGIAaQBsAGkAdAB5ACAAYQBuAGQAIABhAHIAZQAgAGkAbgBjAG8A
# cgBwAG8AcgBhAHQAZQBkACAAaABlAHIAZQBpAG4AIABiAHkAIAByAGUAZgBlAHIA
# ZQBuAGMAZQAuMAsGCWCGSAGG/WwDFTASBgNVHRMBAf8ECDAGAQH/AgEAMHkGCCsG
# AQUFBwEBBG0wazAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29t
# MEMGCCsGAQUFBzAChjdodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRBc3N1cmVkSURSb290Q0EuY3J0MIGBBgNVHR8EejB4MDqgOKA2hjRodHRwOi8v
# Y3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMDqg
# OKA2hjRodHRwOi8vY3JsNC5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURS
# b290Q0EuY3JsMB0GA1UdDgQWBBQVABIrE5iymQftHt+ivlcNK2cCzTAfBgNVHSME
# GDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzANBgkqhkiG9w0BAQUFAAOCAQEARlA+
# ybcoJKc4HbZbKa9Sz1LpMUerVlx71Q0LQbPv7HUfdDjyslxhopyVw1Dkgrkj0bo6
# hnKtOHisdV0XFzRyR4WUVtHruzaEd8wkpfMEGVWp5+Pnq2LN+4stkMLA0rWUvV5P
# sQXSDj0aqRRbpoYxYqioM+SbOafE9c4deHaUJXPkKqvPnHZL7V/CSxbkS3BMAIke
# /MV5vEwSV/5f4R68Al2o/vsHOE8Nxl2RuQ9nRc3Wg+3nkg2NsWmMT/tZ4CMP0qqu
# AHzunEIOz5HXJ7cW7g/DvXwKoO4sCFWFIrjrGBpN/CohrUkxg0eVd3HcsRtLSxwQ
# nHcUwZ1PL1qVCCkQJjGCBAMwggP/AgEBMIGGMHIxCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAv
# BgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EC
# EAS/j+KrDcLYfUmxXxwPKbAwCQYFKw4DAhoFAKBAMBkGCSqGSIb3DQEJAzEMBgor
# BgEEAYI3AgEEMCMGCSqGSIb3DQEJBDEWBBS97smjF8FVDrT3sOfkylR4RCFzNDAN
# BgkqhkiG9w0BAQEFAASCAQCj1L/1Ztz0NCCQE/gJX25289O7IG5VM0ITK7hQL7QI
# lGpFEGcYekpfnZEiJo2CFRt0ljxQwXV7JpdEyf6O7testDRyIb7fNYAQH956PsrC
# GIOtZ61JHR4uvjbRGUiW8NYm/N3Uqg6d+OoNnahGWQA1/hVldxd3z554p8qNAKKp
# 4kKoZfCXPwoIIzU1zFU/boXWYaN+IsKxgCJDO+F2F8JjbpItqcZeKRadtNI/ULkv
# 7zNIzcuO2vktwy/KzScz9X42PRXWQodgE+jpqpCFfz5+kw+eybC4NTAzF9WvuNQf
# 9+zSgahV5crFzvT9ub0JuTeFv+ghu8Ck83+5n+tfTUSyoYICDzCCAgsGCSqGSIb3
# DQEJBjGCAfwwggH4AgEBMHYwYjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lD
# ZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGln
# aUNlcnQgQXNzdXJlZCBJRCBDQS0xAhADAZoCOv9YsWvW1ermF/BmMAkGBSsOAwIa
# BQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0x
# NTEwMDUyMTExMzRaMCMGCSqGSIb3DQEJBDEWBBRSHLPJYLONeQasbazEUxmm/nIy
# YTANBgkqhkiG9w0BAQEFAASCAQAdX5VrPm1iMwwLMThgj6ejfWVDYxQdIML+RTDF
# 2TOtYgbf2+3V4Scska64KKZrrEuouIQQ5Ck22sSRcMWbnhzGSyM3NE73R7FDKx3s
# 0MOr6kUPQtUrE+wP2/4zGsgElKLimhcrt02vPOTjtoiZLCZdrbz+FHxoq59ynBTL
# OU7afWpwfiF6Ha1G8F9mVmNXL/YRNC2tIBCU4cOVFZEmVAfecv3YkKgc7RWN9YAn
# jLBqdM71WDrRXy13zYbyZRyaIeiCZ7doN9p84FS10XyFXSVV6cZURftrtCjBH2fM
# TaTuKFvyIKMnrOr0cjj61WGfIEh+jN0SlfHcMqszFGAu8F41
# SIG # End signature block
