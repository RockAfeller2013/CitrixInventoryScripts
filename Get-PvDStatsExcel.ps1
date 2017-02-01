#Requires -Version 3

<#
.SYNOPSIS
	Creates an Excel spreadsheet containing Citrix XenDesktop Personal vDisk (PvD) usage statistics.
.DESCRIPTION
	Take a XenDesktop Catalog name (or all Catalogs) and gathers PvD usage stats.  
	Creates a Summary worksheet with users who have any PvD with > 90% usage.
	Summary worksheet has the user's Active Directory name and email address.
.PARAMETER XDCatalogName
	XenDesktop Catalog name.  If not entered, process all Catalogs.
.PARAMETER AdminAddress
	Specifies the address of a XenDesktop controller that the PowerShell script will connect to. 
	This can be provided as a host name or an IP address. 
.PARAMETER XDVersion
	Specifies the XenDesktop version.
	5 is for all versions of XenDesktop 5.
	7 is for all versions of XenDesktop 7.
	Default is XenDesktop 7. 
.PARAMETER SmtpServer
	Specifies the optional email server to send emails to users on the Summary worksheet. 
.PARAMETER SmtpPort
	Specifies the SMTP port. 
.PARAMETER UseSSL
	Specifies whether to use SSL for the SmtpServer.
	Default is False.
.PARAMETER From
	Specifies the username for the From email address.
	If SmtpServer is used, this is a required parameter.
.PARAMETER InputFile
	The file that contains the text for the body of the email.
	If SmtpServer is used, this is a required parameter.
	Text is treated as HTML.
.EXAMPLE
	PS C:\PSScript > .\Get-PvDStatsExcel.ps1
	
	Will process all XenDesktop Catalogs.
	AdminAddress is the computer running the script for the AdminAddress.
	Will load Citrix SnapIn for XenDesktop 7.
.EXAMPLE
	PS C:\PSScript > .\Get-PvDStatsExcel.ps1 -XDVersion 5
	
	Will process all XenDesktop Catalogs.
	AdminAddress is the computer running the script for the AdminAddress.
	Will load Citrix SnapIn for XenDesktop 5.
.EXAMPLE
	PS C:\PSScript .\Get-PvDStatsExcel.ps1 -XDCatalogName Win7Image1 -AdminAddress DDC01

	Will gather PvD usage stats for the Win7Image1 Catalog 
	using a Controller named DDC01 for the AdminAddress.
	Will load Citrix SnapIn for XenDesktop 7.
.EXAMPLE
	PS C:\PSScript > .\Get-PvDStatsExcel.ps1 -SmtpServer mail.domain.tld -From XDAdmin@domain.tld -InputFile C:\Scripts\EmailBodyText.txt
	
	Will process all XenDesktop Catalogs.
	AdminAddress is the computer running the script for the AdminAddress.
	Will load Citrix SnapIn for XenDesktop 7.
	Uses mail.domain.tld as the email server.
	Uses the default SMTP port 25.
	Will not use SSL.
	Uses XDAdmin@domain.tld for the email's Sent From address.
	Reads the contents of C:\Scripts\EmailBodyText.txt and uses the contents as the email body.
.EXAMPLE
	PS C:\PSScript > .\Get-PvDStatsExcel.ps1 -SmtpServer smtp.office365.com -SmtpPort 587 -UseSSL -From Webster@CarlWebster.com -InputFile C:\Scripts\EmailBodyText.txt
	
	Will process all XenDesktop Catalogs.
	AdminAddress is the computer running the script for the AdminAddress.
	Will load Citrix SnapIn for XenDesktop 7.
	Uses smtp.office365.com.
	User SMTP port 587.
	Uses Webster@CarlWebster.com for the email's Sent From address (this must be a valid address on the email system).
	Reads the contents of C:\Scripts\EmailBodyText.txt and uses the contents as the email body.
.Example
	Sample input file.
	<br />
	<br />
	Your Personal vDisk space has reached over 90% of its capacity.   <br />
	You should resolve this as soon as possible before your Personal <br />
	vDisk fills up and needs to be reset.  When that happens, you <br />
	will need to reinstall your applications.<br />
	<br />
	You should move your 'User Data' from your C and or PvD drive <br />
	and it should be kept on your home drive.  The free space on your <br />
	workstation is to be used for application installs, temp files, <br />
	and other general system and application configurations.<br />
	<br />
	Let us know if you need any help.<br />
	<br />
	Thanks<br />
	<br />
	<br />
	The XenDesktop Help Desk Team at CarlWebster.com<br />
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	Creates a Microsoft Excel spreadsheet.
	This has only been tested with Microsoft Excel 2010.
	Optional, creates and sends an email for every user on the Summary Worksheet.
.NOTES
	NAME: Get-PvDStatsExcel.ps1
	VERSION: 2.1
	AUTHOR: Carl Webster with help, as usual, from Michael B. Smith
	LASTEDIT: February 1, 2014
#>


[CmdletBinding( SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "P1" ) ]

Param(
	[parameter(ParameterSetName="P1",
	Position = 0, 
	Mandatory=$false )
	] 
	[parameter(ParameterSetName="P2",
	Position = 0, 
	Mandatory=$false )
	] 
	[String]$XDCatalogName="", 

	[parameter(ParameterSetName="P1",
	Position = 1, 
	Mandatory=$false )
	] 
	[parameter(ParameterSetName="P2",
	Position = 1, 
	Mandatory=$false )
	] 
	[string]$AdminAddress="",

	[parameter(ParameterSetName="P1",
	Position = 2, 
	Mandatory=$false )
	] 
	[parameter(ParameterSetName="P2",
	Position = 2, 
	Mandatory=$false )
	] 
	[int]$XDVersion=7,

	[parameter(ParameterSetName="P2",
	Position = 3, 
	Mandatory=$True )
	] 
	[string]$SmtpServer="",

	[parameter(ParameterSetName="P2",
	Position = 4, 
	Mandatory=$False )
	] 
	[int]$SmtpPort=25,

	[parameter(ParameterSetName="P2",
	Position = 5, 
	Mandatory=$False )
	] 
	[switch]$UseSSL=$False,

	[parameter(ParameterSetName="P2",
	Position = 6, 
	Mandatory=$True )
	] 
	[string]$From="",

	[parameter(ParameterSetName="P2",
	Position = 7, 
	Mandatory=$True )
	] 
	[string]$InputFile="")

#Changes to version 2.0
#	Added XDVersion parameter so script could work with either XenDesktop 5.x or 7.x
#	Changed Write-Host statements to Write-Verbose
#	Changed hard coded numbers for colors to variables	
#	Changed hard coded numbers 50 and 90 to variables $WarningLimit and $ErrorLimit
#	If there are no PvDs with > 90% utilization then no Summary worksheet is created
#	Make other code adjustments if there is no Summary worksheet created
#	Delete CSV files after they are used
#	Changed from use .\ for the working folder to $Pwd
#	Removed -Port parameter on Send-MailMessage as it is not supported in PowerShell 2.0
#	Change the code for quitting Excel
#	At the end of the script, if the Excel.exe process is still running stop the process
#	Added elapsed time
#	General code cleanup
#Changes to version 2.1
#	Require PowerShell 3
#	Added SmtpPort and UseSSL parameters
#	Added getting email logon credentials
#	Changed input file to be treated as HTML
#	Added display text to show information about the script environment and parameters

#force -verbose on
$PSDefaultParameterValues = @{"*:Verbose"=$True}
If($SmtpServer -eq $Null)
{
	$SmtpServer = ""
}
If($SmtpPort -eq $Null)
{
	$SmtpPort = 25
}
If($UseSSL -eq $Null)
{
	$UseSSL = $False
}
If($From -eq $Null)
{
	$From = ""
}
If($InputFile -eq $Null)
{
	$InputFile = ""
}

If($XDVersion –eq 5)
{
	$XDSnapinName = "Citrix.Broker.Admin.V1"
}
Else
{
	$XDSnapinName = "Citrix.Broker.Admin.V2"
}
	
Set-StrictMode -Version 2

Function CheckExcelPrereq
{
	If ((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Excel.Application) -eq $False)
	{
		Write-Warning  "This script directly outputs to Microsoft Excel, please install Microsoft Excel"
		Exit
	}

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId
	
	#Find out if excel is running in our session
	[bool]$excelrunning = ((Get-Process 'Excel' -ea 0)|?{$_.SessionId -eq $SessionID}) -ne $null
	If ($excelrunning)
	{
		Write-Warning  "Please close all instances of Microsoft Excel before running this report."
		Exit
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
		Return $False
	}
	Else
	{
		Return $True
	}
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
		$module = Import-Module -Name $ModuleName -PassThru -EA 0
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

#Script begins

$script:startTime = Get-Date

#color values are from http://msdn.microsoft.com/en-us/library/cc296089(v=office.12).aspx
[int]$XLRed = 3
[int]$XLYellow = 6
[int]$XLGrey = 15

#warning and error values
[int]$WarningLimit = 50
[int]$ErrorLimit = 90

If (!(Check-NeededPSSnapins $XDSnapinName))
{
    #We're missing Citrix Snapins that we need
    Write-Error "Missing Citrix PowerShell Snap-ins Detected, check the console above for more information. Script will now close."
    Break
}

If(!(Check-LoadedModule "ActiveDirectory"))
{
	Write-Warning "The Active Directory module does not exist, RSAT needs to be installed, script cannot continue."
	Break
}

CheckExcelPrereq

$Params = @{
Filter = "{ Name -eq $XDCatalogName }";
MaxRecordCount = 65536; 
EA = 0;
}

If(![System.String]::IsNullOrEmpty( $AdminAddress ))
{
	$Params.AdminAddress = $AdminAddress
}

If( ![System.String]::IsNullOrEmpty( $XDCatalogName ) )
{
	$status = $(try {Get-BrokerCatalog @Params} catch {$null})
	If ($status -ne $null) 
	{
	 	Write-Verbose  "$(Get-Date): $XDCatalogName is a valid XenDesktop Catalog and is being retrieved"
		$Params.CatalogName = $XDCatalogName
	} 
	Else 
	{
	 	Write-Warning "Catalog $XDCatalogName was not found.  Script cannot continue."
	 	Exit
	}
}
Else
{
	Write-Verbose  "$(Get-Date): Retrieving all XenDesktop Catalog names"	
}

If(![System.String]::IsNullOrEmpty( $InputFile ))
{
	If ( !(Test-Path $InputFile) )
	{
	 	Write-Warning "The input file $InputFile was not found.  Script cannot continue."
	 	Exit
	}
}

Write-Verbose "$(Get-Date): "
Write-Verbose "$(Get-Date): "
If([System.String]::IsNullOrEmpty( $AdminAddress ))
{
	Write-Verbose "$(Get-Date): Using Controller   : $env:ComputerName"
}
Else
{
	Write-Verbose "$(Get-Date): Using Controller   : $AdminAddress"
}
Write-Verbose "$(Get-Date): XenDesktop version : $XDVersion"
If(![System.String]::IsNullOrEmpty( $SmtpServer ))
{
	Write-Verbose "$(Get-Date): Smtp Server        : $SmtpServer"
	Write-Verbose "$(Get-Date): Smtp Port          : $SmtpPort"
	Write-Verbose "$(Get-Date): Use SSL            : $UseSSL"
	Write-Verbose "$(Get-Date): From               : $From"
	Write-Verbose "$(Get-Date): Input File         : $InputFile"
}
Write-Verbose "$(Get-Date): PoSH version       : $($Host.Version)"
Write-Verbose "$(Get-Date): "
Write-Verbose "$(Get-Date): Script start       : $($Script:StartTime)"
Write-Verbose "$(Get-Date): "
Write-Verbose "$(Get-Date): "

#only get catalogs where desktops are registered, have a DNSName and are powered on
#if all the desktops in a catalog are powered off, there is no need to to get the PvD stats

$Params = @{
Filter = "{ RegistrationState -eq 'Registered' -and DNSName -ne '' -and PowerState -eq 'On'}";
MaxRecordCount = 65536; 
EA = 0;
}

If(![System.String]::IsNullOrEmpty( $AdminAddress ))
{
	$Params.AdminAddress = $AdminAddress
}

$Catalogs = Get-BrokerDesktop @Params | Select -unique CatalogName | Sort-Object CatalogName

If( !($?) -or $Catalogs -eq $Null)
{
	Write-Warning "Unable to retrieve Catalog Names.  Script cannot proceed."
	Write-Error “cmdlet failed $($error[ 0 ].ToString())”
     Return $null
}

If($Catalogs -is [Array])
{
    $TotalCatalogs = $Catalogs.count
}
Else
{
    $TotalCatalogs = 1
}

#Create CSV files first
$CSVFiles = @()
$SummaryObjects = @()

ForEach($Catalog in $Catalogs)
{
	#name the CSV file for the catalog
	$FileName = $Catalog.CatalogName 
	
	Write-Verbose  "$(Get-Date): Get VMs for Catalog $($Catalog.CatalogName)"		
	#only get VMs that are registered, have a DNSName and are powered on
	#if a VM is powered off, there is no use in getting PvD stats as the PvD Service cannot be reached
	$VMs = Get-BrokerDesktop `
	-Filter { RegistrationState -eq 'Registered' -and DNSName -ne "" -and PowerState -eq "On"} `
	-CatalogName $Catalog.CatalogName -MaxRecordCount 65536 -EA 0  | `
	Sort-Object HostedMachineName

	If( $? -and $VMs )
	{
		
		$VMName                     = ""
		$UserName                   = ""
		$ADUserName                 = ""
		$PVDServiceStatus           = ""
		$PVDStatus                  = ""
		[double]$AppGB              = 0.0
		[double]$AppPercentUsed     = 0.00
		[double]$ProfileGB          = 0.0
		[double]$ProfilePercentUsed = 0.00
		[double]$TotalPercentUsed   = 0.00
		$UpdateStatus               = ""
		$PowerState                 = ""
		$PVDObjects                 = @()
		$EmailAddress               = ""
		$tmp                        = ""
	
		ForEach($VM in $VMs)
		{
			$VMName = $VM.HostedMachineName
			Write-Verbose  "$(Get-Date): `tProcessing VM $VMName"
			$PowerState =  $VM.powerstate
			If($VM.AssociatedUserFullNames -ne $null)
			{
				$UserName     = $VM.AssociatedUserFullNames[0]
				$UserADName   = $VM.AssociatedUserNames[0]
				$tmp          = ($UserADName.Split("\")[1])
				$EmailAddress = $(try {(Get-ADUser $tmp -Properties mail -EA 0).mail} catch {$null})
				
				If($EmailAddress -eq $Null -or $EmailAddress -eq "")
				{
					$EmailAddress = "Unassigned"
				}
			}
			Else
			{
				$UserName     = "Unassigned"
				$UserADName   = "Unassigned"
				$EmailAddress = "Unassigned"
			}
							
			If ($PowerState -eq "On")
			{
				$vns = $null
				try
				{
					$vns = Get-Service -ComputerName $VMName -name "Citrix Personal vDisk" -EA 0
				}
				catch
				{
					write-warning "Get-Service -ComputerName $($VMName) -name Citrix Personal vDisk: failed"
				}
				If($vns -eq $null)
				{
					$PVDServiceStatus   = "Unknown"
					$PVDStatus          = "Unknown"
					$AppGB              = 0.0
					$AppPercentUsed     = 0.00
					$ProfileGB          = 0.0
					$ProfilePercentUsed = 0.00
					$TotalPercentUsed   = 0.00
					$UpdateStatus       = "Unknown"
				}
				Else
				{
					$PVDServiceStatus = $vns.status
					$PvDPool = $null
					try
					{
						$PvDPool = Get-WmiObject -ComputerName $VMName -Namespace root\Citrix -Class Citrix_PvDPool -EA 0
					}
					catch
					{
						write-warning "Get-WmiObject -ComputerName $($VMName) -Namespace root\Citrix -Class Citrix_PvDPool: failed"
					}
					$PvDOk=$false
					If ($PvDPool -ne $null) 
					{
						If ( $PvDPool.IsActive -eq $true )
						{

							If([long]$PvDPool.TotalAppSizeBytes -eq 0)
							{
								$AppPercentUsed = 0.00
							}
							Else
							{
								$AppPercentUsed = ([long]$PvDPool.CurrentAppSizeBytes / [long]$PvDPool.TotalAppSizeBytes)
							}
							$AppGB = [long]$PvDPool.TotalAppSizeBytes / 1GB
							
							If([long]$PvDPool.TotalProfSizeBytes -eq 0)
							{
								$ProfilePercentUsed = 0.00
							}
							Else
							{
								$ProfilePercentUsed = ([long]$PvDPool.CurrentProfSizeBytes / [long]$PvDPool.TotalProfSizeBytes)
							}
							$ProfileGB = [long]$PvDPool.TotalProfSizeBytes / 1GB
							
							If(([long]$PvDPool.TotalAppSizeBytes + [long]$PvDPool.TotalProfSizeBytes) -eq 0)
							{
								$TotalPercentUsed = 0.00
							}
							Else
							{
								$TotalPercentUsed = ([long]$PvDPool.CurrentAppSizeBytes + [long]$PvDPool.CurrentProfSizeBytes) / `
												([long]$PvDPool.TotalAppSizeBytes + [long]$PvDPool.TotalProfSizeBytes)
							}
		
							$PVDStatus = "Running"
							$PvDOk     = $true
						}
						Else
						{
							$AppPercentUsed     = 0.00
							$AppGB              = 0.0
							$ProfilePercentUsed = 0.00
							$ProfileGB          = 0.0
							$TotalPercentUsed   = 0.00
							$PVDStatus          = "No"
						}
					}
					Else
					{
						$AppPercentUsed     = 0.00
						$AppGB              = 0.0
						$ProfilePercentUsed = 0.00
						$ProfileGB          = 0.0
						$TotalPercentUsed   = 0.00
						$PVDStatus          = "No"
					}
						
					# No PvD, check image update state
					If($PvDPool -eq $Null)
					{
						$UpdateStatus = "Unknown"
					}
					ElseIf ($PvDPool -ne $Null -and $PvDOk -eq $false)
					{
						If ($PvDPool.StatusText -ne "") 
						{
							$UpdateStatus = $PvDPool.StatusText
						}
						Else
						{
							$UpdateStatus = "Unknown"
						}
					}
					Else
					{
						$UpdateStatus = "OK"
					}
				}
			
			}
			Else
			{
				$PVDServiceStatus   = "VM Off"
				$PVDStatus          = "VM Off"
				$AppPercentUsed     = 0.00
				$AppGB              = 0.0
				$ProfilePercentUsed = 0.00
				$ProfileGB          = 0.0
				$TotalPercentUsed   = 0.00
				$UpdateStatus       = "VM Off"
			}
	
			$AppPercentUsed     *= 100
			$ProfilePercentUsed *= 100
			$TotalPercentUsed   *= 100
			
			$PVDObject = New-Object -Type PSObject -Property @{   
				VMName             = $VMName
				UserName           = $UserName         
				UserADName         = $UserADName         
				PVDServiceStatus   = $PVDServiceStatus
				PVDStatus          = $PVDStatus
				AppGB              = "{0:N1}" -f $AppGB
				AppPercentUsed     = "{0:F2}" -f $AppPercentUsed
				ProfileGB          = "{0:N1}" -f $ProfileGB
				ProfilePercentUsed = "{0:F2}" -f $ProfilePercentUsed
				TotalPercentUsed   = "{0:F2}" -f $TotalPercentUsed
				UpdateStatus       = $UpdateStatus
				}
					
			$PVDObjects += $PVDObject
			
			#gather summary data
			If(([double]$AppPercentUsed -ge $ErrorLimit) -or `
				([double]$ProfilePercentUsed -ge $ErrorLimit) -or `
				([double]$TotalPercentUsed -ge $ErrorLimit))
			{
				$SummaryObject = New-Object -Type PSObject -Property @{   
					VMName             = $VMName
					UserName           = $UserName         
					UserADName         = $UserADName         
					EmailAddress       = $EmailAddress
					AppGB              = "{0:N1}" -f $AppGB
					AppPercentUsed     = "{0:F2}" -f $AppPercentUsed
					ProfileGB          = "{0:N1}" -f $ProfileGB
					ProfilePercentUsed = "{0:F2}" -f $ProfilePercentUsed
					TotalPercentUsed   = "{0:F2}" -f $TotalPercentUsed
					}
				$SummaryObjects += $SummaryObject
			}
		}
		
		Write-Verbose  "$(Get-Date): Creating CSV file $($pwd.path)\$($FileName)_PvD_Stats.csv"
		$PVDObjects | `
		select-object VMName, UserName, UserADName, PVDServiceStatus, PVDStatus,
		AppGB, AppPercentUsed, ProfileGB, ProfilePercentUsed, TotalPercentUsed, UpdateStatus | `
		Sort-Object VMname | `
		Export-CSV "$($pwd.path)\$($FileName)_PvD_Stats.csv" -NoTypeInformation
		$CSVFiles += "$($pwd.path)\$($FileName)_PvD_Stats.csv"
	}
	Else
	{
		Write-Warning "Unable to retrieve VMs for Catalog $($Catalog.CatalogName)"
	}
}

[bool]$DoSummary = $False
If($SummaryObjects -ne $Null)
{
	$DoSummary = $True
	Write-Verbose  "$(Get-Date): Creating Summary CSV file $($pwd.path)\Summary_PvD_Stats.csv"
	$SummaryObjects | `
	Select-Object VMName, UserName, UserADName, EmailAddress, AppGB, 
	AppPercentUsed, ProfileGB, ProfilePercentUsed, TotalPercentUsed | `
	Sort-Object VMname | `
	Export-CSV "$($pwd.path)\Summary_PvD_Stats.csv" -NoTypeInformation
	$CSVFiles += "$($pwd.path)\Summary_PvD_Stats.csv"

	If(![System.String]::IsNullOrEmpty( $SmtpServer ))
	{
		Write-Verbose  "$(Get-Date): Sending emails"
		#send an email to each person on the summary worksheet
		$emailSmtpServer = $SmtpServer
		[int]$emailSmtpPort = $SmtpPort
		$emailFrom = $From
		$emailBodyText = Get-Content $InputFile
		$emailCredentials = Get-Credential -Message "Enter the email account and password to send emails"
		ForEach($Item in $SummaryObjects)
		{
			If($Item.EmailAddress -ne "Unassigned")
			{
				Write-Verbose  "$(Get-Date): `tSending email to $($Item.EmailAddress)"
				$emailTo = $Item.EmailAddress
				$emailSubject = "$($Item.VMName) - Action needed to continue use of your Virtual Desktop"
				$emailBody = @"
`nHello $($Item.UserName),

$($emailBodyText)
"@ 
				If($UseSSL)
				{
					Send-MailMessage -To $emailTo -Subject $emailSubject -Body $emailBody -SmtpServer $emailSmtpServer -From $emailFrom -Port $emailSmtpPort -BodyAsHtml -credential $emailCredentials -UseSSL
				}
				Else
				{
					Send-MailMessage -To $emailTo -Subject $emailSubject -Body $emailBody -SmtpServer $emailSmtpServer -From $emailFrom -Port $emailSmtpPort -BodyAsHtml -credential $emailCredentials
				}
			}
		}
	}
}
Else
{
	$DoSummary = $False
	Write-Verbose  "$(Get-Date): No Summary Data"
}

Write-Verbose  "$(Get-Date): CSV Processing complete"
Write-Verbose  "$(Get-Date): Start creating Excel file and worksheets"

# Setup excel for output
Write-Verbose  "$(Get-Date): Setup Excel"
$Excel = New-Object -com Excel.Application

$Excel.Visible = $False

If($TotalCatalogs -lt 1)
{
	If($DoSummary)
	{
		$Excel.sheetsInNewWorkbook = 2
	}
	Else
	{
		$Excel.sheetsInNewWorkbook = 1
	}
}
Else
{
	If($DoSummary)
	{
		$Excel.sheetsInNewWorkbook = $TotalCatalogs + 1
	}
	Else
	{
		$Excel.sheetsInNewWorkbook = $TotalCatalogs
	}
}

$WB = $Excel.WorkBooks.Add()
#get active worksheet
$WS = $WB.ActiveSheet
[int]$i = 0

If($DoSummary)
{
	#process summary CSV file first so it is the first worksheet
	$i++
	$WS = $WB.WorkSheets.Item("Sheet$i")

	[void] $WS.Activate() 
	$WS.Application.ActiveWindow.SplitRow = 1
	$WS.Application.ActiveWindow.FreezePanes = $true

	#name the worksheet for the catalog
	$WS.Name = "Summary"

	Write-Verbose  "$(Get-Date): Get Summary CSV file"		

	$CSVFile  = "$($pwd.path)\Summary_PvD_Stats.csv"
	If((Test-Path "$($pwd.path)\Summary_PvD_Stats.csv"))
	{
		$Stats    = Import-Csv -Path $csvFile

		$Cells = $WS.Cells
		[int]$xRow = 1
		$Cells.Item($xRow,1).Interior.ColorIndex = $XLGrey
		$Cells.Item($xRow,1).Font.Bold = $True
		$Cells.Item($xRow,1) = "VMName"
		$Cells.Item($xRow,2).Interior.ColorIndex = $XLGrey
		$Cells.Item($xRow,2).Font.Bold = $True
		$Cells.Item($xRow,2) = "User Name"
		$Cells.Item($xRow,3).Interior.ColorIndex = $XLGrey
		$Cells.Item($xRow,3).Font.Bold = $True
		$Cells.Item($xRow,3) = "User AD Name"
		$Cells.Item($xRow,4).Interior.ColorIndex = $XLGrey
		$Cells.Item($xRow,4).Font.Bold = $True
		$Cells.Item($xRow,4) = "Email Address"
		$Cells.Item($xRow,5).Interior.ColorIndex = $XLGrey
		$Cells.Item($xRow,5).Font.Bold = $True
		$Cells.Item($xRow,5) = "App GB"
		$Cells.Item($xRow,6).Interior.ColorIndex = $XLGrey
		$Cells.Item($xRow,6).Font.Bold = $True
		$Cells.Item($xRow,6) = "App % Used"
		$Cells.Item($xRow,7).Interior.ColorIndex = $XLGrey
		$Cells.Item($xRow,7).Font.Bold = $True
		$Cells.Item($xRow,7) = "Profile GB"
		$Cells.Item($xRow,8).Interior.ColorIndex = $XLGrey
		$Cells.Item($xRow,8).Font.Bold = $True
		$Cells.Item($xRow,8) = "Profile % Used"
		$Cells.Item($xRow,9).Interior.ColorIndex = $XLGrey
		$Cells.Item($xRow,9).Font.Bold = $True
		$Cells.Item($xRow,9) = "Total % Used"

		$xrow = 2

		ForEach($Stat in $Stats)
		{
			Write-Verbose  "$(Get-Date): `tAdding summary row for $($Stat.VMName)"
			
			$Cells.Item($xRow,1) = $stat.VMName
			$Cells.Item($xRow,2) = $Stat.UserName
			$Cells.Item($xRow,3) = $Stat.UserADName
			$Cells.Item($xRow,4) = $Stat.EmailAddress
			If([double]$Stat.AppGB -eq 0)
			{
				$Cells.Item($xRow,5).Interior.ColorIndex = $XLRed
				$Cells.Item($xRow,5).Font.Bold = $True
			}
			$Cells.Item($xRow,5) = [double]$Stat.AppGB
			Switch ([double]$Stat.AppPercentUsed)
			{
				{($_ -ge $WarningLimit)  -and ($_ -lt $ErrorLimit)} 
					{
						$Cells.Item($xRow,6).Interior.ColorIndex = $XLYellow
						$Cells.Item($xRow,6).Font.Bold = $True
					}
				{($_ -ge $ErrorLimit) -or ($_ -eq 0) } 
					{
						$Cells.Item($xRow,6).Interior.ColorIndex = $XLRed
						$Cells.Item($xRow,6).Font.Bold = $True
				}
			}
			
			$Cells.Item($xRow,6) = [double]$Stat.AppPercentUsed
			If([double]$Stat.ProfileGB -eq 0)
			{
				$Cells.Item($xRow,7).Interior.ColorIndex = $XLRed
				$Cells.Item($xRow,7).Font.Bold = $True
			}
			$Cells.Item($xRow,7) = [double]$Stat.ProfileGB
			Switch ([double]$Stat.ProfilePercentUsed)
			{
				{($_ -ge $WarningLimit) -and ($_ -lt $ErrorLimit)} 
					{
						$Cells.Item($xRow,8).Interior.ColorIndex = $XLYellow
						$Cells.Item($xRow,8).Font.Bold = $True
				}
				{($_ -ge $ErrorLimit) -or ($_ -eq 0)} 
					{
						$Cells.Item($xRow,8).Interior.ColorIndex = $XLRed
						$Cells.Item($xRow,8).Font.Bold = $True
					}
			}
			$Cells.Item($xRow,8) = [double]$Stat.ProfilePercentUsed
			Switch ([double]$Stat.TotalPercentUsed)
			{
				{($_ -ge $WarningLimit) -and ($_ -lt $ErrorLimit)} 
					{
						$Cells.Item($xRow,9).Interior.ColorIndex = $XLYellow
						$Cells.Item($xRow,9).Font.Bold = $True
					}
				{($_ -ge $ErrorLimit) -or ($_ -eq 0) } 
					{
						$Cells.Item($xRow,9).Interior.ColorIndex = $XLRed
						$Cells.Item($xRow,9).Font.Bold = $True
					}
			}
			$Cells.Item($xRow,9) = [double]$Stat.TotalPercentUsed
			
			$xRow++
		}
		$ws.columns.item("E:E").EntireColumn.NumberFormat = "#0.0"
		$ws.columns.item("F:F").EntireColumn.NumberFormat = "#0.00"
		$ws.columns.item("G:G").EntireColumn.NumberFormat = "#0.0"
		$ws.columns.item("H:I").EntireColumn.NumberFormat = "#0.00"
		$ws.columns.item("A:I").EntireColumn.AutoFit() | out-null
		
		#no longer need CSV file so delete it
		Write-Verbose "$(Get-Date): Deleting $($csvFile)"
		Remove-Item $csvFile -EA 0		
		Write-Verbose  "$(Get-Date): Summary sheet completed"
	}
}
Else
{
	Write-Verbose  "$(Get-Date): No Summary information to create a Summary worksheet"
}
Write-Verbose  "$(Get-Date): Start processing catalogs"

#now process the rest of the CSV files
ForEach($Catalog in $Catalogs)
{
	#increment $i to get the next worksheet
	$i++
	$WS = $WB.WorkSheets.Item("Sheet$i")
	
	[void] $WS.Activate() 
	$WS.Application.ActiveWindow.SplitRow = 1
	$WS.Application.ActiveWindow.FreezePanes = $true
	
	#name the worksheet for the catalog
	$WS.Name = $Catalog.CatalogName 
	
	Write-Verbose  "$(Get-Date): Get CSV file for Catalog $($Catalog.CatalogName)"		

	#the CSV file is named for the catalog
	$FileName = $Catalog.CatalogName 
	$CSVFile  = "$($pwd.path)\$($FileName)_PvD_Stats.csv"
	$Stats    = Import-Csv -Path $csvFile

	$Cells = $WS.Cells
	[int]$xRow = 1
	$Cells.Item($xRow,1).Interior.ColorIndex = $XLGrey
	$Cells.Item($xRow,1).Font.Bold = $True
	$Cells.Item($xRow,1) = "VMName"
	$Cells.Item($xRow,2).Interior.ColorIndex = $XLGrey
	$Cells.Item($xRow,2).Font.Bold = $True
	$Cells.Item($xRow,2) = "User Name"
	$Cells.Item($xRow,3).Interior.ColorIndex = $XLGrey
	$Cells.Item($xRow,3).Font.Bold = $True
	$Cells.Item($xRow,3) = "User AD Name"
	$Cells.Item($xRow,4).Interior.ColorIndex = $XLGrey
	$Cells.Item($xRow,4).Font.Bold = $True
	$Cells.Item($xRow,4) = "PvD Service Status"
	$Cells.Item($xRow,5).Interior.ColorIndex = $XLGrey
	$Cells.Item($xRow,5).Font.Bold = $True
	$Cells.Item($xRow,5) = "PvD Status"
	$Cells.Item($xRow,6).Interior.ColorIndex = $XLGrey
	$Cells.Item($xRow,6).Font.Bold = $True
	$Cells.Item($xRow,6) = "App GB"
	$Cells.Item($xRow,7).Interior.ColorIndex = $XLGrey
	$Cells.Item($xRow,7).Font.Bold = $True
	$Cells.Item($xRow,7) = "App % Used"
	$Cells.Item($xRow,8).Interior.ColorIndex = $XLGrey
	$Cells.Item($xRow,8).Font.Bold = $True
	$Cells.Item($xRow,8) = "Profile GB"
	$Cells.Item($xRow,9).Interior.ColorIndex = $XLGrey
	$Cells.Item($xRow,9).Font.Bold = $True
	$Cells.Item($xRow,9) = "Profile % Used"
	$Cells.Item($xRow,10).Interior.ColorIndex = $XLGrey
	$Cells.Item($xRow,10).Font.Bold = $True
	$Cells.Item($xRow,10) = "Total % Used"
	$Cells.Item($xRow,11).Interior.ColorIndex = $XLGrey
	$Cells.Item($xRow,11).Font.Bold = $True
	$Cells.Item($xRow,11) = "Update Status"

	$xrow = 2
	
	ForEach($Stat in $Stats)
	{
		Write-Verbose  "$(Get-Date): `tAdding row for $($Stat.VMName)"
		
		$Cells.Item($xRow,1) = $stat.VMName
		$Cells.Item($xRow,2) = $Stat.UserName
		$Cells.Item($xRow,3) = $Stat.UserADName
		If($Stat.PVDServiceStatus -ne "Running")
		{
			$Cells.Item($xRow,4).Interior.ColorIndex = $XLRed
			$Cells.Item($xRow,4).Font.Bold = $True
		}
		$Cells.Item($xRow,4) = $Stat.PVDServiceStatus
		If($Stat.PVDStatus -ne "Running")
		{
			$Cells.Item($xRow,5).Interior.ColorIndex = $XLRed
			$Cells.Item($xRow,5).Font.Bold = $True
		}
		$Cells.Item($xRow,5) = $Stat.PVDStatus
		If([double]$Stat.AppGB -eq 0)
		{
			$Cells.Item($xRow,6).Interior.ColorIndex = $XLRed
			$Cells.Item($xRow,6).Font.Bold = $True
		}
		$Cells.Item($xRow,6) = [double]$Stat.AppGB
		Switch ([double]$Stat.AppPercentUsed)
		{
			{($_ -ge $WarningLimit)  -and ($_ -lt $ErrorLimit)} 
				{
					$Cells.Item($xRow,7).Interior.ColorIndex = $XLYellow
					$Cells.Item($xRow,7).Font.Bold = $True
				}
			{($_ -ge $ErrorLimit) -or ($_ -eq 0) } 
				{
					$Cells.Item($xRow,7).Interior.ColorIndex = $XLRed
					$Cells.Item($xRow,7).Font.Bold = $True
			}
		}
		
		$Cells.Item($xRow,7) = [double]$Stat.AppPercentUsed
		If([double]$Stat.ProfileGB -eq 0)
		{
			$Cells.Item($xRow,8).Interior.ColorIndex = $XLRed
			$Cells.Item($xRow,8).Font.Bold = $True
		}
		$Cells.Item($xRow,8) = [double]$Stat.ProfileGB
		Switch ([double]$Stat.ProfilePercentUsed)
		{
			{($_ -ge $WarningLimit) -and ($_ -lt $ErrorLimit)} 
				{
					$Cells.Item($xRow,9).Interior.ColorIndex = $XLYellow
					$Cells.Item($xRow,9).Font.Bold = $True
			}
			{($_ -ge $ErrorLimit) -or ($_ -eq 0)} 
				{
					$Cells.Item($xRow,9).Interior.ColorIndex = $XLRed
					$Cells.Item($xRow,9).Font.Bold = $True
				}
		}
		$Cells.Item($xRow,9) = [double]$Stat.ProfilePercentUsed
		Switch ([double]$Stat.TotalPercentUsed)
		{
			{($_ -ge $WarningLimit) -and ($_ -lt $ErrorLimit)} 
				{
					$Cells.Item($xRow,10).Interior.ColorIndex = $XLYellow
					$Cells.Item($xRow,10).Font.Bold = $True
				}
			{($_ -ge $ErrorLimit) -or ($_ -eq 0) } 
				{
					$Cells.Item($xRow,10).Interior.ColorIndex = $XLRed
					$Cells.Item($xRow,10).Font.Bold = $True
				}
		}
		$Cells.Item($xRow,10) = [double]$Stat.TotalPercentUsed
		If($Stat.UpdateStatus -ne "OK")
		{
			$Cells.Item($xRow,11).Interior.ColorIndex = $XLRed
			$Cells.Item($xRow,11).Font.Bold = $True
		}
		$Cells.Item($xRow,11) = $Stat.UpdateStatus
		
		$xRow++
	}
	$ws.columns.item("F:F").EntireColumn.NumberFormat = "#0.0"
	$ws.columns.item("G:G").EntireColumn.NumberFormat = "#0.00"
	$ws.columns.item("H:H").EntireColumn.NumberFormat = "#0.0"
	$ws.columns.item("I:J").EntireColumn.NumberFormat = "#0.00"
	$ws.columns.item("A:K").EntireColumn.AutoFit() | out-null
	#no longer need CSV file so delete it
	Write-Verbose "$(Get-Date): Deleting $($csvFile)"
	Remove-Item $csvFile -EA 0		
}

If($DoSummary)
{
	#activate the Summary worksheet so when the file is opened, the Summary sheet is displayed
	$WS = $WB.WorkSheets.Item("Summary")
	[void] $WS.Activate() 
}

Write-Verbose  "$(Get-Date): Processing worksheets is complete"
Write-Verbose  "$(Get-Date): Saving Excel file"
$Excel.DisplayAlerts = $False
#xlsx 
$xlOpenXMLWorkbook = 51
$wb.saveas("$($pwd.path)\PvDStats_$(Get-Date -f yyyy-MM-dd).xlsx", $xlOpenXMLWorkbook)
$Excel.Quit()
#[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
While( [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Cells)){}
While( [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WS)){}
While( [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WB)){}
While( [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)){}
Remove-Variable Excel
Write-Verbose  "$(Get-Date): Excel file $($pwd.path)\PvDStats_$(Get-Date -f yyyy-MM-dd).xlsx is ready for use"

#If the Excel.exe process is still running for the user's sessionID, kill it
$SessionID = (Get-Process -PID $PID).SessionId
(Get-Process 'Excel' -ea 0 | ?{$_.sessionid -eq $Sessionid}) | stop-process

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