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
# SIG # Begin signature block
# MIIiywYJKoZIhvcNAQcCoIIivDCCIrgCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUCENi/SowMCxz1TdrgV9S2j87
# mJ+ggh41MIIDtzCCAp+gAwIBAgIQDOfg5RfYRv6P5WD8G/AwOTANBgkqhkiG9w0B
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
# 8jCCBmowggVSoAMCAQICEAOf7e3LeVuN7TIMiRnwNokwDQYJKoZIhvcNAQEFBQAw
# YjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQ
# d3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgQXNzdXJlZCBJRCBD
# QS0xMB4XDTEzMDUyMTAwMDAwMFoXDTE0MDYwNDAwMDAwMFowRzELMAkGA1UEBhMC
# VVMxETAPBgNVBAoTCERpZ2lDZXJ0MSUwIwYDVQQDExxEaWdpQ2VydCBUaW1lc3Rh
# bXAgUmVzcG9uZGVyMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAumlK
# gU1vpRQWqorNZ75Lv8Zpj1gc4HnoHp1YJpjaXNR8o/nbK4wSNsP8+WQGsbvCqJgK
# Fw3hletAtOuWbZi/po95z7yKknttnBgGUdilGFMyAScZYeiEQd/G8OjK/netX9ie
# e4xgb4VcRr1r5w+AzucDw3wxz7dlVcb74JkI5HNa+5fa0Ey+tLbGD38mkqm4/Dju
# tOQ6pEjQTOqpRidbz5IRk5wWp/7SrR8ixR6swXHvvErbAQlE35gcLWe6qIoDM8lR
# tfcCTQmkTf6AXsXXRcN9CKoBM8wz2E8wFuT/IjIu63478PkeMuuVJdLy/m1UhLrV
# 5dTR3RuvvVl7lIUwAQIDAQABo4IDNTCCAzEwDgYDVR0PAQH/BAQDAgeAMAwGA1Ud
# EwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwggG/BgNVHSAEggG2MIIB
# sjCCAaEGCWCGSAGG/WwHATCCAZIwKAYIKwYBBQUHAgEWHGh0dHBzOi8vd3d3LmRp
# Z2ljZXJ0LmNvbS9DUFMwggFkBggrBgEFBQcCAjCCAVYeggFSAEEAbgB5ACAAdQBz
# AGUAIABvAGYAIAB0AGgAaQBzACAAQwBlAHIAdABpAGYAaQBjAGEAdABlACAAYwBv
# AG4AcwB0AGkAdAB1AHQAZQBzACAAYQBjAGMAZQBwAHQAYQBuAGMAZQAgAG8AZgAg
# AHQAaABlACAARABpAGcAaQBDAGUAcgB0ACAAQwBQAC8AQwBQAFMAIABhAG4AZAAg
# AHQAaABlACAAUgBlAGwAeQBpAG4AZwAgAFAAYQByAHQAeQAgAEEAZwByAGUAZQBt
# AGUAbgB0ACAAdwBoAGkAYwBoACAAbABpAG0AaQB0ACAAbABpAGEAYgBpAGwAaQB0
# AHkAIABhAG4AZAAgAGEAcgBlACAAaQBuAGMAbwByAHAAbwByAGEAdABlAGQAIABo
# AGUAcgBlAGkAbgAgAGIAeQAgAHIAZQBmAGUAcgBlAG4AYwBlAC4wCwYJYIZIAYb9
# bAMVMB8GA1UdIwQYMBaAFBUAEisTmLKZB+0e36K+Vw0rZwLNMB0GA1UdDgQWBBRj
# L8nfeZJ7tSPKu+Gk7jN+4+Kd+jB9BgNVHR8EdjB0MDigNqA0hjJodHRwOi8vY3Js
# My5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURDQS0xLmNybDA4oDagNIYy
# aHR0cDovL2NybDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEQ0EtMS5j
# cmwwdwYIKwYBBQUHAQEEazBpMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdp
# Y2VydC5jb20wQQYIKwYBBQUHMAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydEFzc3VyZWRJRENBLTEuY3J0MA0GCSqGSIb3DQEBBQUAA4IBAQCr
# dL1AAEx2FSVXPdMcA/99RchFEmbnKGVg2N87s/oNwawzj/SBuWHxnfuYVdfeR0O6
# gD3xSMw/ZzBWH8700EyEvYeknsXhD6gGXdAvbl7cGejwh+rgTq89bCCOc29+1ocY
# 4IbTmvye6oxy6UEPuHG1OCz4KbLVHKKdG+xfKrjcNyDhy7vw0GxspbPLn0r2VOMm
# ND0uuMErHLf2wz3+0S0eUPSUyPj97nPbSbUb9PX/pZDBORQb2O1xG2qY+/pAmkSp
# KQ5VXni4t6SDw3AB8GZA5a55NOErTQOhLebbVGIY7dUJi6Kq1gzITxq+mSV4aZmJ
# 1FmJ3t+I8NNnXnSlnaZEMIIGkDCCBXigAwIBAgIQBKVRftX3ANDrw0+OjYS9xjAN
# BgkqhkiG9w0BAQUFADBvMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQg
# SW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMS4wLAYDVQQDEyVEaWdpQ2Vy
# dCBBc3N1cmVkIElEIENvZGUgU2lnbmluZyBDQS0xMB4XDTExMDkzMDAwMDAwMFoX
# DTE0MTAwODEyMDAwMFowXDELMAkGA1UEBhMCVVMxCzAJBgNVBAgTAlROMRIwEAYD
# VQQHEwlUdWxsYWhvbWExFTATBgNVBAoTDENhcmwgV2Vic3RlcjEVMBMGA1UEAxMM
# Q2FybCBXZWJzdGVyMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAz2g4
# Kup2X6Mscbuq96HnetDDiITbncV1LtQ8Rxf8ZtN00+O/TliIZsWtufMq7GsLj1D8
# ikWfcgWGqMngWMsVYB4vdr1B8aQuHmKWld7W+j8FhKp3l+rNuFviTGa62sR6fEVW
# 1N6lDtJJHpfSIg/FUFfAqOKl0gFc45PU7iWCh08+oG5FJdhZ3WY0SosS1QujKEA4
# riSjeXPV6XSLsAHTE/fmHlGuu7NzJyMUzNNz2gPOFxYupHygbduhM5aAItD6GJ1h
# ajlovRt71tAMyeIPWNjj9B2luXxfRbgO9eufw91uFrXnougBPa7/eQ25YdW3NcGf
# tosYjvVI6Ptw/AaSiQIDAQABo4IDOTCCAzUwHwYDVR0jBBgwFoAUe2jOKarAF75J
# euHlP9an90WPNTIwHQYDVR0OBBYEFMHndyU+4pRT+JRECX9EG4y1laDkMA4GA1Ud
# DwEB/wQEAwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzBzBgNVHR8EbDBqMDOgMaAv
# hi1odHRwOi8vY3JsMy5kaWdpY2VydC5jb20vYXNzdXJlZC1jcy0yMDExYS5jcmww
# M6AxoC+GLWh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9hc3N1cmVkLWNzLTIwMTFh
# LmNybDCCAcQGA1UdIASCAbswggG3MIIBswYJYIZIAYb9bAMBMIIBpDA6BggrBgEF
# BQcCARYuaHR0cDovL3d3dy5kaWdpY2VydC5jb20vc3NsLWNwcy1yZXBvc2l0b3J5
# Lmh0bTCCAWQGCCsGAQUFBwICMIIBVh6CAVIAQQBuAHkAIAB1AHMAZQAgAG8AZgAg
# AHQAaABpAHMAIABDAGUAcgB0AGkAZgBpAGMAYQB0AGUAIABjAG8AbgBzAHQAaQB0
# AHUAdABlAHMAIABhAGMAYwBlAHAAdABhAG4AYwBlACAAbwBmACAAdABoAGUAIABE
# AGkAZwBpAEMAZQByAHQAIABDAFAALwBDAFAAUwAgAGEAbgBkACAAdABoAGUAIABS
# AGUAbAB5AGkAbgBnACAAUABhAHIAdAB5ACAAQQBnAHIAZQBlAG0AZQBuAHQAIAB3
# AGgAaQBjAGgAIABsAGkAbQBpAHQAIABsAGkAYQBiAGkAbABpAHQAeQAgAGEAbgBk
# ACAAYQByAGUAIABpAG4AYwBvAHIAcABvAHIAYQB0AGUAZAAgAGgAZQByAGUAaQBu
# ACAAYgB5ACAAcgBlAGYAZQByAGUAbgBjAGUALjCBggYIKwYBBQUHAQEEdjB0MCQG
# CCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wTAYIKwYBBQUHMAKG
# QGh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRENv
# ZGVTaWduaW5nQ0EtMS5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQUFAAOC
# AQEAm1zhveo2Zy2lp8UNpR2E2CE8/NvEk0NDLszcBBuMda3N8Du23CikXCgrVvE0
# 3mMaeu/cIMDVU01ityLaqvDuovmTsvAKqaSJNztV9yTeWK9H4+h+35UEIU5TvYLs
# uzEW+rI5M2KcCXR6/LF9ZPmnBf9hHnK44hweHpmDWbo8HPqMatnIo7ideucuDn/D
# BM6s63eTMsFQCPYwte5vxuyVLqodOubLvIOMezZzByrpvJp9+gWAL151CE4qR6xQ
# jpgk5KqSkkkyvl72D+3PhNwZuxZDbZil5PIcrjmaBYoG8wfJzoNrtPFq3aG8dnQr
# xjXJjl+IN1iHYehBAUoBX98EozCCBqMwggWLoAMCAQICEA+oSQYV1wCgviF2/cXs
# bb0wDQYJKoZIhvcNAQEFBQAwZTELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lD
# ZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGln
# aUNlcnQgQXNzdXJlZCBJRCBSb290IENBMB4XDTExMDIxMTEyMDAwMFoXDTI2MDIx
# MDEyMDAwMFowbzELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZ
# MBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEuMCwGA1UEAxMlRGlnaUNlcnQgQXNz
# dXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EtMTCCASIwDQYJKoZIhvcNAQEBBQADggEP
# ADCCAQoCggEBAJx8+aCPCsqJS1OaPOwZIn8My/dIRNA/Im6aT/rO38bTJJH/qFKT
# 53L48UaGlMWrF/R4f8t6vpAmHHxTL+WD57tqBSjMoBcRSxgg87e98tzLuIZARR9P
# +TmY0zvrb2mkXAEusWbpprjcBt6ujWL+RCeCqQPD/uYmC5NJceU4bU7+gFxnd7XV
# b2ZklGu7iElo2NH0fiHB5sUeyeCWuAmV+UuerswxvWpaQqfEBUd9YCvZoV29+1aT
# 7xv8cvnfPjL93SosMkbaXmO80LjLTBA1/FBfrENEfP6ERFC0jCo9dAz0eotyS+BW
# tRO2Y+k/Tkkj5wYW8CWrAfgoQebH1GQ7XasCAwEAAaOCA0MwggM/MA4GA1UdDwEB
# /wQEAwIBhjATBgNVHSUEDDAKBggrBgEFBQcDAzCCAcMGA1UdIASCAbowggG2MIIB
# sgYIYIZIAYb9bAMwggGkMDoGCCsGAQUFBwIBFi5odHRwOi8vd3d3LmRpZ2ljZXJ0
# LmNvbS9zc2wtY3BzLXJlcG9zaXRvcnkuaHRtMIIBZAYIKwYBBQUHAgIwggFWHoIB
# UgBBAG4AeQAgAHUAcwBlACAAbwBmACAAdABoAGkAcwAgAEMAZQByAHQAaQBmAGkA
# YwBhAHQAZQAgAGMAbwBuAHMAdABpAHQAdQB0AGUAcwAgAGEAYwBjAGUAcAB0AGEA
# bgBjAGUAIABvAGYAIAB0AGgAZQAgAEQAaQBnAGkAQwBlAHIAdAAgAEMAUAAvAEMA
# UABTACAAYQBuAGQAIAB0AGgAZQAgAFIAZQBsAHkAaQBuAGcAIABQAGEAcgB0AHkA
# IABBAGcAcgBlAGUAbQBlAG4AdAAgAHcAaABpAGMAaAAgAGwAaQBtAGkAdAAgAGwA
# aQBhAGIAaQBsAGkAdAB5ACAAYQBuAGQAIABhAHIAZQAgAGkAbgBjAG8AcgBwAG8A
# cgBhAHQAZQBkACAAaABlAHIAZQBpAG4AIABiAHkAIAByAGUAZgBlAHIAZQBuAGMA
# ZQAuMBIGA1UdEwEB/wQIMAYBAf8CAQAweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUF
# BzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6
# Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5j
# cnQwgYEGA1UdHwR6MHgwOqA4oDaGNGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9E
# aWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwOqA4oDaGNGh0dHA6Ly9jcmw0LmRp
# Z2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwHQYDVR0OBBYE
# FHtozimqwBe+SXrh5T/Wp/dFjzUyMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6en
# IZ3zbcgPMA0GCSqGSIb3DQEBBQUAA4IBAQB7ch1k/4jIOsG36eepxIe725SS15BZ
# M/orh96oW4AlPxOPm4MbfEPE5ozfOT7DFeyw2jshJXskwXJduEeRgRNG+pw/alE4
# 3rQly/Cr38UoAVR5EEYk0TgPJqFhkE26vSjmP/HEqpv22jVTT8nyPdNs3CPtqqBN
# ZwnzOoA9PPs2TJDndqTd8jq/VjUvokxl6ODU2tHHyJFqLSNPNzsZlBjU1ZwQPNWx
# HBn/j8hrm574rpyZlnjRzZxRFVtCJnJajQpKI5JA6IbeIsKTOtSbaKbfKX8GuTwO
# vZ/EhpyCR0JxMoYJmXIJeUudcWn1Qf9/OXdk8YSNvosesn1oo6WQsQz/MIIGzTCC
# BbWgAwIBAgIQBv35A5YDreoACus/J7u6GzANBgkqhkiG9w0BAQUFADBlMQswCQYD
# VQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGln
# aWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVkIElEIFJvb3QgQ0Ew
# HhcNMDYxMTEwMDAwMDAwWhcNMjExMTEwMDAwMDAwWjBiMQswCQYDVQQGEwJVUzEV
# MBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29t
# MSEwHwYDVQQDExhEaWdpQ2VydCBBc3N1cmVkIElEIENBLTEwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQDogi2Z+crCQpWlgHNAcNKeVlRcqcTSQQaPyTP8
# TUWRXIGf7Syc+BZZ3561JBXCmLm0d0ncicQK2q/LXmvtrbBxMevPOkAMRk2T7It6
# NggDqww0/hhJgv7HxzFIgHweog+SDlDJxofrNj/YMMP/pvf7os1vcyP+rFYFkPAy
# IRaJxnCI+QWXfaPHQ90C6Ds97bFBo+0/vtuVSMTuHrPyvAwrmdDGXRJCgeGDboJz
# PyZLFJCuWWYKxI2+0s4Grq2Eb0iEm09AufFM8q+Y+/bOQF1c9qjxL6/siSLyaxhl
# scFzrdfx2M8eCnRcQrhofrfVdwonVnwPYqQ/MhRglf0HBKIJAgMBAAGjggN6MIID
# djAOBgNVHQ8BAf8EBAMCAYYwOwYDVR0lBDQwMgYIKwYBBQUHAwEGCCsGAQUFBwMC
# BggrBgEFBQcDAwYIKwYBBQUHAwQGCCsGAQUFBwMIMIIB0gYDVR0gBIIByTCCAcUw
# ggG0BgpghkgBhv1sAAEEMIIBpDA6BggrBgEFBQcCARYuaHR0cDovL3d3dy5kaWdp
# Y2VydC5jb20vc3NsLWNwcy1yZXBvc2l0b3J5Lmh0bTCCAWQGCCsGAQUFBwICMIIB
# Vh6CAVIAQQBuAHkAIAB1AHMAZQAgAG8AZgAgAHQAaABpAHMAIABDAGUAcgB0AGkA
# ZgBpAGMAYQB0AGUAIABjAG8AbgBzAHQAaQB0AHUAdABlAHMAIABhAGMAYwBlAHAA
# dABhAG4AYwBlACAAbwBmACAAdABoAGUAIABEAGkAZwBpAEMAZQByAHQAIABDAFAA
# LwBDAFAAUwAgAGEAbgBkACAAdABoAGUAIABSAGUAbAB5AGkAbgBnACAAUABhAHIA
# dAB5ACAAQQBnAHIAZQBlAG0AZQBuAHQAIAB3AGgAaQBjAGgAIABsAGkAbQBpAHQA
# IABsAGkAYQBiAGkAbABpAHQAeQAgAGEAbgBkACAAYQByAGUAIABpAG4AYwBvAHIA
# cABvAHIAYQB0AGUAZAAgAGgAZQByAGUAaQBuACAAYgB5ACAAcgBlAGYAZQByAGUA
# bgBjAGUALjALBglghkgBhv1sAxUwEgYDVR0TAQH/BAgwBgEB/wIBADB5BggrBgEF
# BQcBAQRtMGswJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBD
# BggrBgEFBQcwAoY3aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0
# QXNzdXJlZElEUm9vdENBLmNydDCBgQYDVR0fBHoweDA6oDigNoY0aHR0cDovL2Ny
# bDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDA6oDig
# NoY0aHR0cDovL2NybDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9v
# dENBLmNybDAdBgNVHQ4EFgQUFQASKxOYspkH7R7for5XDStnAs0wHwYDVR0jBBgw
# FoAUReuir/SSy4IxLVGLp6chnfNtyA8wDQYJKoZIhvcNAQEFBQADggEBAEZQPsm3
# KCSnOB22WymvUs9S6TFHq1Zce9UNC0Gz7+x1H3Q48rJcYaKclcNQ5IK5I9G6OoZy
# rTh4rHVdFxc0ckeFlFbR67s2hHfMJKXzBBlVqefj56tizfuLLZDCwNK1lL1eT7EF
# 0g49GqkUW6aGMWKoqDPkmzmnxPXOHXh2lCVz5Cqrz5x2S+1fwksW5EtwTACJHvzF
# ebxMElf+X+EevAJdqP77BzhPDcZdkbkPZ0XN1oPt55INjbFpjE/7WeAjD9KqrgB8
# 7pxCDs+R1ye3Fu4Pw718CqDuLAhVhSK46xgaTfwqIa1JMYNHlXdx3LEbS0scEJx3
# FMGdTy9alQgpECYxggQAMIID/AIBATCBgzBvMQswCQYDVQQGEwJVUzEVMBMGA1UE
# ChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMS4wLAYD
# VQQDEyVEaWdpQ2VydCBBc3N1cmVkIElEIENvZGUgU2lnbmluZyBDQS0xAhAEpVF+
# 1fcA0OvDT46NhL3GMAkGBSsOAwIaBQCgQDAZBgkqhkiG9w0BCQMxDAYKKwYBBAGC
# NwIBBDAjBgkqhkiG9w0BCQQxFgQUPBgaXvKkz3f6eHLoZP4o3+MxIOowDQYJKoZI
# hvcNAQEBBQAEggEAwjbw3qdFlPi7ecxO3IKIIDpLm98C43xaq5Ed0P3ADOqyy6lw
# LnfbumyeqruQleOzKblZebsf3qo/B3Ib/6UHnljbnOKPL7UM5PXUEWVaGfmZHgGo
# T9qhqfFM6hVElb+R6J8w0zawDVHiqU5kWiLLcEGUOn4Lai6AFuVRkyPivMymewjC
# GWEx4Ge/G0xol4aKrM64BBxeSEDDq2p11TsBl4I1/cfLonIh9h3usKR6uePgAI1k
# a4ldY6SI8ZKwEmhgaX0CBRLXkyQklm4eRsqNZYYxY2awyLX+IFLs9uvlP/GN7Q+E
# y0j+5z5ZpbU4D9DPR3Am0jzOMtcqIavFvzc+V6GCAg8wggILBgkqhkiG9w0BCQYx
# ggH8MIIB+AIBATB2MGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJ
# bmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0
# IEFzc3VyZWQgSUQgQ0EtMQIQA5/t7ct5W43tMgyJGfA2iTAJBgUrDgMCGgUAoF0w
# GAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMTQwMjAx
# MTYzMjMzWjAjBgkqhkiG9w0BCQQxFgQUYjMk+yTxbmNR0pEA/gB6qPlb3kUwDQYJ
# KoZIhvcNAQEBBQAEggEAISYIEj5CqsfjcLpyeGjHDVMr/5gFmES3S+01EyPJJ+X2
# aGiecvy0Y62xC4ZqRawxxwB48QuUdGQqYnuJyUl5jj79pQlV/vSJbl28ifA6HCHe
# sG+EZxzTIc9zzHyLIL8tMztFoNHA3CnaUD4dOAL7apZj+Z/XUa+TzBaKqRuyayMt
# uflWD+FQFyDkmbLMB/QqwPGLjYzHUCR4WMWD/10yOIVID4kluxKpTPGwFnRNT9Ql
# 1FtZXRGgYPpHFbU7b4upeNCf64ndgTQoD7NwNpZEKaFKa00K2qNsRd9xD5gbW2OY
# KXYdnZsKuu+ke9XGI6FuGe24MU/IMDxwfwV0IiNgcA==
# SIG # End signature block
