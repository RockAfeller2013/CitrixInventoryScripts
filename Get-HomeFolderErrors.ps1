<#
.SYNOPSIS
	Find Home Folder errors on XenApp 6.5 servers.
.DESCRIPTION
	Builds a list of all XenApp 6.5 servers in a Farm.
	Process each server looking for Home Folder errors (Event ID 1060, Source TermService) 
	within, by default, the last 30 days.
	Builds a list of unique user names and servers unable to process.
	Creates the two text files, by default, in the folder where the script is run.
	Optionally, can specify the output folder.
.PARAMETER StartDate
	Start date, in MM/DD/YYYY format.
	Default is today's date minus 30 days.
.PARAMETER EndDate
	End date, in MM/DD/YYYY HH:MM format.
	Default is today's date.
.PARAMETER Folder
	Specifies the optional output folder to save the output reports. 
.EXAMPLE
	PS C:\PSScript > .\Get-HomeFolderErrors.ps1
	
.EXAMPLE
	PS C:\PSScript > .\Get-HomeFolderErrors.ps1 -StartDate "08/01/2016" -EndDate "08/02/2016" 
	
	Will return all Folder Redirection errors from "08/01/2016" through "08/02/2016".
.EXAMPLE
	PS C:\PSScript > .\Get-HomeFolderErrors.ps1 -StartDate "08/01/2016" -EndDate "08/01/2016" 
	
	Will return all Configuration Logging entries from "08/01/2016" through "08/01/2016".
.EXAMPLE
	PS C:\PSScript > .\Get-HomeFolderErrors.ps1 -Folder \\FileServer\ShareName
	
	Output files will be saved in the path \\FileServer\ShareName
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  This script creates two text files.
.NOTES
	NAME: Get-HomeFolderErrors.ps1
	VERSION: 1.00
	AUTHOR: Carl Webster, Sr. Solutions Architect for Choice Solutions, LLC
	LASTEDIT: August 26, 2016
#>


#Created by Carl Webster, CTP 24-Aug-2016
#Sr. Solutions Architect for Choice Solutions, LLC
#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com
#
#Version 1.00 released to the community on 26-Aug-2016

[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "Default") ]

Param(
	[parameter(ParameterSetName="Default",Mandatory=$False)] 
	[Datetime]$StartDate = ((Get-Date -displayhint date).AddDays(-30)),

	[parameter(ParameterSetName="Default",Mandatory=$False)] 
	[Datetime]$EndDate = (Get-Date -displayhint date),
	
	[parameter(ParameterSetName="Default",Mandatory=$False)] 
	[string]$Folder=""
	
	)

Write-Host "$(Get-Date): Setting up script"

If($StartDate -eq $Null)
{
	$StartDate = ((Get-Date -displayhint date).AddDays(-30))
}
If($EndDate -eq $Null)
{
	$EndDate = (Get-Date -displayhint date)
}
If($Folder -eq $Null)
{
	$Folder = ""
}

If(!(Test-Path Variable:StartDate))
{
	$StartDate = ((Get-Date -displayhint date).AddDays(-30))
}
If(!(Test-Path Variable:EndDate))
{
	$EndDate = ((Get-Date -displayhint date))
}
If(!(Test-Path Variable:Folder))
{
	$Folder = ""
}

If($Folder -ne "")
{
	Write-Host "$(Get-Date): Testing folder path"
	#does it exist
	If(Test-Path $Folder -EA 0)
	{
		#it exists, now check to see if it is a folder and not a file
		If(Test-Path $Folder -pathType Container -EA 0)
		{
			#it exists and it is a folder
			Write-Host "$(Get-Date): Folder path $Folder exists and is a folder"
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

[string]$FileName1 = "$($pwdpath)\HomeFolderErrors.txt"
[string]$FileName2 = "$($pwdpath)\OfflineServers.txt"

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
				Write-Host "$(Get-Date): Loading Windows PowerShell snap-in: $snapin"
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

Write-Host "$(Get-Date): Loading XenApp snapin"
If(!(Check-NeededPSSnapins "Citrix.XenApp.Commands"))
{
	#We're missing Citrix Snapins that we need
	$ErrorActionPreference = $SaveEAPreference
	Write-Error "Missing Citrix PowerShell Snap-ins Detected, check the console above for more information. Script will now close."
	Exit
}

[bool]$Remoting = $False
$RemoteXAServer = Get-XADefaultComputerName -EA 0 
If(![String]::IsNullOrEmpty($RemoteXAServer))
{
	$Remoting = $True
}

If($Remoting)
{
	Write-Host "$(Get-Date): Remoting is enabled to XenApp server $RemoteXAServer"
	#now need to make sure the script is not being run against a session-only host
	$Server = Get-XAServer -ServerName $RemoteXAServer -EA 0 
	If($Server.ElectionPreference -eq "WorkerMode")
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Warning "This script cannot be run remotely against a Session-only Host Server."
		Write-Warning "Use Set-XADefaultComputerName XA65ControllerServerName or run the script on a controller."
		Write-Error "Script cannot continue.  See messages above."
		Exit
	}
}
Else
{
	Write-Host "$(Get-Date): Remoting is not being used"
	
	#now need to make sure the script is not being run on a session-only host
	$ServerName = (Get-Childitem env:computername).value
	$Server = Get-XAServer -ServerName $ServerName -EA 0
	If($Server.ElectionPreference -eq "WorkerMode")
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Warning "This script cannot be run on a Session-only Host Server if Remoting is not enabled."
		Write-Warning "Use Set-XADefaultComputerName XA65ControllerServerName or run the script on a controller."
		Write-Error "Script cannot continue.  See messages above."
		Exit
	}
}

$startTime = Get-Date

Write-Host "$(Get-Date): Getting XenApp servers"
$servers = Get-XAServer -ea 0 | select ServerName | Sort ServerName

If($? -and $Null -ne $servers)
{
	If($servers -is [Array])
	{
		[int]$Total = $servers.count
	}
	Else
	{
		[int]$Total = 1
	}
	Write-Host "$(Get-Date): Found $($Total) XenApp servers"
	$ErrorArray = @()
	$cnt = 0
	ForEach($server in $servers)
	{
		$cnt++
		Write-Host "$(Get-Date): Processing server $($Server.ServerName) $($Total - $cnt) left"
		If(Test-Connection -ComputerName $server.servername -quiet -EA 0)
		{
			try
			{
				$Errors = Get-EventLog -logname system -computername $Server.ServerName -source "TermService" `
				-entrytype "Error" -after $StartDate.ToShortDateString() -before $EndDate.ToShortDateString() -EA 0
			}
			
			catch
			{
				Write-Host "$(Get-Date): `tServer $($Server.ServerName) had error being accessed"
				Out-File -FilePath $Filename2 -Append -InputObject "Server $($Server.ServerName) had error being accessed $(Get-Date)"
				Continue
			}
			
			If($? -and $Null -ne $Errors)
			{
				$data = @()
				
				ForEach($Item in $Errors)
				{
					If($Item.Message -like "*home directory*")
					{
						$obj = New-Object -TypeName PSObject
						$obj | Add-Member -MemberType NoteProperty -Name UserName -Value $Item.ReplacementStrings[0]
						$obj | Add-Member -MemberType NoteProperty -Name DomainName -Value $Item.ReplacementStrings[1]
						
						$data += $obj
					}
				}	

				$Errors = $data | Sort DomainName, UserName -Unique
				
				$ErrorCount = 0
				If($Errors -is [Array])
				{
					$ErrorCount = $Errors.Count
				}
				Else
				{
					$ErrorCount = 1
					[array]$Errors = $Errors
				}
				
				Write-Host "$(Get-Date): `t$($ErrorCount) Home Folder errors found"
				$ErrorArray += $Errors
				[array]$ErrorArray = $ErrorArray | Sort UserName -Unique
				
				$ErrorArrayCount = 0
				If($ErrorArray -is [Array])
				{
					$ErrorArrayCount = $ErrorArray.Count
				}
				Else
				{
					$ErrorArrayCount = 1
				}
				Write-Host "$(Get-Date): `t`t$($ErrorArrayCount) total Home Folder errors found"
			}
			Else
			{
				Write-Host "$(Get-Date): `tNo Home Folder errors found"
			}
		}
		Else
		{
			Write-Host "$(Get-Date): `tServer $($Server.ServerName) is not online"
			Out-File -FilePath $Filename2 -Append -InputObject "Server $($Server.ServerName) was not online $(Get-Date)"
		}
	}
	$ErrorArray = $ErrorArray | Sort UserName -Unique
	
	Write-Host "$(Get-Date): Output Home Folder errors to file"
	Out-File -FilePath $Filename1 -InputObject $ErrorArray

	If(Test-Path "$($FileName2)")
	{
		Write-Host "$(Get-Date): $($FileName2) is ready for use"
	}
	If(Test-Path "$($FileName1)")
	{
		Write-Host "$(Get-Date): $($FileName1) is ready for use"
	}
}
ElseIf($? -and $Null -eq $servers)
{
	Write-Warning "Server information could not be retrieved"
}
Else
{
	Write-Warning "No results returned for Server information"
}

Write-Host "$(Get-Date): Script started: $($StartTime)"
Write-Host "$(Get-Date): Script ended: $(Get-Date)"
$runtime = $(Get-Date) - $StartTime
$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds", `
	$runtime.Days, `
	$runtime.Hours, `
	$runtime.Minutes, `
	$runtime.Seconds,
	$runtime.Milliseconds)
Write-Host "$(Get-Date): Elapsed time: $($Str)"
$runtime = $Null
