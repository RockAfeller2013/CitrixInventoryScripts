<#
.SYNOPSIS
	Gets the average, minimum and maximum Group Policy processing time on XenApp 6.5 servers.
.DESCRIPTION
	Builds a list of all XenApp 6.5 servers in a Farm.
	Process each server looking in the Microsoft-Windows-GroupPolicy/Operational for all Event ID 8001.
	Displays the Avergage, Minimum and Maximum processing times.
	
	All events where processing time is 0 are ignored. A 0 time means a local account was used for login.
	
	There is a bug with Get-WinEvent and PowerShell versions later than 2 or culture other than en-US,
	the Message property is not returned.

	There are two work-arounds:
	1. PowerShell.exe -Version 2
	2. Add this line to the script: 
	[System.Threading.Thread]::CurrentThread.CurrentCulture = New-Object "System.Globalization.CultureInfo" "en-US"
	
.PARAMETER MaxSeconds
	Specifies the number of seconds to use for the cutoff for GPO processing time.
	Any value greater than or equal to MaxSeconds is recorded along with the user name and server name.
	Default is 30.
.PARAMETER Folder
	Specifies the optional output folder to save the output reports. 
.EXAMPLE
	PS C:\PSScript > .\Get-GPOTimeV1.1.ps1
.EXAMPLE
	PS C:\PSScript > .\Get-GPOTimeV1.1.ps1 -Folder \\ServerName\Share
	
	Saves the two output text files in \\ServerName\Share.
.EXAMPLE
	PS C:\PSScript > .\Get-GPOTimeV1.1.ps1 -MaxSeconds 10
	
	When the total group policy processing time is greater than or equal 10 seconds,
	the time, user name and server name are recorded in LongGPOTimes.txt.
.EXAMPLE
	PS C:\PSScript > .\Get-GPOTimeV1.1.ps1 -MaxSeconds 17 -Folder c:\LogFiles
	
	When the total group policy processing time is greater than or equal 17 seconds,
	the time, user name and server name are recorded in LongGPOTimes.txt.
	
	Saves the two output text files in C:\LogFiles.
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.
	The script creates two text files:
		LongGPOTimes.txt
		GPOAvgMinMaxTimes.txt
		
	By default, the two text files are stored in the folder where the script is run.
.NOTES
	NAME: Get-GPOTime.ps1
	VERSION: 1.1
	AUTHOR: Carl Webster
	LASTEDIT: March 24, 2016
#>


#Created by Carl Webster, CTP and independent consultant 05-Mar-2016
#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com

[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "Default") ]

Param(
	[parameter(ParameterSetName="Default",Mandatory=$False)] 
	[Int]$MaxSeconds = 30,

	[parameter(ParameterSetName="Default",Mandatory=$False)] 
	[string]$Folder=""
	
	)

#Version 1.1 24-Mar-2016
#	Allows you to specify the maximum number of seconds group policy processing should take. Any number greater than or equal to that number is recorded in LongGPOTimes.txt.
#	Allows you to specify an output folder.
#	Records the long GPO times in an text file.
#	Records the Average, Minimum and Maximum processing time to GPOAvgMinMaxTimes.txt.
#	GPOAvgMinMaxTimes.txt is a cumulative file and records the Average, Minimum and Maximum times for each run of the script.

Write-Host "$(Get-Date): Setting up script"

If($MaxSeconds -eq $Null)
{
	$MaxSeconds = 30
}
If($Folder -eq $Null)
{
	$Folder = ""
}

If(!(Test-Path Variable:Seconds))
{
	$MaxSeconds = 30
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

[string]$FileName1 = "$($pwdpath)\LongGPOTimes.txt"
[string]$FileName2 = "$($pwdpath)\GPOAvgMinMaxTimes.txt"

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
	Write-Host "$(Get-Date): Remoting is not used"
	
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
$servers = Get-XAServer -ea 0 | Select ServerName | Sort ServerName

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
	$TimeArray = @()
	$LongGPOsArray = @()
	$cnt = 0
	ForEach($server in $servers)
	{
		$cnt++
		Write-Host "$(Get-Date): Processing server $($Server.ServerName) $($Total - $cnt) left"
		If(Test-Connection -ComputerName $server.servername -quiet -EA 0)
		{
			try
			{
				$GPTime = Get-WinEvent -logname Microsoft-Windows-GroupPolicy/Operational `
				-computername $server.servername | Where {$_.id -eq "8001"} | Select message
			}
			
			catch
			{
				Write-Host "$(Get-Date): `tServer $($Server.ServerName) had error being accessed"
				Continue
			}
			
			If($? -and $Null -ne $GPTime)
			{
				ForEach($GPT in $GPTime)
				{
					$tmparray = $GPT.Message.ToString().Split(" ")
					[int]$GPOTime = $tmparray[8]
					If($GPOTime -ne 0)
					{
						$TimeArray += $GPOTime
					}
					If($GPOTime -ge $MaxSeconds)
					{
						$obj = New-Object -TypeName PSObject
						$obj | Add-Member -MemberType NoteProperty -Name MaxSeconds	-Value $GPOTime
						$obj | Add-Member -MemberType NoteProperty -Name User		-Value $tmparray[6]
						$obj | Add-Member -MemberType NoteProperty -Name Server		-Value $server.servername
						$LongGPOsArray += $obj
					}
					
				}
			}
		}
		Else
		{
			Write-Host "$(Get-Date): `tServer $($Server.ServerName) is not online"
		}
	}
	
	Write-Host "$(Get-Date): Output long GPO times to file"
	#first sort array by seconds, longest to shortest
	$LongGPOsArray = $LongGPOsArray | Sort MaxSeconds -Descending
	Out-File -FilePath $Filename1 -InputObject $LongGPOsArray

	If(Test-Path "$($FileName1)")
	{
		Write-Host "$(Get-Date): $($FileName1) is ready for use"
	}

	$Avg = ($TimeArray | Measure-Object -Average -minimum -maximum)
	Write-Host "Average: " $Avg.Average
	Write-Host "Minimum: " $Avg.Minimum
	Write-Host "Maximum: " $Avg.Maximum

	Write-Host "$(Get-Date): Output GPO Avg/Min/Max times to file"
	Out-File -FilePath $Filename2 -Append -InputObject " "
	Out-File -FilePath $Filename2 -Append -InputObject "$(Get-Date): Average: $($Avg.Average) seconds"
	Out-File -FilePath $Filename2 -Append -InputObject "$(Get-Date): Minimum: $($Avg.Minimum) seconds"
	Out-File -FilePath $Filename2 -Append -InputObject "$(Get-Date): Maximum: $($Avg.Maximum) seconds"
	Out-File -FilePath $Filename2 -Append -InputObject " "

	If(Test-Path "$($FileName2)")
	{
		Write-Host "$(Get-Date): $($FileName2) is ready for use"
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

# SIG # Begin signature block
# MIIgCgYJKoZIhvcNAQcCoIIf+zCCH/cCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUP+qT91wXSk0F0lQDCysTf0/n
# wPKgghtxMIIDtzCCAp+gAwIBAgIQDOfg5RfYRv6P5WD8G/AwOTANBgkqhkiG9w0B
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
# /AeB+ova+YJJ92JuoVP6EpQYhS6SkepobEQysmah5xikmmRR7zCCBT8wggQnoAMC
# AQICEAmkTdj/HQvKi5Whef7gyA8wDQYJKoZIhvcNAQELBQAwcjELMAkGA1UEBhMC
# VVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0
# LmNvbTExMC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIENvZGUgU2ln
# bmluZyBDQTAeFw0xNTEwMjkwMDAwMDBaFw0xNjExMDIxMjAwMDBaMHwxCzAJBgNV
# BAYTAlVTMQswCQYDVQQIEwJUTjESMBAGA1UEBxMJVHVsbGFob21hMSUwIwYDVQQK
# ExxDYXJsIFdlYnN0ZXIgQ29uc3VsdGluZywgTExDMSUwIwYDVQQDExxDYXJsIFdl
# YnN0ZXIgQ29uc3VsdGluZywgTExDMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIB
# CgKCAQEA35g9yG7Fh7/h1rbQmW2x6BmWEWCBw6qwOKfXDJyMMeSunAKZ+rnBYX3K
# T1ERQYMYi2/tK1/hNcgW3ja6sSqwEWBde/nLmqdkzMJb2pUPGUhVP0ZMO7KCS8oz
# Ed5FPpT4Hete/8OQyGKTdU16Ne2xhWzgVvKP1g0zLXJojIWYB4+kKOY2OCl8oPhX
# LwMlQEraFUz39JDkwumteT2/MEjORclAAJ+odAk9R1jjOD5p5GzLRi27vDrBUDq2
# wNsHgejZrq4mbyLiNqdZnFKUeQCzCF8YF32U9E0O+fdhY4QvTM2Jdtusz1d/IIz/
# JqM2AjkDkEXUMK6nQ3015j9yoOAQiQIDAQABo4IBxTCCAcEwHwYDVR0jBBgwFoAU
# WsS5eyoKo6XqcQPAYPkt9mV1DlgwHQYDVR0OBBYEFLdZN8kA2rYz8RkS85RNuO4I
# GxMHMA4GA1UdDwEB/wQEAwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzB3BgNVHR8E
# cDBuMDWgM6Axhi9odHRwOi8vY3JsMy5kaWdpY2VydC5jb20vc2hhMi1hc3N1cmVk
# LWNzLWcxLmNybDA1oDOgMYYvaHR0cDovL2NybDQuZGlnaWNlcnQuY29tL3NoYTIt
# YXNzdXJlZC1jcy1nMS5jcmwwTAYDVR0gBEUwQzA3BglghkgBhv1sAwEwKjAoBggr
# BgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAIBgZngQwBBAEw
# gYQGCCsGAQUFBwEBBHgwdjAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNl
# cnQuY29tME4GCCsGAQUFBzAChkJodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20v
# RGlnaUNlcnRTSEEyQXNzdXJlZElEQ29kZVNpZ25pbmdDQS5jcnQwDAYDVR0TAQH/
# BAIwADANBgkqhkiG9w0BAQsFAAOCAQEAGz9cEmjU3FosI30XHF355vqavCPByB2F
# TYvGpToMODFnVKul0dQjbF9CWWNeuknYfVjmYBKOgBaFkF/eAy4yfk41tmZZnN9D
# j4Ngenvbrx7ZJqC/ZMNgoIM7un1WLrqZKS5tOaFpBwaEeAIzfU9dHHE27zchIoAJ
# x5aDQbnP6SVWitxa/jGa78b9pDslLpv7Pm4KAEv5d2NYiQ7nhvHShFnWY6wMNBTE
# i+q5rSNcm4TzYsyYSoYT+bGs21vvSAlMSKlvsL0oMWLHMdsMKtC+1Wp2sE4Fshdt
# 9K8DBkl33XhdprC2KabgZa6GTz5NA/rV4FW6oDUidts19XbWIjlB7DCCBmowggVS
# oAMCAQICEAMBmgI6/1ixa9bV6uYX8GYwDQYJKoZIhvcNAQEFBQAwYjELMAkGA1UE
# BhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2lj
# ZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgQXNzdXJlZCBJRCBDQS0xMB4XDTE0
# MTAyMjAwMDAwMFoXDTI0MTAyMjAwMDAwMFowRzELMAkGA1UEBhMCVVMxETAPBgNV
# BAoTCERpZ2lDZXJ0MSUwIwYDVQQDExxEaWdpQ2VydCBUaW1lc3RhbXAgUmVzcG9u
# ZGVyMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAo2Rd/Hyz4II14OD2
# xirmSXU7zG7gU6mfH2RZ5nxrf2uMnVX4kuOe1VpjWwJJUNmDzm9m7t3LhelfpfnU
# h3SIRDsZyeX1kZ/GFDmsJOqoSyyRicxeKPRktlC39RKzc5YKZ6O+YZ+u8/0SeHUO
# plsU/UUjjoZEVX0YhgWMVYd5SEb3yg6Np95OX+Koti1ZAmGIYXIYaLm4fO7m5zQv
# MXeBMB+7NgGN7yfj95rwTDFkjePr+hmHqH7P7IwMNlt6wXq4eMfJBi5GEMiN6ARg
# 27xzdPpO2P6qQPGyznBGg+naQKFZOtkVCVeZVjCT88lhzNAIzGvsYkKRrALA76Tw
# iRGPdwIDAQABo4IDNTCCAzEwDgYDVR0PAQH/BAQDAgeAMAwGA1UdEwEB/wQCMAAw
# FgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwggG/BgNVHSAEggG2MIIBsjCCAaEGCWCG
# SAGG/WwHATCCAZIwKAYIKwYBBQUHAgEWHGh0dHBzOi8vd3d3LmRpZ2ljZXJ0LmNv
# bS9DUFMwggFkBggrBgEFBQcCAjCCAVYeggFSAEEAbgB5ACAAdQBzAGUAIABvAGYA
# IAB0AGgAaQBzACAAQwBlAHIAdABpAGYAaQBjAGEAdABlACAAYwBvAG4AcwB0AGkA
# dAB1AHQAZQBzACAAYQBjAGMAZQBwAHQAYQBuAGMAZQAgAG8AZgAgAHQAaABlACAA
# RABpAGcAaQBDAGUAcgB0ACAAQwBQAC8AQwBQAFMAIABhAG4AZAAgAHQAaABlACAA
# UgBlAGwAeQBpAG4AZwAgAFAAYQByAHQAeQAgAEEAZwByAGUAZQBtAGUAbgB0ACAA
# dwBoAGkAYwBoACAAbABpAG0AaQB0ACAAbABpAGEAYgBpAGwAaQB0AHkAIABhAG4A
# ZAAgAGEAcgBlACAAaQBuAGMAbwByAHAAbwByAGEAdABlAGQAIABoAGUAcgBlAGkA
# bgAgAGIAeQAgAHIAZQBmAGUAcgBlAG4AYwBlAC4wCwYJYIZIAYb9bAMVMB8GA1Ud
# IwQYMBaAFBUAEisTmLKZB+0e36K+Vw0rZwLNMB0GA1UdDgQWBBRhWk0ktkkynUoq
# eRqDS/QeicHKfTB9BgNVHR8EdjB0MDigNqA0hjJodHRwOi8vY3JsMy5kaWdpY2Vy
# dC5jb20vRGlnaUNlcnRBc3N1cmVkSURDQS0xLmNybDA4oDagNIYyaHR0cDovL2Ny
# bDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEQ0EtMS5jcmwwdwYIKwYB
# BQUHAQEEazBpMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20w
# QQYIKwYBBQUHMAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2Vy
# dEFzc3VyZWRJRENBLTEuY3J0MA0GCSqGSIb3DQEBBQUAA4IBAQCdJX4bM02yJoFc
# m4bOIyAPgIfliP//sdRqLDHtOhcZcRfNqRu8WhY5AJ3jbITkWkD73gYBjDf6m7Gd
# JH7+IKRXrVu3mrBgJuppVyFdNC8fcbCDlBkFazWQEKB7l8f2P+fiEUGmvWLZ8Cc9
# OB0obzpSCfDscGLTYkuw4HOmksDTjjHYL+NtFxMG7uQDthSr849Dp3GdId0UyhVd
# kkHa+Q+B0Zl0DSbEDn8btfWg8cZ3BigV6diT5VUW8LsKqxzbXEgnZsijiwoc5ZXa
# rsQuWaBh3drzbaJh6YoLbewSGL33VVRAA5Ira8JRwgpIr7DUbuD0FAo6G+OPPcqv
# ao173NhEMIIGzTCCBbWgAwIBAgIQBv35A5YDreoACus/J7u6GzANBgkqhkiG9w0B
# AQUFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVk
# IElEIFJvb3QgQ0EwHhcNMDYxMTEwMDAwMDAwWhcNMjExMTEwMDAwMDAwWjBiMQsw
# CQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cu
# ZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBBc3N1cmVkIElEIENBLTEw
# ggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDogi2Z+crCQpWlgHNAcNKe
# VlRcqcTSQQaPyTP8TUWRXIGf7Syc+BZZ3561JBXCmLm0d0ncicQK2q/LXmvtrbBx
# MevPOkAMRk2T7It6NggDqww0/hhJgv7HxzFIgHweog+SDlDJxofrNj/YMMP/pvf7
# os1vcyP+rFYFkPAyIRaJxnCI+QWXfaPHQ90C6Ds97bFBo+0/vtuVSMTuHrPyvAwr
# mdDGXRJCgeGDboJzPyZLFJCuWWYKxI2+0s4Grq2Eb0iEm09AufFM8q+Y+/bOQF1c
# 9qjxL6/siSLyaxhlscFzrdfx2M8eCnRcQrhofrfVdwonVnwPYqQ/MhRglf0HBKIJ
# AgMBAAGjggN6MIIDdjAOBgNVHQ8BAf8EBAMCAYYwOwYDVR0lBDQwMgYIKwYBBQUH
# AwEGCCsGAQUFBwMCBggrBgEFBQcDAwYIKwYBBQUHAwQGCCsGAQUFBwMIMIIB0gYD
# VR0gBIIByTCCAcUwggG0BgpghkgBhv1sAAEEMIIBpDA6BggrBgEFBQcCARYuaHR0
# cDovL3d3dy5kaWdpY2VydC5jb20vc3NsLWNwcy1yZXBvc2l0b3J5Lmh0bTCCAWQG
# CCsGAQUFBwICMIIBVh6CAVIAQQBuAHkAIAB1AHMAZQAgAG8AZgAgAHQAaABpAHMA
# IABDAGUAcgB0AGkAZgBpAGMAYQB0AGUAIABjAG8AbgBzAHQAaQB0AHUAdABlAHMA
# IABhAGMAYwBlAHAAdABhAG4AYwBlACAAbwBmACAAdABoAGUAIABEAGkAZwBpAEMA
# ZQByAHQAIABDAFAALwBDAFAAUwAgAGEAbgBkACAAdABoAGUAIABSAGUAbAB5AGkA
# bgBnACAAUABhAHIAdAB5ACAAQQBnAHIAZQBlAG0AZQBuAHQAIAB3AGgAaQBjAGgA
# IABsAGkAbQBpAHQAIABsAGkAYQBiAGkAbABpAHQAeQAgAGEAbgBkACAAYQByAGUA
# IABpAG4AYwBvAHIAcABvAHIAYQB0AGUAZAAgAGgAZQByAGUAaQBuACAAYgB5ACAA
# cgBlAGYAZQByAGUAbgBjAGUALjALBglghkgBhv1sAxUwEgYDVR0TAQH/BAgwBgEB
# /wIBADB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRp
# Z2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2NhY2VydHMuZGlnaWNlcnQu
# Y29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDCBgQYDVR0fBHoweDA6oDig
# NoY0aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9v
# dENBLmNybDA6oDigNoY0aHR0cDovL2NybDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0
# QXNzdXJlZElEUm9vdENBLmNybDAdBgNVHQ4EFgQUFQASKxOYspkH7R7for5XDStn
# As0wHwYDVR0jBBgwFoAUReuir/SSy4IxLVGLp6chnfNtyA8wDQYJKoZIhvcNAQEF
# BQADggEBAEZQPsm3KCSnOB22WymvUs9S6TFHq1Zce9UNC0Gz7+x1H3Q48rJcYaKc
# lcNQ5IK5I9G6OoZyrTh4rHVdFxc0ckeFlFbR67s2hHfMJKXzBBlVqefj56tizfuL
# LZDCwNK1lL1eT7EF0g49GqkUW6aGMWKoqDPkmzmnxPXOHXh2lCVz5Cqrz5x2S+1f
# wksW5EtwTACJHvzFebxMElf+X+EevAJdqP77BzhPDcZdkbkPZ0XN1oPt55INjbFp
# jE/7WeAjD9KqrgB87pxCDs+R1ye3Fu4Pw718CqDuLAhVhSK46xgaTfwqIa1JMYNH
# lXdx3LEbS0scEJx3FMGdTy9alQgpECYxggQDMIID/wIBATCBhjByMQswCQYDVQQG
# EwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNl
# cnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQgQ29kZSBT
# aWduaW5nIENBAhAJpE3Y/x0LyouVoXn+4MgPMAkGBSsOAwIaBQCgQDAZBgkqhkiG
# 9w0BCQMxDAYKKwYBBAGCNwIBBDAjBgkqhkiG9w0BCQQxFgQUY4gwPdOlNe4j3DcD
# 8WywCJmMLrEwDQYJKoZIhvcNAQEBBQAEggEAlFogVz+3lkhYl0i4mlbkZsKa3OiD
# 9sXPSlzIU0CN4H783oF2+cA1LkFQQnTVllsKavNdJRDMeRNFMAOV4vp257fWq0Pl
# FeXWcwT+R94RchG8dqhWpNs5SzQOQYw9atWYprFeQMLtrYz4PI9n4c0bQ14PIBAN
# 4p4w2ZmejjH5GMgUiDn1wusDOTVu5uYSvp3bP5NfwWFoIlhd8b1ytqORA00cCPQI
# kNI/89t4ttDst44BWllljcP4H2cnY/9QgwcxplWrS5f5K03Hk8caqWbjqBOctjkN
# Q3PhQGvVDX5nn8aswMran1aaF6KJTDfeRDHr5UMgeKGUeaMe+0SQgsKgDKGCAg8w
# ggILBgkqhkiG9w0BCQYxggH8MIIB+AIBATB2MGIxCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAf
# BgNVBAMTGERpZ2lDZXJ0IEFzc3VyZWQgSUQgQ0EtMQIQAwGaAjr/WLFr1tXq5hfw
# ZjAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG
# 9w0BCQUxDxcNMTYwNzI0MDEyMTQxWjAjBgkqhkiG9w0BCQQxFgQUmYp16JdSgVom
# h5qoxgRTSrlmJmYwDQYJKoZIhvcNAQEBBQAEggEAIPcclkHl+EE5hHj1jEtPR9ky
# V+kLAD6sBOX5PMMSLVM/GbCrC+ZZd6EJTyI4rLKoUHrqTpWBtglaHilgWnTdlfgn
# 85a0ZrOV0WnTzs2qm0VWHvR8INBNqbRZF9u8jXwWrXfuF4tCv4JjJXiNS3k2nqUE
# Bk8IqhJ06Ptb7dX470GAO1IZbQPlBSXMlimpJZcxAoOlvgE+ZOnFpF3MxzgcQXRz
# hdOMbHzctHOip6o0ZRq5pw7m8nVvP7pN0iZ2Yl0Mu6K1T8KT98uC2ywSnkjflDJf
# MmdUZkvccR6LYW/sHi/6w2Nsc7THuN2jJvk4sOagrIohdE6ow2OTun/x94xYRQ==
# SIG # End signature block
