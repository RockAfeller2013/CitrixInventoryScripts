<#
.SYNOPSIS
	Creates an XML file of a StoreFront server cluster
	
.DESCRIPTION
	First part of a 2-script process to document a Citrix StoreFront server (group) installation.
	Uses the StoreFront PowerShell cmdlets to create an XML file of the StoreFront configuration.
	This XML file can then be processed on any machine running the corresponding StoreFront client 
	documentation script (SFClient.ps1).
	Look for the section called "GUI Customizations" to find the lines that you may wish to modify  
	for your script needs.

.PARAMETER OutputDirectory
	Alias: OUTDIR
	Output directory for the resulting XML file
	Default: Current directory.
	
.PARAMETER OutputFile
	Alias: OUT
	Output filename (no extension) for the resulting XML file. The .xml extension will be added.
	Default: StoreFront

.PARAMETER GUI
	Use a graphical form to accept parameters from the user. 
	The GUI will accept parameters passed on the command line.
	Default: True. Use -GUI:$False to turn off.
	
.PARAMETER Hardware
	Use WMI to gather hardware information on: Computer System, Disks, Processor and Network Interface Cards
	This parameter may require the script be run from an elevated PowerShell session 
	using an account with permission to retrieve hardware information (i.e. Domain Admin or Local Administrator).
	Selecting this parameter will add to both the time it takes to run the script and size of the report.
	Default: False

.PARAMETER Software
	Read the registry to obtain a list of installed software, and use WMI to get Citrix services installed 
	on the host machine (and their current states).
	This parameter may require the script be run from an elevated PowerShell session 
	using an account with permission to read the registry (i.e. Domain Admin or Local Administrator).
	Selecting this parameter will add to both the time it takes to run the script and size of the report.
	Default: False
	
.EXAMPLE
	PS C:\PSScript > .\SFServer.ps1
	
	Will use all default values (Use the GUI prepopulated with the current directory
	and the filename StoreFront).
	
.EXAMPLE
	PS C:\PSScript > .\SFServer.ps1 -GUI:$False -OUT "IPM-SF" -OUTDIR "C:\Output" -Hardware
	
	Will gather hardware information for the host server, and
	will create the file IPM-SF.xml in directory C:\Output without using the GUI.
	
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  
	This script creates an XML document.
.NOTES
	NAME:     SFServer.ps1
	VERSION:  1.00
	AUTHOR:   Sam Jacobs
	LASTEDIT: July 20, 2015
#>

Param(

	[parameter(Mandatory=$False)] 
	[Alias("OUTDIR")]
	[ValidateNotNullOrEmpty()]
	[string]$OutputDir=$pwd,
    
	[parameter(Mandatory=$False)] 
	[Alias("OUT")]
	[ValidateNotNullOrEmpty()]
	[string]$OutputFile="StoreFront",
	
	[parameter(Mandatory=$False)] 
	[Switch]$GUI=$True,

	[parameter(Mandatory=$False)] 
	[Switch]$Hardware=$False,

	[parameter(Mandatory=$False)] 
	[Switch]$Software=$False

	)

Function LoadStorefrontSnapins()
{
    write-host "Adding snapins"

    # Do not attempt to load these snapins, they are loaded
    # by the modules too and cause a clash
    $excludedSnapins = @("Citrix.DeliveryServices.ConfigurationProvider","Citrix.DeliveryServices.ClusteredCredentialWallet.Install","Citrix.DeliveryServices.Workflow.WCF.Install" )

    $availableSnapins = Get-PSSnapin -Name "Citrix.DeliveryServices.*" -Registered | Select -ExpandProperty "Name"
    $loadedSnapins = Get-PSSnapin -Name "Citrix.DeliveryServices.*" -ErrorAction SilentlyContinue | Select -ExpandProperty "Name"

    foreach ($snapin in $availableSnapins)
    {
        if (($excludedSnapins -notcontains $snapin) -and ($loadedSnapins -notcontains $snapin))
        {
            Add-PSSnapin -Name $snapin
        }
    }
}

Function ImportStorefrontModules()
{
    write-host "Importing modules"
    $dsInstallProp = Get-ItemProperty -Path HKLM:\SOFTWARE\Citrix\DeliveryServicesManagement -Name InstallDir
    $dsInstallDir = $dsInstallProp.InstallDir

    $dsModules = Get-ChildItem -Path "$dsInstallDir\Cmdlets" | Where { $_.FullName.EndsWith('psm1') } | foreach { $_.FullName }


    foreach ($dsModule in $dsModules)
    {
            Import-Module $dsModule
    }
}

Set-StrictMode -Version 2

#force -verbose on
$PSDefaultParameterValues = @{"*:Verbose"=$True}

Try {
  Write-Verbose "$(Get-Date): Checking for valid StoreFront installation"
  $dsInstallProp = Get-ItemProperty -Path HKLM:\SOFTWARE\Citrix\DeliveryServicesManagement -Name InstallDir
  ##$dsInstallDir = $dsInstallProp.InstallDir 
  ##& $dsInstallDir\..\Scripts\ImportModules.ps1

  LoadStorefrontSnapins
  ImportStorefrontModules
 
} Catch {
  $dsInstallProp = $Null
}

if ($dsInstallProp -eq $Null) {
  Write-Verbose "$(Get-Date): Server does not have StoreFront installed ..."
  Write-Host "Server does not have StoreFront installed ... Aborting script."
  Return
}

#region Documentation GUI
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#
#   GUI component for StoreFront server documentation script
#   Author:      Sam Jacobs, IPM
#   Created:     July, 2014
#   Version:     1.0
#   Last Update: August 24, 2014 
#
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

$continueProcessing = $True

#~~< GUI Customizations go here >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# title used for the form and any pop-up message boxes
$GUI_title  = "StoreFront Documentation Script - Part 1 of 2"

#~~< Message Box buttons >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
[int]$MB_OK					= 0
[int]$MB_OK_CANCEL			= 1
[int]$MB_ABORT_RETRY_IGNORE = 2
[int]$MB_YES_NO_CANCEL		= 3
[int]$MB_YES_NO				= 4
[int]$MB_RETRY_CANCEL		= 5

#~~< Message Box icons >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
[int]$MB_ICON_CRITICAL		= 16
[int]$MB_ICON_QUESTION		= 32
[int]$MB_ICON_WARNING		= 48
[int]$MB_ICON_INFORMATIONAL	= 64

#~~< GUI Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Function getDirectory($prompt) {
	$objShell = New-Object -com Shell.Application
	$selectedFolder = $objShell.BrowseForFolder(0,$prompt,0,0)
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objShell) | Out-Null
	return $selectedFolder
}

Function populateDirectory ($startDir) {
	$selectedDir = getDirectory("Please select output directory:")
	if ($selectedDir -ne $Null) {
		$txtOutputDir.Text = $selectedDir.Self.Path
	}
}

Function Abort_Script() {
	$Script:continueProcessing = $False
	$frmServer.Close()
}

Function Continue_Script() {
	# save fields needed from form before closing

	$Script:OutputFile = $txtOutputFile.Text
	$Script:OutputDir  = $txtOutputDir.Text
	$Script:Software = ($chkSoftware.Checked -eq $True)
	$Script:Hardware = ($chkHardware.Checked -eq $True)
	
	# make sure the output directory actually exists!
	If (!(Test-Path ($Script:OutputDir))) {
		[System.Windows.Forms.MessageBox]::Show("Output directory does not exist!" , 
			$GUI_title, $MB_OK, $MB_ICON_CRITICAL)
		Return
	}
	$Script:continueProcessing = $True
	$frmServer.Close()
}

if ($GUI -eq $True) {
	Write-Verbose "$(Get-Date): Displaying GUI"
	#~~< create the GUI >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") |  Out-Null
	[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") |  Out-Null

	$frmServer = New-Object System.Windows.Forms.Form

	$btnExit = New-Object System.Windows.Forms.Button
	$btnGenerate = New-Object System.Windows.Forms.Button
	$groupBox3 = New-Object System.Windows.Forms.GroupBox
	$chkHardware = New-Object System.Windows.Forms.CheckBox
	$chkSoftware = New-Object System.Windows.Forms.CheckBox
	$groupBox1 = New-Object System.Windows.Forms.GroupBox
	$lblExt = New-Object System.Windows.Forms.Label
	$btnSelectOutputDir = New-Object System.Windows.Forms.Button
	$txtOutputDir = New-Object System.Windows.Forms.TextBox
	$txtOutputFile = New-Object System.Windows.Forms.TextBox
	$label4 = New-Object System.Windows.Forms.Label
	$label3 = New-Object System.Windows.Forms.Label
	## 
	## btnExit
	## 
	$btnExit.Location = New-Object System.Drawing.Point(268, 228)
	$btnExit.Name = "btnExit"
	$btnExit.Size = New-Object System.Drawing.Size(147, 30)
	$btnExit.TabIndex = 20
	$btnExit.Text = "Exit"
	$btnExit.UseVisualStyleBackColor = $True
	$btnExit.add_Click({Abort_Script})
	## 
	## btnGenerate
	## 
	$btnGenerate.Location = New-Object System.Drawing.Point(57, 228)
	$btnGenerate.Name = "btnGenerate"
	$btnGenerate.Size = New-Object System.Drawing.Size(147, 30)
	$btnGenerate.TabIndex = 19
	$btnGenerate.Text = "Generate"
	$btnGenerate.UseVisualStyleBackColor = $True
	$btnGenerate.Add_Click({Continue_Script})
	## 
	## groupBox3
	## 
	$groupBox3.Controls.Add($chkHardware)
	$groupBox3.Controls.Add($chkSoftware)
	$groupBox3.Location = New-Object System.Drawing.Point(36, 123)
	$groupBox3.Name = "groupBox3"
	$groupBox3.Size = New-Object System.Drawing.Size(403, 88)
	$groupBox3.TabIndex = 18
	$groupBox3.TabStop = $False
	$groupBox3.Text = " Optional "
	## 
	## chkHardware
	## 
	$chkHardware.AutoSize = $True
	$chkHardware.Location = New-Object System.Drawing.Point(29, 55)
	$chkHardware.Name = "chkHardware"
	$chkHardware.Size = New-Object System.Drawing.Size(217, 17)
	$chkHardware.TabIndex = 4
	$chkHardware.Text = "Use WMI to gather hardware information"
	$chkHardware.UseVisualStyleBackColor = $True
	$chkHardware.Checked = ($Hardware -eq $True)
	## 
	## chkSoftware
	## 
	$chkSoftware.AutoSize = $True
	$chkSoftware.Location = New-Object System.Drawing.Point(29, 26)
	$chkSoftware.Name = "chkSoftware"
	$chkSoftware.Size = New-Object System.Drawing.Size(189, 17)
	$chkSoftware.TabIndex = 3
	$chkSoftware.Text = "Query registry for installed software"
	$chkSoftware.UseVisualStyleBackColor = $True
	$chkSoftware.Checked = ($Software -eq $True)
	## 
	## groupBox1
	## 
	$groupBox1.Controls.Add($lblExt)
	$groupBox1.Controls.Add($btnSelectOutputDir)
	$groupBox1.Controls.Add($txtOutputDir)
	$groupBox1.Controls.Add($txtOutputFile)
	$groupBox1.Controls.Add($label4)
	$groupBox1.Controls.Add($label3)
	$groupBox1.Location = New-Object System.Drawing.Point(34, 12)
	$groupBox1.Name = "groupBox1"
	$groupBox1.Size = New-Object System.Drawing.Size(405, 100)
	$groupBox1.TabIndex = 21
	$groupBox1.TabStop = $False
	$groupBox1.Text = " Output "
	## 
	## lblExt
	## 
	$lblExt.AutoSize = $True
	$lblExt.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
	$lblExt.Location = New-Object System.Drawing.Point(340, 62)
	$lblExt.Name = "lblExt"
	$lblExt.Size = New-Object System.Drawing.Size(34, 15)
	$lblExt.TabIndex = 11
	$lblExt.Text = ".xml"
	## 
	## btnSelectOutputDir
	## 
	$btnSelectOutputDir.Location = New-Object System.Drawing.Point(338, 28)
	$btnSelectOutputDir.Name = "btnSelectOutputDir"
	$btnSelectOutputDir.Size = New-Object System.Drawing.Size(37, 20)
	$btnSelectOutputDir.TabIndex = 10
	$btnSelectOutputDir.Text = "..."
	$btnSelectOutputDir.UseVisualStyleBackColor = $True
	$btnSelectOutputDir.Add_Click({populateDirectory($OutputDir)})
	## 
	## txtOutputDir
	## 
	$txtOutputDir.Location = New-Object System.Drawing.Point(114, 28)
	$txtOutputDir.Name = "txtOutputDir"
	$txtOutputDir.Size = New-Object System.Drawing.Size(215, 20)
	$txtOutputDir.TabIndex = 9
	$txtOutputDir.Text = $OutputDir
	## 
	## txtOutputFile
	## 
	$txtOutputFile.Location = New-Object System.Drawing.Point(114, 60)
	$txtOutputFile.Name = "txtOutputFile"
	$txtOutputFile.Size = New-Object System.Drawing.Size(215, 20)
	$txtOutputFile.TabIndex = 8
	$txtOutputFile.Text = $OutputFile
	## 
	## label4
	## 
	$label4.AutoSize = $True
	$label4.Location = New-Object System.Drawing.Point(25, 28)
	$label4.Name = "label4"
	$label4.Size = New-Object System.Drawing.Size(49, 13)
	$label4.TabIndex = 1
	$label4.Text = "Directory"
	## 
	## label3
	## 
	$label3.AutoSize = $True
	$label3.Location = New-Object System.Drawing.Point(25, 61)
	$label3.Name = "label3"
	$label3.Size = New-Object System.Drawing.Size(49, 13)
	$label3.TabIndex = 0
	$label3.Text = "Filename"
	## 
	## frmServer
	## 
	$frmServer.ClientSize = New-Object System.Drawing.Size(475, 287)
	$frmServer.Controls.Add($groupBox1)
	$frmServer.Controls.Add($btnExit)
	$frmServer.Controls.Add($btnGenerate)
	$frmServer.Controls.Add($groupBox3)
	$frmServer.Name = "frmServer"
	$frmServer.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
	$frmServer.Text = $GUI_title

#region formIcon

# form icon - convert to base64
[string] $iconBase64=@"
AAABAAEAICAAAAEAIACoEAAAFgAAACgAAAAgAAAAQAAAAAEAIAAAAAAAgBAAAMQOAADEDgAAAAAAAAAA
AAD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8ASEhImk9PT7lKSko9////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wBISEjhVVVV/0lJSVn///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AEdHR9RQUFD/SEhIVP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8AR0dH1FBQUP9ISEhU////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wBHR0fUUFBQ/0dHR1X///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wBQUFA7TU1NRExMTBP///8A////AEdHR9RQUFD/SUlJRE5O
ThlPT092UlJSkE5OTmpLS0sQ////AP///wD///8A////AEZGRjVUVFREXl5eGv///wD///8AQkJCFlFR
UURFRUUr////AP///wB+fn4HTk5OQ0ZGRkQ9PT0H////AFJSUudSUlL/R0dHTf///wD///8AR0dH1E1N
Tf9HR0eZSEhI81ZWVv9UVFT/UlJS/0pKStpEREQk////AP///wD///8ASEhI01hYWP9fX19n////AP//
/wBHR0dYVlZW/0dHR6v///8A////AH9/fx1SUlL/TExM/0ZGRhz///8AUVFR3U5OTv9ISEhJ////AP//
/wBHR0fUSkpK/05OTv9HR0fOTk5Ofk5OTodJSUnpUVFR/0hISNleXl4H////AP///wBHR0fKVFRU/19f
X2L///8A////AEhISFRSUlL/R0dHpP///wD///8Ae3t7HE5OTvtISEj/SEhIG////wBRUVHcTk5O/0hI
SEn///8A////AEdHR9RPT0//TExMzWVlZQT///8A////AFBQUChKSkrrUFBQ/0pKSmL///8A////AEdH
R8lUVFT/X19fYv///wD///8ASEhIVFJSUv9HR0ej////AP///wB7e3scT09P+khISP9ISEgb////AFFR
UdxOTk7/SEhISf///wD///8AR0dH1FBQUP9PT09W////AP///wD///8A////AEpKSo5SUlL/TU1NsP//
/wD///8AR0dHyVRUVP9fX19i////AP///wBISEhUUlJS/0dHR6P///8A////AHt7exxPT0/6SEhI/0hI
SBv///8AUVFR3E5OTv9ISEhJ////AP///wBHR0fYS0tL/09PTxz///8A////AP///wD///8AR0dHX1JS
Uv9JSUnJ////AP///wBHR0fJVFRU/1xcXGL///8A////AEhISFNSUlL/SEhIof///wD///8Ae3t7HE9P
T/pISEj/SEhIG////wBRUVHcTk5O/0hISEn///8A////AEdHR9ZOTk7/T09POf///wD///8A////AP//
/wBLS0t1UlJS/01NTb7///8A////AEdHR8lSUlL/Tk5OXf///wD///8ASEhIU1BQUP9OTk7G////AP//
/wBtbW0bTU1N+khISP9ISEgb////AFFRUdxOTk7/SEhISf///wD///8AR0dH1FBQUP9PT0+j////AP//
/wD///8A////AEpKSs9SUlL/TU1Ng////wD///8AR0dHyVFRUf9NTU12////AP///wBPT0+DUFBQ/0lJ
SeD///8A////AEdHRxRJSUn6SkpK/09PTxz///8AUVFR3E5OTv9ISEhJ////AP///wBHR0fUS0tL/0xM
TP9JSUl8VFRULFRUVDVLS0uuT09P/0lJSftPT08f////AP///wBHR0fJTk5O/0lJSeNQUFA7VFRUNUlJ
SeBLS0v/T09P/05OToVUVFQsSUlJh05OTv9JSUnoVFRUBf///wBSUlLtVFRU/0lJSU////8A////AEhI
SOVRUVH/S0tL8lFRUf9PT0//TU1N/1FRUf9MTEz/R0dHX////wD///8A////AEdHR9lSUlL/UFBQ/1BQ
UP9OTk7/T09P/0lJSexMTEzRU1NT/01NTf9NTU3/UlJS/0lJSZH///8A////AFFRUYVNTU2ZSUlJLP//
/wD///8AR0dHgE1NTZlNTU03T09Pb0xMTM9NTU3lTU1NvUtLS03///8A////AP///wD///8AR0dHeVFR
UZlfX19cS0tLqU1NTd9LS0u8U1NTPHV1dQxNTU2XTU1N4E1NTdtOTk6eVFRUC////wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AKKiogr///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8AtbW1Bv///wD///8A////AP///wCfn58Hy8vLBP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AABs1l0AduyGAG3aG////wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wAXc9ggAXTm/wCA//8Abdik////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////ABN53RYBb9z/AID+/wBt15T///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AABq1UkAcuR8AGXMCf///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A///////////////////////////8f////H////x////8f////H///4wD
xjCMAcYwjADGMIwwxjCMeMYwjHjGMIx4xjCMeMYwjADAAIwBwAGMA8AB/+/95/////+P////D////w//
//+P//////////////////////////////8=
"@
		$iconStream=[System.IO.MemoryStream][System.Convert]::FromBase64String($iconBase64)
		$iconBmp=[System.Drawing.Bitmap][System.Drawing.Image]::FromStream($iconStream)
		$iconHandle=$iconBmp.GetHicon()
		$icon=[System.Drawing.Icon]::FromHandle($iconHandle)
		$frmServer.icon = $icon
#endregion formIcon

	# display the form
	[System.Windows.Forms.Application]::EnableVisualStyles()
	[System.Windows.Forms.Application]::Run($frmServer)

	If ($continueProcessing -eq $False) { 
		Write-Verbose "$(Get-Date): Script cancelled by user."
		Return 
	}
	#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	#  End of GUI for documentation scripts
	#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
}
#endregion Documentation GUI

#~~< script functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Function TranslateMethod($methodKey) {
	Switch($methodKey) {
		"ExplicitForms" 	{$methodValue = "User name and password"}
		"IntegratedWindows"	{$methodValue = "Domain pass-through"}
		"CitrixAGBasic"		{$methodValue = "Pass-through from NetScaler"}
		"HttpBasic"		{$methodValue = "HTTP Basic"}
		"Certificate"		{$methodValue = "Smart card"}
		Default			{$methodValue = "(Unknown authentication method)"}
	}
	return $methodValue
}

Function TranslateNSVersion($versionKey) {
	Switch($versionKey) {
		"Version10_0_69_4" 	{$versionValue = "10.0 (Build 69.4) or later"}
		"Version9x"		{$versionValue = "9.x"}
		"Version5x"		{$versionValue = "5.x"}
		Default			{$versionValue = $versionKey}
	}
	return $versionValue
}

Function TranslateHTML5Deployment($HTML5Key) {
	Switch($HTML5Key) {
		"Fallback" 	{$HTML5Value = "Use Receiver for HTML5 if local install fails"}
		"Always"	{$HTML5Value = "Always use Receiver for HTML5"}
		"Off"		{$HTML5Value = "Citrix Receiver installed locally"}
		Default		{$HTML5Value = $HTML5Key}
	}
	return $HTML5Value
}

Function TranslateLogonType($logonKey) {
	Switch($logonKey) {
		"DomainAndRSA" 	{$logonValue = "Domain and security token"}
		"Domain"	{$logonValue = "Domain"}
		"RSA"		{$logonValue = "Security token"}
		"SMS"		{$logonValue = "SMS authentication"}
		"SmartCard"	{$logonValue = "Smart card"}
		"None"		{$logonValue = "None"}
		Default		{$logonValue = $logonKey}
	}
	return $logonValue
}

Function TranslatePasswordOptions($pwKey) {
	Switch($pwKey) {
		"Always" 	{$pwValue = "At any time"}
		"ExpiredOnly" 	{$pwValue = "When expired"}
		"Never" 	{$pwValue = "Never"}
		Default		{$pwValue = $pwKey}
	}
	return $pwValue
}

Function YesNo ([boolean] $condition) {
    if ($condition -eq $True) {
	return "Yes"
    } Else {
	return "No"
    }
}

Function OutputGatewayDetails($service)
{
    $xmlOut = $Script:xmlWriter
    $allGws = (Get-DSGlobalGateways).Gateways
    foreach($gw in $service.Service.GatewayRefs)
    {         
        $vpnGws = $allGws | Where-Object { $_.ID -eq $gw.RefId }
        $Script:fullVPN = $service.Service.ServiceType -eq "VPN"
        if ($gw.Default) {$Script:defaultGW = $vpnGws.Name}
	$xmlOut.WriteElementString('RemoteGW', $($vpnGws.Name))
    }
}
 
#~~< WMI Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Function GetComputerWMIInfo
{
	[string]$ComputerName = $env:ComputerName
	$xmlOut = $Script:xmlWriter
	
	# original routines by Kees Baggerman, 
	# Senior Technical Consultant @ Inter Access
	# k.baggerman@myvirtualvision.com
	# @kbaggerman on Twitter
	# http://blog.myvirtualvision.com

	#Get Computer info
	Write-Verbose "$(Get-Date): `t`tProcessing WMI Computer information"
	Write-Verbose "$(Get-Date): `t`t`tHardware information"
	
	$xmlOut.WriteStartElement('Hardware')
		$xmlOut.WriteStartElement('General')
	
			[bool]$GotComputerItems = $True
			
			Try
			{
				$Results = Get-WmiObject -computername $ComputerName win32_computersystem
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
				Write-Verbose "$(Get-Date): Get-WmiObject win32_computersystem failed for $($ComputerName)"
				Write-Warning "Get-WmiObject win32_computersystem failed for $($ComputerName)"
				$xmlOut.WriteElementString('ErrorMessage', "Get-WmiObject win32_computersystem failed for $($ComputerName)")
			}
			Else
			{
				Write-Verbose "$(Get-Date): No results returned for Computer information"
				$xmlOut.WriteElementString('InfoMessage', "No results returned for Computer information")
			}
		$xmlOut.WriteEndElement()	# close General
			
		#Get Disk info
		Write-Verbose "$(Get-Date): `t`t`tDrive information"
		$xmlOut.WriteStartElement('DriveInfo')

		[bool]$GotDrives = $True
		
		Try
		{
			$Results = Get-WmiObject -computername $ComputerName Win32_LogicalDisk
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
					$xmlOut.WriteStartElement('Drive')
					OutputDriveItem $drive
					$xmlOut.WriteEndElement()	# close Drive
				}
			}
		}
		ElseIf(!$?)
		{
			Write-Verbose "$(Get-Date): Get-WmiObject Win32_LogicalDisk failed for $($ComputerName)"
			Write-Warning "Get-WmiObject Win32_LogicalDisk failed for $($ComputerName)"
			$xmlOut.WriteElementString('ErrorMessage', "Get-WmiObject Win32_LogicalDisk failed for $($ComputerName)")
		}
		Else
		{
			Write-Verbose "$(Get-Date): No results returned for Drive information"
			$xmlOut.WriteElementString('InfoMessage', "No results returned for Drive information")
		}
		$xmlOut.WriteEndElement()	# close DiskInfo

		$xmlOut.WriteStartElement('CPUs')
		#Get CPU's and stepping
		Write-Verbose "$(Get-Date): `t`t`tProcessor information"

		[bool]$GotProcessors = $True
		
		Try
		{
			$Results = Get-WmiObject -computername $ComputerName win32_Processor
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
				$xmlOut.WriteStartElement('CPU')
				OutputProcessorItem $processor
				$xmlOut.WriteEndElement()	# close CPU
			}
		}
		ElseIf(!$?)
		{
			Write-Verbose "$(Get-Date): Get-WmiObject win32_Processor failed for $($ComputerName)"
			Write-Warning "Get-WmiObject win32_Processor failed for $($ComputerName)"
			$xmlOut.WriteElementString('ErrorMessage', "Get-WmiObject win32_Processor failed for $($ComputerName)")
		}
		Else
		{
			Write-Verbose "$(Get-Date): No results returned for Processor information"
			$xmlOut.WriteElementString('InfoMessage', "No results returned for Processor information")
		}
	$xmlOut.WriteEndElement()	# close CPUs
	
	#Get Nics
	Write-Verbose "$(Get-Date): `t`t`tNIC information"
	$xmlOut.WriteStartElement('Network')

		[bool]$GotNics = $True
		
		Try
		{
			$Results = Get-WmiObject -computername $ComputerName win32_networkadapterconfiguration
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
						$ThisNic = Get-WmiObject -computername $ComputerName win32_networkadapter | Where {$_.index -eq $nic.index}
					}
					
					Catch 
					{
						$ThisNic = $Null
					}
					
					If($? -and $ThisNic -ne $Null)
					{
						$xmlOut.WriteStartElement('NIC')
						OutputNicItem $Nic $ThisNic
						$xmlOut.WriteEndElement()	# close NIC
					}
					ElseIf(!$?)
					{
						Write-Warning "$(Get-Date): Error retrieving NIC information"
						Write-Verbose "$(Get-Date): Get-WmiObject win32_networkadapterconfiguration failed for $($ComputerName)"
						Write-Warning "Get-WmiObject win32_networkadapterconfiguration failed for $($ComputerName)"
						$xmlOut.WriteElementString('ErrorMessage', "Error retrieving NIC information")
					}
					Else
					{
						Write-Verbose "$(Get-Date): No results returned for NIC information"
						$xmlOut.WriteElementString('InfoMessage', "No results returned for NIC information")
					}
				}
			}	
		}
		ElseIf(!$?)
		{
			Write-Warning "$(Get-Date): Error retrieving NIC configuration information"
			Write-Verbose "$(Get-Date): Get-WmiObject win32_networkadapterconfiguration failed for $($ComputerName)"
			Write-Warning "Get-WmiObject win32_networkadapterconfiguration failed for $($ComputerName)"
			$xmlOut.WriteElementString('ErrorMessage', "Error retrieving NIC configuration information")
		}
		Else
		{
			Write-Verbose "$(Get-Date): No results returned for NIC configuration information"
			$xmlOut.WriteElementString('InfoMessage', "No results returned for NIC configuration information")
		}
		
		$Results = $Null
		$ComputerItems = $Null
		$Drives = $Null
		$Processors = $Null
		$Nics = $Null
		$xmlOut.WriteEndElement()	# close Network
	$xmlOut.WriteEndElement()	# close Hardware
}

Function OutputComputerItem
{
	Param([object]$Item)

		$xmlOut.WriteElementString('Manufacturer', $Item.manufacturer)
		$xmlOut.WriteElementString('Model', $Item.model)
		$xmlOut.WriteElementString('Domain', $Item.domain)
		$xmlOut.WriteElementString('Ram', "$($Item.totalphysicalram) GB")

}

Function OutputDriveItem
{
	Param([object]$Drive)

		$xmlOut.WriteElementString('Caption', $Drive.caption)
		$xmlOut.WriteElementString('Size', "$($drive.drivesize) GB")
		If(![String]::IsNullOrEmpty($drive.filesystem))
		{
			$xmlOut.WriteElementString('FileSystem', $Drive.filesystem)
		}
		$xmlOut.WriteElementString('FreeSpace', "$($drive.drivefreespace) GB")
		If(![String]::IsNullOrEmpty($drive.volumename))
		{
			$xmlOut.WriteElementString('VolumeName', $Drive.volumename)
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
			$xmlOut.WriteElementString('isDirty', $tmp)

		}
		If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
		{
			$xmlOut.WriteElementString('SerialNumber', $Drive.volumeserialnumber)
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
		$xmlOut.WriteElementString('DriveType', $tmp)
}

Function OutputProcessorItem
{
	Param([object]$Processor)

		$xmlOut.WriteElementString('Name', $Processor.name)
		$xmlOut.WriteElementString('Description', $Processor.description)
		$xmlOut.WriteElementString('MaxClockSpeed',"$($processor.maxclockspeed) MHz")
		
		If($processor.l2cachesize -gt 0)
		{
			$xmlOut.WriteElementString('L2CacheSize', "$($processor.l2cachesize) KB")
		}
		If($processor.l3cachesize -gt 0)
		{
			$xmlOut.WriteElementString('L3 Cache Size', "$($processor.l3cachesize) KB")
		}
		If($processor.numberofcores -gt 0)
		{
			$xmlOut.WriteElementString('NumberOfCores', $Processor.numberofcores)
		}
		If($processor.numberoflogicalprocessors -gt 0)
		{
			$xmlOut.WriteElementString('NumLogicalCPUs', $Processor.numberoflogicalprocessors)
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
		$xmlOut.WriteElementString('Availability', $tmp)
}

Function OutputNicItem
{
	Param([object]$Nic, [object]$ThisNic)

		If($ThisNic.Name -eq $nic.description)
		{
			$xmlOut.WriteElementString('Name', $ThisNic.Name)
		}
		Else
		{
			$xmlOut.WriteElementString('Name', $ThisNic.Name)
			$xmlOut.WriteElementString('Description', $Nic.description)
		}
		$xmlOut.WriteElementString('ConnectionID', $ThisNic.NetConnectionID)
	##	$xmlOut.WriteElementString('Manufacturer', $Nic.manufacturer)
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
		$xmlOut.WriteElementString('Availability', $tmp)
		$xmlOut.WriteElementString('PhysicalAddress', $Nic.macaddress)
		$xmlOut.WriteElementString('IPAddress', $Nic.ipaddress)
		$xmlOut.WriteElementString('DefaultGateway', $Nic.Defaultipgateway)
		$xmlOut.WriteElementString('SubnetMask', $Nic.ipsubnet)
		If($nic.dhcpenabled)
		{
			$DHCPLeaseObtainedDate = $nic.ConvertToDateTime($nic.dhcpleaseobtained)
			$DHCPLeaseExpiresDate = $nic.ConvertToDateTime($nic.dhcpleaseexpires)
			$xmlOut.WriteElementString('DHCPEnabled', $Nic.dhcpenabled)
			$xmlOut.WriteElementString('DHCPLeaseObtained', $dhcpleaseobtaineddate)
			$xmlOut.WriteElementString('DHCPLeaseExpires', $dhcpleaseexpiresdate)
			$xmlOut.WriteElementString('DHCPServer', $Nic.dhcpserver)
		}
		If(![String]::IsNullOrEmpty($nic.dnsdomain))
		{
			$xmlOut.WriteElementString('DNSDomain', $Nic.dnsdomain)
		}
		If($nic.dnsdomainsuffixsearchorder -ne $Null -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
		{
			[int]$x = 1
			$xmlOut.WriteStartElement('DNSSearchSuffixes')
			$nicdnsdomainsuffixsearchorder = $nic.dnsdomainsuffixsearchorder
			
			ForEach($DNSDomain in $nicdnsdomainsuffixsearchorder)
			{
				$xmlOut.WriteElementString('DNSDomain', $DNSDomain)
			}
			$xmlOut.WriteEndElement()	# close DNSSearchSuffixes
		}
		If($nic.dnsenabledforwinsresolution)
		{
			$tmp = "Yes"
		}
		Else
		{
			$tmp = "No"
		}
		$xmlOut.WriteElementString('WINSEnabled', $tmp)
		If($nic.dnsserversearchorder -ne $Null -and $nic.dnsserversearchorder.length -gt 0)
		{
			$xmlOut.WriteStartElement('DNSServers')
			$nicdnsserversearchorder = $nic.dnsserversearchorder
			
			ForEach($DNSServer in $nicdnsserversearchorder)
			{
				$xmlOut.WriteElementString('DNSServer', "$($DNSServer)")
			}
			$xmlOut.WriteEndElement()	# close DNSServers
		}
		Switch ($nic.TcpipNetbiosOptions)
		{
			0	{$tmp = "Use NetBIOS setting from DHCP Server"}
			1	{$tmp = "Enable NetBIOS"}
			2	{$tmp = "Disable NetBIOS"}
			Default	{$tmp = "Unknown"}
		}
		$xmlOut.WriteElementString('NetBIOS', $tmp)
		If($nic.winsenablelmhostslookup)
		{
			$tmp = "Yes"
		}
		Else
		{
			$tmp = "No"
		}
		$xmlOut.WriteElementString('EnabledLMHosts', $tmp)
		If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
		{
			$xmlOut.WriteElementString('HostLookupFile', $Nic.winshostlookupfile)
		}
		If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
		{
			$xmlOut.WriteElementString('PrimaryServer', $Nic.winsprimaryserver)
		}
		If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
		{
			$xmlOut.WriteElementString('SecondaryServer', $Nic.winssecondaryserver)
		}
		If(![String]::IsNullOrEmpty($nic.winsscopeid))
		{
			$xmlOut.WriteElementString('ScopeID', $Nic.winsscopeid)
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
       # original work by Shaun Ritchie, Jeff Wouters, Webster
 
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


Function InstalledSoftware
{
	#get list of applications installed on server
	# original code by Shaun Ritchie, Jeff Wouters, Webster, Michael B. Smith
	$InstalledApps = @()
	$JustApps = @()
	$xmlOut = $Script:xmlWriter


	#Define the variable to hold the location of Currently Installed Programs
	$UninstallKey1="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall" 

	#Create an instance of the Registry Object and open the HKLM base key
	$reg=[microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$env:ComputerName) 

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
	$xmlOut.WriteStartElement('InstalledApplications')
	Write-Verbose "$(Get-Date): `t`tProcessing installed applications for server $($env:ComputerName)"
	
	ForEach($app in $JustApps)
	{
		Write-Verbose "$(Get-Date): `t`t`tProcessing installed application $($app.DisplayName)"
		$xmlOut.WriteStartElement('Application')
		$xmlOut.WriteElementString('Name', $app.DisplayName)
		$xmlOut.WriteElementString('Version', $app.DisplayVersion)
		$xmlOut.WriteEndElement()	# close Application
	}
	$xmlOut.WriteEndElement()	# close InstalledApplications
}

Function CitrixServices
{				
	#list citrix services
	$xmlOut = $Script:xmlWriter

		Write-Verbose "$(Get-Date): `t`tProcessing Citrix services for server $($env:ComputerName) by calling Get-Service"

		Try
		{
			#Iain Brighton optimization 5-Jun-2014
			#Replaced with a single call to retrieve services via WMI. The repeated
			## "Get-WMIObject Win32_Service -Filter" calls were the major delays in the script.
			## If we need to retrieve the StartUp type might as well just use WMI.
			$Services = Get-WMIObject Win32_Service -ComputerName $env:ComputerName -EA 0 | Where {$_.DisplayName -like "*Citrix*"} | Sort DisplayName
		}
		
		Catch
		{
			$Services = $Null
		}

		$xmlOut.WriteStartElement('CitrixServices')
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
			Write-Verbose "$(Get-Date): `t`t $NumServices Services found"
			$xmlOut.WriteElementString('Count', $NumServices)
			
			ForEach($Service in $Services) 
			{
				#Write-Verbose "$(Get-Date): `t`t`t Processing service $($Service.DisplayName)";
				$xmlOut.WriteStartElement('Service')
				$xmlOut.WriteElementString('Name', $Service.DisplayName)
				$xmlOut.WriteElementString('State', $Service.State)
				$xmlOut.WriteElementString('StartMode', $Service.StartMode)
				$xmlOut.WriteEndElement()	# close Service
			}
		}
		ElseIf(!$?)
		{
			Write-Warning "No services were retrieved."
			$xmlOut.WriteElementString('InfoMessage',  "No services were retrieved.")
		}
		Else
		{
			Write-Warning "Services retrieval was successful but no services were returned."
			$xmlOut.WriteElementString('InfoMessage',  "Services retrieval was successful but no services were returned." )
		}
		$xmlOut.WriteEndElement()	# close CitrixServices
}

Function GetSFVersion {
	$aSFVersion = (Get-DSVersion).StoreFrontVersion.Split(".")
	$dblVersion = [double] ($aSFVersion[0]+"."+$aSFVersion[1])
	return $dblVersion
}

#~~< script begins here >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$StartDateTime = Get-Date;

if ($OutputDir.EndsWith("\")) {
	$FullOutputName = $OutputDir + $OutputFile + ".xml"
} Else {
	$FullOutputName = $OutputDir + "\" + $OutputFile + ".xml"
}

# use the XMLTextWriter class to create the XML document
Write-Verbose "$(Get-Date): Creating XML document $FullOutputName"
Try {
	$XmlWriter = New-Object System.XMl.XmlTextWriter($FullOutputName,$Null)
} Catch {
	Write-Error "Could not create output file ... script is terminating."
	Exit
}

# choose a pretty formatting:
$xmlWriter.Formatting = 'Indented'
$xmlWriter.Indentation = 1
$XmlWriter.IndentChar = "`t"
 
# write the header / root element
$xmlWriter.WriteStartDocument()

$XmlWriter.WriteComment(' StoreFront documentation file ')
$XmlWriter.WriteComment(' Sam Jacobs, sjacobs@ipm.com ')
$XmlWriter.WriteComment(' Last update: July 20, 2015 ')

$xmlWriter.WriteStartElement('StoreFront')
	$xmlWriter.WriteStartElement('HostInfo')
	$XmlWriter.WriteElementString('ServerName', $env:ComputerName)
	$XmlWriter.WriteElementString('SFVersion', (get-DSVersion).StoreFrontVersion)
	$XmlWriter.WriteElementString('AsOf', $(Get-Date))
	$xmlWriter.WriteEndElement()	# close HostInfo

# server group
Write-Verbose "$(Get-Date): Processing server group"
$baseURL = $(get-DSHostBaseUrl).hostBaseUrl
$clusterMemberCount = $(get-DSClusterMembersCount).MembersCount
$xmlWriter.WriteStartElement('ServerGroup')
	$XmlWriter.WriteElementString('baseURL', $baseURL)
	$XmlWriter.WriteElementString('MemberCount', $clusterMemberCount)
	If (($clusterMemberCount -gt 1) -and ((GetSFVersion) -ge 2.5) ) {
		$lastSource = (Get-DSClusterConfigurationUpdateState).UpdateState.LastSourceServerName
		$XmlWriter.WriteElementString('LastSource', $lastSource)
	}
	$clusterMemberNames = $(get-DSClusterMembersName).HostNames
	$xmlWriter.WriteStartElement('Members')
	ForEach ($member in $clusterMemberNames) 
	{
		$lastSync = $(Get-DSXdServerGroupConfigurationUpdateState $member).LastEndTime.ToString()
		$xmlWriter.WriteStartElement('Member')
		$XmlWriter.WriteAttributeString('Name', $member)
		If ($clusterMemberCount -gt 1) {
			$XmlWriter.WriteAttributeString('LastSync', $lastSync)
		}
		$xmlWriter.WriteEndElement()
	}
	$xmlWriter.WriteEndElement()	# close Members
$xmlWriter.WriteEndElement()	# close ServerGroup

# authentication
Write-Verbose "$(Get-Date): Processing authentication"
$auth = $(Get-DSAuthenticationServicesSummary)
$tokenURL = $auth.TokenIssuerUrl + "/validate"
If($auth.UseHttps -eq $True)
{
	$status = "Service using HTTPS"
	$daysToExpire = (New-TimeSpan -End $auth.IISCertificate.NotAfter).Days
}
Else
{
	$status = "Service NOT using HTTPS"
	$daysToExpire = ""
}
$xmlWriter.WriteStartElement('Authentication')
	$xmlWriter.WriteStartElement('Methods')
	[int]$enabledMethods = 0
	ForEach ($protocol in $auth.protocols)
	{
		If($protocol.DisplayInConsole -eq $True)
		{
			$method = TranslateMethod($protocol.choice) 

			Switch($protocol.enabled)
			{
				$True	{$enabled = "Yes"; ++$enabledMethods}
				Default	{$enabled = "No"}
			}
			$xmlWriter.WriteStartElement('Method')
			$XmlWriter.WriteAttributeString('name', $method)
			$XmlWriter.WriteAttributeString('enabled', $enabled)
			$xmlWriter.WriteEndElement()	# close Method
		}
	}
	$xmlWriter.WriteEndElement()	# close Methods
	
	$xmlWriter.WriteElementString('tokenURL', $tokenURL)
	$XmlWriter.WriteElementString('EnabledMethods', $enabledMethods)
	$XmlWriter.WriteElementString('status', $status)
	if($daysToExpire -is [int])
	{
        $XmlWriter.WriteElementString('SSLExpiration',$auth.IISCertificate.NotAfter)
		$XmlWriter.WriteElementString('daysUntilExpiration', $daysToExpire)
	}

	# authentication domains
	$domainInfo = Get-DSExplicitCommonDataModel ($auth.SiteID) $auth.VirtualPath
	$defDomain = ($domaininfo.DefaultDomain).DefaultDomain
	$changePW = TranslatePasswordOptions($domainInfo.AllowUserToChangePassword)
	
	$xmlWriter.WriteStartElement('TrustedDomains')
	$XmlWriter.WriteElementString('DomainCount', $domainInfo.Domains.Count)
	If($domainInfo.Domains.Count -gt 0)
	{
		$xmlWriter.WriteElementString('AllowLogOn', 'Trusted domains only')
		$XmlWriter.WriteElementString('DefaultDomain', $defDomain)
		if ((GetSFVersion) -ge 2.6) {
			$showDomains = YesNo($domainInfo.ShowDomainField)
			$XmlWriter.WriteElementString('ShowDomains', $showDomains)
		}
		
		ForEach($domain in $domainInfo.Domains)
		{
			$xmlWriter.WriteStartElement('Domain')
			$XmlWriter.WriteAttributeString('name', $domain)
			$xmlWriter.WriteEndElement()	# close Domain
		}		
		
	} Else {
		$xmlWriter.WriteElementString('AllowLogOn', 'Any domain')
	}
	$xmlWriter.WriteEndElement()	# close TrustedDomains

	$xmlWriter.WriteElementString('changePW', $changePW)
	
$xmlWriter.WriteEndElement()	# close authentication

# stores
Write-Verbose "$(Get-Date): Processing stores"
$accounts = @(((Get-DSGlobalAccounts).Accounts) | Sort Name)
$xmlWriter.WriteStartElement('Stores')

	foreach ($account in $accounts) {

		#$acctInfo = (Get-DSGlobalAccount -AccountId $account.Id).Account
		#$advertised = YesNo($acctInfo.Published)
		#$store = Get-DSStoreServicesSummary | where {$_.FriendlyName -eq $acctInfo.Name}

		$advertised = YesNo($account.Published)
		$store = Get-DSStoreServicesSummary | where {$_.FriendlyName -eq $account.Name}

		$xmlWriter.WriteStartElement('Store')
		$friendlyName = $store.FriendlyName
		$URL = $store.Url
		if ($store.GatewayKeys.Count -gt 0) { 
			$access = "Internal and external networks"
		} Else { 
			$access = "Internal networks only"
		}
		if ($store.UseHttps -eq $True) {
			$status = "Service using HTTPS"
		} Else {
			$status = "Service using HTTP"
		}

		if ((GetSFVersion) -ge 2.5) {
			$locked = YesNo($store.IsLockedDown)
			$authenticated = YesNo(!$store.IsAnonymous)
                        $filterTypes = Get-DSResourceFilterType $store.SiteID $store.VirtualPath
                        $filterKeywords = Get-DSResourceFilterkeyword $store.SiteID $store.VirtualPath
                        $includeKeywords = @($filterKeywords.Include)
                        $excludeKeywords = @($filterKeywords.Exclude)
		}

		$xmlWriter.WriteElementString('Name', $friendlyName)
		$xmlWriter.WriteElementString('StoreURL', $URL)
		$xmlWriter.WriteElementString('Access', $access)
		if ((GetSFVersion) -ge 2.5) {
			$xmlWriter.WriteElementString('Locked', $locked)
			$xmlWriter.WriteElementString('Authenticated', $authenticated)
			$xmlWriter.WriteStartElement('ResourceFilters')
			   $xmlWriter.WriteStartElement('IncludedTypes')
			   foreach ($resource in $filterTypes) {
				$xmlWriter.WriteElementString('includedType', $resource)
			   }
			   $xmlWriter.WriteEndElement()  # close IncludedTypes
			   $xmlWriter.WriteStartElement('IncludedKeywords')
			   foreach ($resource in $includeKeywords) {
				if ($resource -ne $Null) {
				   $xmlWriter.WriteElementString('includedKeyword', $resource)
				}
			   }
			   $xmlWriter.WriteEndElement()  # close IncludedKeywords
			   $xmlWriter.WriteStartElement('ExcludedKeywords')
			   foreach ($resource in $excludeKeywords) {
				$xmlWriter.WriteElementString('excludedKeyword', $resource)
			   }
			   $xmlWriter.WriteEndElement()  # close ExcludedKeywords
			$xmlWriter.WriteEndElement()  # close ResourceFilters
		}
		$xmlWriter.WriteElementString('Advertised', $advertised)
		$xmlWriter.WriteElementString('Status', $status)

		$farmsets = @($store.Farmsets)
		foreach ($farmset in $farmsets) {	
	   	$farms = @($farmset.Farms)
		$xmlWriter.WriteStartElement('Farms')
			foreach ($farm in $farms) {
				$farmName = $farm.FarmName 
				$farmType = $farm.FarmType
				$farmServers = $farm.Servers
				$transportType = $farm.TransportType
				$port = $farm.ServicePort
				$sslRelayPort = $farm.SSLRelayPort
				$loadBalance = YesNo($farm.LoadBalance)
				$xmlWriter.WriteStartElement('Farm')
					$xmlWriter.WriteElementString('FarmName', $farmName)
					$xmlWriter.WriteElementString('FarmType', $farmType)
					$xmlWriter.WriteStartElement('Servers')
					foreach ($server in $farmServers)
					{
						$xmlWriter.WriteStartElement('Server')
						$XmlWriter.WriteAttributeString('name', $server)
						$xmlWriter.WriteEndElement()	# close Server
					}
					$xmlWriter.WriteEndElement()  # close Servers
					$xmlWriter.WriteElementString('XMLPort', $port)
					if ($farmType -ne "AppController") {
						$xmlWriter.WriteElementString('TransportType', $transportType)
						$xmlWriter.WriteElementString('SSLRelayPort', $sslRelayPort)
						$xmlWriter.WriteElementString('LoadBalance', $loadBalance)
					}
				$xmlWriter.WriteEndElement()  # close Farm
			}		
			$xmlWriter.WriteEndElement()  # close Farms
		}
		
		$xmlWriter.WriteStartElement('CitrixOnline')
			$GoToMeeting = YesNo($store.IsGoToMeetingEnabled)
			$GotoWebinar = YesNo($store.IsGoToWebinarEnabled)
			$GoToTraining = YesNo($store.IsGoToTrainingEnabled)
			$xmlWriter.WriteElementString('GoToMeeting', $GoToMeeting)
			$xmlWriter.WriteElementString('GoToWebinar', $GotoWebinar)
			$xmlWriter.WriteElementString('GoToTraining', $GoToTraining)
		$xmlWriter.WriteEndElement()	# close CitrixOnline

		# remote access
		$defaultGW = ""
		$fullVPN = $False
		$xmlWriter.WriteStartElement('RemoteAccess')
		$vpnService = Get-DSGlobalService -ServiceRef "VPN_$($store.ServiceRef)"
		if($vpnService.Service)
		{
    			OutputGatewayDetails($vpnService)
		}
		else
		{
    			$service = Get-DSGlobalService -ServiceRef $store.ServiceRef
   			 OutputGatewayDetails($service)
		}

		switch ($defaultGW)
		{
   			""		{ $xmlWriter.WriteElementString('RemoteType', "None") }
			default	
			{
				$xmlWriter.WriteElementString('DefaultGW', $defaultGW)
				if ($fullVPN)	{ $xmlWriter.WriteElementString('RemoteType', "Full VPN Tunnel") }
				else		{ $xmlWriter.WriteElementString('RemoteType', "No VPN Tunnel") }
			}
		}				

		$xmlWriter.WriteEndElement()	# close RemoteAccess

		$xmlWriter.WriteEndElement()  # close Store
	}
$xmlWriter.WriteEndElement()  # close stores

# Receiver for Web
Write-Verbose "$(Get-Date): Processing Receiver for Web sites"
$receivers = @(Get-DSWebReceiversSummary)
$xmlWriter.WriteStartElement('ReceiverForWeb')

	foreach ($receiver in $receivers) {
	
		$name = $receiver.FriendlyName
		$WebUrl  = $receiver.Url
		if ((GetSFVersion) -ge 2.5) {
			$authenticated = YesNo(!$receiver.IsAnonymousStore)
			$HTML5version = $receiver.HTML5Version
			$authMethods = @($receiver.AllowedAuthMethods)
		}
		$storeURL = $receiver.StoreUrl
		$aStore = $storeURL -split "/"
		$store = $aStore[$aStore.Count-1]
		$deployment = TranslateHTML5Deployment($receiver.HTML5Configuration)
		$shortcuts = Get-DSAppShortcutsTrustedUrls $receiver.SiteId $receiver.VirtualPath
		$xmlWriter.WriteStartElement('Receiver')
			$xmlWriter.WriteElementString('Name', $name)
			$xmlWriter.WriteElementString('WebURL', $WebUrl)
			$xmlWriter.WriteElementString('Store', $store)
			$xmlWriter.WriteElementString('StoreURL', $storeURL)

			if ((GetSFVersion) -ge 2.5) {
			   $xmlWriter.WriteElementString('Authenticated', $authenticated)
			   $xmlWriter.WriteElementString('HTML5Version', $HTML5version)
			
			   $xmlWriter.WriteStartElement('AuthMethods')
			   foreach ($authMethod in $authMethods) {
				$method = TranslateMethod($authMethod)
				$xmlWriter.WriteStartElement('AuthMethod')
				$XmlWriter.WriteAttributeString('name', $method)
				$xmlWriter.WriteEndElement()	# close AuthMethod
			   }
			   $xmlWriter.WriteEndElement()	# close AuthMethods
			}
			
			if ($shortcuts -ne $Null) {
				$TrustedURLs = @($shortcuts.TrustedUrls)
				$xmlWriter.WriteStartElement('Shortcuts')
				foreach ($url in $TrustedURLs) {
					$xmlWriter.WriteStartElement('Shortcut')
					$XmlWriter.WriteAttributeString('url', $url)
					$xmlWriter.WriteEndElement()	# close Shortcut
				}				
				$xmlWriter.WriteEndElement()	# close Shortcuts

				$xmlWriter.WriteElementString('Deployment', $deployment)
			}
		$xmlWriter.WriteEndElement()	# close Receiver
	}
	$xmlWriter.WriteEndElement()	# close ReceiverForWeb
	
# NetScaler Gateways
Write-Verbose "$(Get-Date): Processing NetScaler Gateways"
$gateways = @((Get-DSGlobalGateways).Gateways)
$xmlWriter.WriteStartElement('Gateways')

	foreach ($gateway in $gateways) {
		$name = $gateway.Name
		$used = "Yes"
		$url = $gateway.Address
		$NSversion = TranslateNSVersion($gateway.AccessGatewayVersion)
		$callbackURL = $gateway.CallbackURL
		$deploymentMode = $gateway.DeploymentType
		$STAs = $gateway.SecureTicketAuthorityURLs
		if ($gateway.SessionReliability -eq $True) {$sessionReliability="Yes"} Else {$sessionReliability="No"}
		if ($gateway.RequestTicketTwoSTA -eq $True) {$request2STATickets="Yes"} Else {$request2STATickets="No"}
		$xmlWriter.WriteStartElement('Gateway')
			$xmlWriter.WriteElementString('Name', $name)
			$xmlWriter.WriteElementString('GatewayURL', $url)
			$xmlWriter.WriteElementString('Version', $NSversion)
			if ($NSversion -ne "5.x") {
				$logonType = TranslateLogonType($gateway.Logon)
				$smartCardFallback = TranslateLogonType($gateway.SmartCardFallback)				
				$xmlWriter.WriteElementString('SubnetIP', $gateway.IPAddress)
				$xmlWriter.WriteElementString('LogonType', $logonType)
				$xmlWriter.WriteElementString('SmartcardFallback', $smartCardFallback)
			}
			$xmlWriter.WriteElementString('CallbackURL', $callbackURL)
			
			$xmlWriter.WriteStartElement('STAs')
			foreach ($sta in $STAs) {
				$xmlWriter.WriteStartElement('STA')
				$XmlWriter.WriteAttributeString('Address', $sta)
				$xmlWriter.WriteEndElement()	# close STA
			}
			$xmlWriter.WriteEndElement()	# close STAs
			$xmlWriter.WriteElementString('SessionReliability', $sessionReliability)
			$xmlWriter.WriteElementString('Request2STAtickets', $request2STATickets)
		$xmlWriter.WriteEndElement()	# close Gateway
	}
$xmlWriter.WriteEndElement()	# close Gateways

# Beacons
Write-Verbose "$(Get-Date): Processing Beacons"
$internalBeacons = @((Get-DSGlobalBeacons "Internal").Beacons)	
$externalBeacons = @((Get-DSGlobalBeacons "External").Beacons)

$xmlWriter.WriteStartElement('Beacons')
	if ($internalBeacons.Count -gt 0) {
		$xmlWriter.WriteStartElement('Internal')
		foreach ($beacon in $internalBeacons) {
			$beaconAddress = ($beacon).Address
			$xmlWriter.WriteElementString('BeaconURL', $beaconAddress)
		}
		$xmlWriter.WriteEndElement()	# close Internal
	}
	
	if ($externalBeacons.Count -gt 0) {
		$xmlWriter.WriteStartElement('External')
		foreach ($beacon in $externalBeacons) {
			$beaconAddress = ($beacon).Address
			$xmlWriter.WriteElementString('BeaconURL', $beaconAddress)
		}
		$xmlWriter.WriteEndElement()	# close External
	}
	
$xmlWriter.WriteEndElement()	# close Beacons

# Was a hardware report requested?
if ($Hardware -eq $True) {
	GetComputerWMIInfo
}

# Was a software report requested?
if ($Software -eq $True) {
	$xmlWriter.WriteStartElement('Software')
	InstalledSoftware
	CitrixServices
	$xmlWriter.WriteEndElement() # close Software
}

Write-Verbose "$(Get-Date): Closing XML document"
# close the root node
$xmlWriter.WriteEndElement()
 
# finalize the document:
$xmlWriter.WriteEndDocument()
$xmlWriter.Flush()
$xmlWriter.Close()

Write-Host ""
Write-Host "$(Get-Date): XML document is available at: $($FullOutputName)" 

$EndDateTime = Get-Date;
Write-Host "$(Get-Date): Script started: $($StartDateTime)"
Write-Host "$(Get-Date): Script ended: $($EndDateTime)"
$runtime = $($EndDateTime) - $StartDateTime
$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds", `
	$runtime.Days, `
	$runtime.Hours, `
	$runtime.Minutes, `
	$runtime.Seconds,
	$runtime.Milliseconds)
Write-Host "$(Get-Date): Elapsed time: $($Str)"
$runtime = $Null
$Str = $Null
		
Write-Host "$(Get-Date): Done."

# SIG # Begin signature block
# MIIcXgYJKoZIhvcNAQcCoIIcTzCCHEsCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU7UsL5Cr5Tb8ZidAw1I1oUhwa
# aWqggheNMIIFFjCCA/6gAwIBAgIQDDQt7Y1XuhbY8RuUWoZC3jANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMB4XDTE1MDcwMzAwMDAwMFoXDTE2MDcw
# NzEyMDAwMFowXTELMAkGA1UEBhMCVVMxETAPBgNVBAgTCE5ldyBZb3JrMREwDwYD
# VQQHEwhCcm9va2x5bjETMBEGA1UEChMKU2FtIEphY29iczETMBEGA1UEAxMKU2Ft
# IEphY29iczCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAOQ4wYU8jGsS
# RWKjY9nEagcepUTxVH5KYvzo9QZ+Zq2aBYxvwvq5A6l4ufm55r8yWinzzty23wsg
# /NYyLiQMpUICPNQmlNow2sJQTwZ2apHaN4EMnOyNSqE96ctmVP8UG+4OUV+47kH0
# xGc+CL8oW4UJPFXNQXDYMotwoMyIBx9idPeGJn7SbQj28fsa6xPCcnyYNT3/KjLU
# PIdSPvyDrunKNfbdh1UUVPqjevEx90Fwgk9rz8Oi7+v8piD+/k3/gPxhIq4okA4s
# Bk9vISAWGV8u89UufbH0p1fPhuoDvV7DW+SdbNiZPkECNL0NKLcC/LG4olnnkcjV
# a3NYgWm8PhECAwEAAaOCAbswggG3MB8GA1UdIwQYMBaAFFrEuXsqCqOl6nEDwGD5
# LfZldQ5YMB0GA1UdDgQWBBQWHubvgEwt1VoZUpyezIIll2vT4jAOBgNVHQ8BAf8E
# BAMCB4AwEwYDVR0lBAwwCgYIKwYBBQUHAwMwdwYDVR0fBHAwbjA1oDOgMYYvaHR0
# cDovL2NybDMuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC1jcy1nMS5jcmwwNaAz
# oDGGL2h0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9zaGEyLWFzc3VyZWQtY3MtZzEu
# Y3JsMEIGA1UdIAQ7MDkwNwYJYIZIAYb9bAMBMCowKAYIKwYBBQUHAgEWHGh0dHBz
# Oi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwgYQGCCsGAQUFBwEBBHgwdjAkBggrBgEF
# BQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tME4GCCsGAQUFBzAChkJodHRw
# Oi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRTSEEyQXNzdXJlZElEQ29k
# ZVNpZ25pbmdDQS5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAQEA
# EvjF6gTkMvIxd2EmcRt/oG9wLL33kuXzIYNUPj0pQIub0yAKsDzJmHpoJdsSaTjW
# Taf7Y9ZLOdvPHJB4si2cAof3C3xFxBHD2IFMZy5dF53/VRWA82Q7SzIBEQbuhIYp
# EzaiLXOmJ7TbKxlmpvHDNVtFevDqgoHR9KTS+J6Ycy0yDkwju+W09OcosMuHH1SR
# kQvixNpVwn41e9tU2lD7oHVh3ct6Oiz6a/J563gDFcYmvvmrW4vJTEBZjzLvkUaB
# U8AbKxy5WHJO12nV3O2SHaLntIn2RY+XH/NIHGnoZzS7O6ZzogBEwrYpKe9vubSp
# +f/CpbZgSrBnDrLNeVUTyDCCBTAwggQYoAMCAQICEAQJGBtf1btmdVNDtW+VUAgw
# DQYJKoZIhvcNAQELBQAwZTELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0
# IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNl
# cnQgQXNzdXJlZCBJRCBSb290IENBMB4XDTEzMTAyMjEyMDAwMFoXDTI4MTAyMjEy
# MDAwMFowcjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcG
# A1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBB
# c3N1cmVkIElEIENvZGUgU2lnbmluZyBDQTCCASIwDQYJKoZIhvcNAQEBBQADggEP
# ADCCAQoCggEBAPjTsxx/DhGvZ3cH0wsxSRnP0PtFmbE620T1f+Wondsy13Hqdp0F
# LreP+pJDwKX5idQ3Gde2qvCchqXYJawOeSg6funRZ9PG+yknx9N7I5TkkSOWkHeC
# +aGEI2YSVDNQdLEoJrskacLCUvIUZ4qJRdQtoaPpiCwgla4cSocI3wz14k1gGL6q
# xLKucDFmM3E+rHCiq85/6XzLkqHlOzEcz+ryCuRXu0q16XTmK/5sy350OTYNkO/k
# tU6kqepqCquE86xnTrXE94zRICUj6whkPlKWwfIPEvTFjg/BougsUfdzvL2FsWKD
# c0GCB+Q4i2pzINAPZHM8np+mM6n9Gd8lk9ECAwEAAaOCAc0wggHJMBIGA1UdEwEB
# /wQIMAYBAf8CAQAwDgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMD
# MHkGCCsGAQUFBwEBBG0wazAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNl
# cnQuY29tMEMGCCsGAQUFBzAChjdodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20v
# RGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3J0MIGBBgNVHR8EejB4MDqgOKA2hjRo
# dHRwOi8vY3JsNC5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0Eu
# Y3JsMDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1
# cmVkSURSb290Q0EuY3JsME8GA1UdIARIMEYwOAYKYIZIAYb9bAACBDAqMCgGCCsG
# AQUFBwIBFhxodHRwczovL3d3dy5kaWdpY2VydC5jb20vQ1BTMAoGCGCGSAGG/WwD
# MB0GA1UdDgQWBBRaxLl7KgqjpepxA8Bg+S32ZXUOWDAfBgNVHSMEGDAWgBRF66Kv
# 9JLLgjEtUYunpyGd823IDzANBgkqhkiG9w0BAQsFAAOCAQEAPuwNWiSz8yLRFcgs
# fCUpdqgdXRwtOhrE7zBh134LYP3DPQ/Er4v97yrfIFU3sOH20ZJ1D1G0bqWOWuJe
# JIFOEKTuP3GOYw4TS63XX0R58zYUBor3nEZOXP+QsRsHDpEV+7qvtVHCjSSuJMbH
# JyqhKSgaOnEoAjwukaPAJRHinBRHoXpoaK+bp1wgXNlxsQyPu6j4xRJon89Ay0BE
# pRPw5mQMJQhCMrI2iiQC/i9yfhzXSUWW6Fkd6fp0ZGuy62ZD2rOwjNXpDd32ASDO
# mTFjPQgaGLOBm0/GkxAG/AeB+ova+YJJ92JuoVP6EpQYhS6SkepobEQysmah5xik
# mmRR7zCCBmowggVSoAMCAQICEAMBmgI6/1ixa9bV6uYX8GYwDQYJKoZIhvcNAQEF
# BQAwYjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UE
# CxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgQXNzdXJlZCBJ
# RCBDQS0xMB4XDTE0MTAyMjAwMDAwMFoXDTI0MTAyMjAwMDAwMFowRzELMAkGA1UE
# BhMCVVMxETAPBgNVBAoTCERpZ2lDZXJ0MSUwIwYDVQQDExxEaWdpQ2VydCBUaW1l
# c3RhbXAgUmVzcG9uZGVyMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA
# o2Rd/Hyz4II14OD2xirmSXU7zG7gU6mfH2RZ5nxrf2uMnVX4kuOe1VpjWwJJUNmD
# zm9m7t3LhelfpfnUh3SIRDsZyeX1kZ/GFDmsJOqoSyyRicxeKPRktlC39RKzc5YK
# Z6O+YZ+u8/0SeHUOplsU/UUjjoZEVX0YhgWMVYd5SEb3yg6Np95OX+Koti1ZAmGI
# YXIYaLm4fO7m5zQvMXeBMB+7NgGN7yfj95rwTDFkjePr+hmHqH7P7IwMNlt6wXq4
# eMfJBi5GEMiN6ARg27xzdPpO2P6qQPGyznBGg+naQKFZOtkVCVeZVjCT88lhzNAI
# zGvsYkKRrALA76TwiRGPdwIDAQABo4IDNTCCAzEwDgYDVR0PAQH/BAQDAgeAMAwG
# A1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwggG/BgNVHSAEggG2
# MIIBsjCCAaEGCWCGSAGG/WwHATCCAZIwKAYIKwYBBQUHAgEWHGh0dHBzOi8vd3d3
# LmRpZ2ljZXJ0LmNvbS9DUFMwggFkBggrBgEFBQcCAjCCAVYeggFSAEEAbgB5ACAA
# dQBzAGUAIABvAGYAIAB0AGgAaQBzACAAQwBlAHIAdABpAGYAaQBjAGEAdABlACAA
# YwBvAG4AcwB0AGkAdAB1AHQAZQBzACAAYQBjAGMAZQBwAHQAYQBuAGMAZQAgAG8A
# ZgAgAHQAaABlACAARABpAGcAaQBDAGUAcgB0ACAAQwBQAC8AQwBQAFMAIABhAG4A
# ZAAgAHQAaABlACAAUgBlAGwAeQBpAG4AZwAgAFAAYQByAHQAeQAgAEEAZwByAGUA
# ZQBtAGUAbgB0ACAAdwBoAGkAYwBoACAAbABpAG0AaQB0ACAAbABpAGEAYgBpAGwA
# aQB0AHkAIABhAG4AZAAgAGEAcgBlACAAaQBuAGMAbwByAHAAbwByAGEAdABlAGQA
# IABoAGUAcgBlAGkAbgAgAGIAeQAgAHIAZQBmAGUAcgBlAG4AYwBlAC4wCwYJYIZI
# AYb9bAMVMB8GA1UdIwQYMBaAFBUAEisTmLKZB+0e36K+Vw0rZwLNMB0GA1UdDgQW
# BBRhWk0ktkkynUoqeRqDS/QeicHKfTB9BgNVHR8EdjB0MDigNqA0hjJodHRwOi8v
# Y3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURDQS0xLmNybDA4oDag
# NIYyaHR0cDovL2NybDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEQ0Et
# MS5jcmwwdwYIKwYBBQUHAQEEazBpMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5k
# aWdpY2VydC5jb20wQQYIKwYBBQUHMAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydEFzc3VyZWRJRENBLTEuY3J0MA0GCSqGSIb3DQEBBQUAA4IB
# AQCdJX4bM02yJoFcm4bOIyAPgIfliP//sdRqLDHtOhcZcRfNqRu8WhY5AJ3jbITk
# WkD73gYBjDf6m7GdJH7+IKRXrVu3mrBgJuppVyFdNC8fcbCDlBkFazWQEKB7l8f2
# P+fiEUGmvWLZ8Cc9OB0obzpSCfDscGLTYkuw4HOmksDTjjHYL+NtFxMG7uQDthSr
# 849Dp3GdId0UyhVdkkHa+Q+B0Zl0DSbEDn8btfWg8cZ3BigV6diT5VUW8LsKqxzb
# XEgnZsijiwoc5ZXarsQuWaBh3drzbaJh6YoLbewSGL33VVRAA5Ira8JRwgpIr7DU
# buD0FAo6G+OPPcqvao173NhEMIIGzTCCBbWgAwIBAgIQBv35A5YDreoACus/J7u6
# GzANBgkqhkiG9w0BAQUFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNl
# cnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdp
# Q2VydCBBc3N1cmVkIElEIFJvb3QgQ0EwHhcNMDYxMTEwMDAwMDAwWhcNMjExMTEw
# MDAwMDAwWjBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkw
# FwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBBc3N1
# cmVkIElEIENBLTEwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDogi2Z
# +crCQpWlgHNAcNKeVlRcqcTSQQaPyTP8TUWRXIGf7Syc+BZZ3561JBXCmLm0d0nc
# icQK2q/LXmvtrbBxMevPOkAMRk2T7It6NggDqww0/hhJgv7HxzFIgHweog+SDlDJ
# xofrNj/YMMP/pvf7os1vcyP+rFYFkPAyIRaJxnCI+QWXfaPHQ90C6Ds97bFBo+0/
# vtuVSMTuHrPyvAwrmdDGXRJCgeGDboJzPyZLFJCuWWYKxI2+0s4Grq2Eb0iEm09A
# ufFM8q+Y+/bOQF1c9qjxL6/siSLyaxhlscFzrdfx2M8eCnRcQrhofrfVdwonVnwP
# YqQ/MhRglf0HBKIJAgMBAAGjggN6MIIDdjAOBgNVHQ8BAf8EBAMCAYYwOwYDVR0l
# BDQwMgYIKwYBBQUHAwEGCCsGAQUFBwMCBggrBgEFBQcDAwYIKwYBBQUHAwQGCCsG
# AQUFBwMIMIIB0gYDVR0gBIIByTCCAcUwggG0BgpghkgBhv1sAAEEMIIBpDA6Bggr
# BgEFBQcCARYuaHR0cDovL3d3dy5kaWdpY2VydC5jb20vc3NsLWNwcy1yZXBvc2l0
# b3J5Lmh0bTCCAWQGCCsGAQUFBwICMIIBVh6CAVIAQQBuAHkAIAB1AHMAZQAgAG8A
# ZgAgAHQAaABpAHMAIABDAGUAcgB0AGkAZgBpAGMAYQB0AGUAIABjAG8AbgBzAHQA
# aQB0AHUAdABlAHMAIABhAGMAYwBlAHAAdABhAG4AYwBlACAAbwBmACAAdABoAGUA
# IABEAGkAZwBpAEMAZQByAHQAIABDAFAALwBDAFAAUwAgAGEAbgBkACAAdABoAGUA
# IABSAGUAbAB5AGkAbgBnACAAUABhAHIAdAB5ACAAQQBnAHIAZQBlAG0AZQBuAHQA
# IAB3AGgAaQBjAGgAIABsAGkAbQBpAHQAIABsAGkAYQBiAGkAbABpAHQAeQAgAGEA
# bgBkACAAYQByAGUAIABpAG4AYwBvAHIAcABvAHIAYQB0AGUAZAAgAGgAZQByAGUA
# aQBuACAAYgB5ACAAcgBlAGYAZQByAGUAbgBjAGUALjALBglghkgBhv1sAxUwEgYD
# VR0TAQH/BAgwBgEB/wIBADB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGGGGh0
# dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2NhY2Vy
# dHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDCBgQYD
# VR0fBHoweDA6oDigNoY0aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0
# QXNzdXJlZElEUm9vdENBLmNybDA6oDigNoY0aHR0cDovL2NybDQuZGlnaWNlcnQu
# Y29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDAdBgNVHQ4EFgQUFQASKxOY
# spkH7R7for5XDStnAs0wHwYDVR0jBBgwFoAUReuir/SSy4IxLVGLp6chnfNtyA8w
# DQYJKoZIhvcNAQEFBQADggEBAEZQPsm3KCSnOB22WymvUs9S6TFHq1Zce9UNC0Gz
# 7+x1H3Q48rJcYaKclcNQ5IK5I9G6OoZyrTh4rHVdFxc0ckeFlFbR67s2hHfMJKXz
# BBlVqefj56tizfuLLZDCwNK1lL1eT7EF0g49GqkUW6aGMWKoqDPkmzmnxPXOHXh2
# lCVz5Cqrz5x2S+1fwksW5EtwTACJHvzFebxMElf+X+EevAJdqP77BzhPDcZdkbkP
# Z0XN1oPt55INjbFpjE/7WeAjD9KqrgB87pxCDs+R1ye3Fu4Pw718CqDuLAhVhSK4
# 6xgaTfwqIa1JMYNHlXdx3LEbS0scEJx3FMGdTy9alQgpECYxggQ7MIIENwIBATCB
# hjByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQL
# ExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3Vy
# ZWQgSUQgQ29kZSBTaWduaW5nIENBAhAMNC3tjVe6FtjxG5RahkLeMAkGBSsOAwIa
# BQCgeDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgor
# BgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3
# DQEJBDEWBBTp9C2yKTBYOzYQpz5piOOeu1/i8jANBgkqhkiG9w0BAQEFAASCAQBP
# cA9GF2qPYU32QYQ8EK21nACE+dpoWa1IlXjL0jClLT87GiumSYVd8lwQoa9tu9OS
# S1XCRpUiMHNO0iHTS60Kdt3FM/rSMl4T8zB36Ct+jKG49dZGRA3E1YIvNsTtZCJV
# aYKiq1tGBhQmruo9QfSXTtrxuMYMqsrKp/x4cL1oFw4FZtYHhhCubFkjmqBlQVnM
# edP+1xO+VsoADC1iMsh5S9+rrLLhSG66w/GcfI3pHxq+zBEPqZuF6a3aJsd5yEr1
# QPtHMPXL9rO1U+satjOzD4Lc7JeIo9C2Vp+BrtwwvsK4iHgyV239fA8YK+irXnx5
# hmuFGAmCIA0MGgwEUKQOoYICDzCCAgsGCSqGSIb3DQEJBjGCAfwwggH4AgEBMHYw
# YjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQ
# d3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgQXNzdXJlZCBJRCBD
# QS0xAhADAZoCOv9YsWvW1ermF/BmMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMx
# CwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xNTA3MjIxNTI0MzRaMCMGCSqG
# SIb3DQEJBDEWBBR9op6akJ9XqD9D2I5lrb90VnhctjANBgkqhkiG9w0BAQEFAASC
# AQBBETI2KSueQvlB54qYvt18djsmzkItrYxbqdUaR3eb6xGBdXfSpdIczckDRH3E
# 4DRe2RnAVIeF6bXQrq6W6H2TP3LpndJyt1HtgeGS45FPgLSOlfcDmBneDeA8PpLP
# mrmYExb9I28byR1Rt9hMgwAlMKdpEXDs8CD2ek2Oriotncl4qY8CU6aslyA2a2N6
# zhLAheGFUPxlw2JJKCoEDhU03AwW3ijNsnky1Zwg857BEZxRQEV/E0d5604+p8Ye
# 9pab+Ud0eQhAXvbhbP37FBvDx+AgZacrTrnVC8C+D3I8HDV39oQDx/eUHMgkRlju
# L8w+uRdm6pcRJIsfajip7dNA
# SIG # End signature block
