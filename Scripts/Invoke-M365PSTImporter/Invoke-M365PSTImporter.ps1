Param(
	#region Log Parameters
	[Parameter(Mandatory = $false, HelpMessage = 'Runlog output location')]
	[Alias("Log")]
	[System.IO.FileInfo]$Runlog,
	#endregion
	#region AZCopy Parameters
	[Parameter(Mandatory = $True, HelpMessage = 'Path to iterate through looking for PST Files')]
	[Alias("PSTPath")]
	[System.IO.FileInfo]$Path,
	[Parameter(Mandatory = $True, HelpMessage = 'Azure Destination URI')]
	[Alias("Destination","Dest")]
	$SASURL,
	[Parameter(Mandatory = $false, HelpMessage = 'Look for PST files in subfolders')]
	[SWITCH]$Recurse,
	[Parameter(Mandatory = $false, HelpMessage = 'Pattern (Set to *.PST)')]
	[SWITCH]$Pattern,
	[Parameter(Mandatory = $false, HelpMessage = 'Output AZCopy logging to the screen')]
	[SWITCH]$AZCopyVerbose,
	#endregion
	#region Switch Parameters
	[Parameter(Mandatory = $false, HelpMessage = 'Opens the compliance center')]
	[Alias("OpenComplianceCenter","SCC")]
	[SWITCH]$ComplianceCenter
	#endregion


)

# Variables
$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path
$Module = -join($ScriptDir,"\Module\O365PSTImport.psm1")

# Checking if Runlog is empty
If(!($Runlog)){
	$Runlog = ".\$(Get-Date -uformat %Y%m%d-%H%M)-ExchangeHybridTest.log"
}

# Importing Module
If(Test-Path $Module){
	Import-Module $Module -DisableNameChecking -Force
}Else{
	Write-Warning -Message "Unable to import script module. Now Exiting on the right!"
	exit
}

#region Testing Prerequisites
Invoke-Logging -LogLevel Title -Message "Prerequisite testing starting..." -Runlog $Runlog
	## Internet connectivity
	$INETCONN = invoke-connectivityCheck -Internet
	If($INETCONN.Internet -eq $true){
		Invoke-Logging -LogLevel SUCCESS -Message 'Internet connectivity check passed!' -runlog $Runlog
	}Else{
		Invoke-Logging -LogLevel ALERT -Message 'Internet connectivity check failed!' -runlog $Runlog
		exit
	}

	## OS info
	$OS = Get-OSVersion
	If($OS.Productname -like "*10*"){
		Invoke-Logging -LogLevel INFO -Message "Running on $($OS.Productname)"
	}Else{
		Invoke-Logging -LogLevel ALERT -Message "We must run on Windows 10 or credentials saving will not work"
		Exit
	}

	## Program Files location setting
	If($OS.'64Bit' -like "True"){
		$Programfiles = [Environment]::GetEnvironmentVariable("ProgramFiles(x86)")
	}Else{
		$Programfiles = [Environment]::GetEnvironmentVariable("ProgramFiles")
	}

	## AZCopy 
	$AZCopyPresent = Test-Path "$ProgramFiles\Microsoft SDKs\Azure\AZCopy\AZCopy.exe"
	If ($AZCopyPresent -ne "True") {
	Invoke-Logging -LogLevel ALERT -Message "Unable to find AZCopy.exe in the default path.."
	$Title = "Download AZCopy?"
	$Message = "Would you like to download a copy of AZCopy now?"
	$Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Download a copy of AZCopy from the Microsoft website now"
	$No = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Exit this script"
	$Options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No)
	$Result = $Host.UI.PromptForChoice($Title, $Message, $Options, 0)
	Switch ($Result){
		0 {Start-Process -FilePath http://aka.ms/downloadazcopy; Exit}
		1 {Exit}
		}
	}
#endregion

#region Main
	## Opening Security and compliance center
	If($ComplianceCenter){
		Invoke-Logging -LogLevel INFO -Message "Opening M365 Security & Compliance Center WebPage"
		Start-Process -FilePath https://protection.office.com
		Invoke-Logging -LogLevel INFO -Message "1. Open the Data Governance"
		Invoke-Logging -LogLevel INFO -Message "2. Select the Import Menu"
		Invoke-Logging -LogLevel INFO -Message "3. Click on the + New import job button"
		Invoke-Logging -LogLevel INFO -Message "4. Name the job"
		Invoke-Logging -LogLevel INFO -Message "5. Hit Next"
		Invoke-Logging -LogLevel INFO -Message "6. Select Upload your Data"
		Invoke-Logging -LogLevel INFO -Message "7. Hit Next"
		Invoke-Logging -LogLevel INFO -Message "8. Click on Show network Upload SAS URL & save it"
		Invoke-Logging -LogLevel INFO -Message "9. "
	}

	## AZ Copy Options Array
	$OptionArray = @()
	
	### Go through all subfolders 
	If($Recurse){
		Invoke-Logging -LogLevel INFO -Message "Adding Recursive Switch"
		$OptionArray += "/S"
	}

	### Use the .pst Pattern
	If($Pattern){
		Invoke-Logging -LogLevel INFO -Message "Adding Pattern Switch"
		$OptionArray += "/Pattern=*.PST"
	}

	### AZ Copy logging option
	If($AZCopyVerbose){
		Invoke-Logging -LogLevel INFO -Message "Adding Verbose Switch"
		$OptionArray += "/V"
	}Else{
		Invoke-Logging -LogLevel INFO -Message "AZCopy output will be logged to $(Get-Date -uformat %Y%m%d-%H%M)-AZCopyVerbose.log"
		$OptionArray += "/V: $ScriptDir\$(Get-Date -uformat %Y%m%d-%H%M)-AZCopyVerbose.log"
	}

	### Allowing Write only Tokens
	$OptionArray += '/Y'

	### Joining array
	$Options = [system.string]::Join(" ",$OptionArray)
	
	### Starting AZCopy Process
	Invoke-Logging -LogLevel INFO -Message "Starting AZCopy..."
	Start-Process -FilePath "$ProgramFiles\Microsoft SDKs\Azure\AZCopy\AZCopy.exe" -ArgumentList "/Source:`"$Path`" /Dest:$SASURL $Options" -PassThru

#endregion