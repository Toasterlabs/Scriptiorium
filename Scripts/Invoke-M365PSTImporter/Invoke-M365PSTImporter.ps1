[CmdletBinding(DefaultParametersetName="Default")] 
Param(
	#region Switch Parameters
	[Parameter(Mandatory = $false, HelpMessage = 'Opens the compliance center', ParameterSetName='Default')]
	[Alias("OpenComplianceCenter","SCC")]
	[SWITCH]$ComplianceCenter,
	[Parameter(Mandatory = $false, HelpMessage = 'Opens webpage to the documentation of this script', ParameterSetName='Default')]
	[SWITCH]$Help,
	[Parameter(Mandatory = $false, HelpMessage = 'Upload PST Files', ParameterSetName='UploadPST')]
	[SWITCH]$UploadPST,
	#endregion
	#region Log Parameters
	[Parameter(Mandatory = $false, HelpMessage = 'Runlog output location')]
	[Alias("Log")]
	[ValidateNotNullOrEmpty()]
	[System.IO.FileInfo]$Runlog,
	#endregion
	#region AZCopy Parameters
	[Parameter(Mandatory = $True, HelpMessage = 'Path to iterate through looking for PST Files', ParameterSetName='UploadPST')]
	[Alias("PSTPath")]
	[ValidateNotNullOrEmpty()]
	$Path,
	[Parameter(Mandatory = $True, HelpMessage = 'Azure Destination URI', ParameterSetName='UploadPST')]
	[Alias("Destination","Dest")]
	[ValidateNotNullOrEmpty()]
	$SASURL,
	[Parameter(Mandatory = $false, HelpMessage = 'Look for PST files in subfolders', ParameterSetName='UploadPST')]
	[SWITCH]$Recurse,
	[Parameter(Mandatory = $false, HelpMessage = 'Pattern (Set to *.PST)', ParameterSetName='UploadPST')]
	[SWITCH]$Pattern,
	[Parameter(Mandatory = $false, HelpMessage = 'Output AZCopy logging to the screen', ParameterSetName='UploadPST')]
	[SWITCH]$AZCopyVerbose,
	[Parameter(Mandatory = $False, HelpMessage = 'Location of mapping file', ParameterSetName='UploadPST')]
	[ValidateNotNullOrEmpty()]
	[SWITCH]$MappingFile,
	[Parameter(Mandatory = $False, HelpMessage = 'EmailAddress of the user which is being processed', ParameterSetName='UploadPST')]
	[ValidateNotNullOrEmpty()]
	[SWITCH]$EmailAddress
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
		Invoke-Logging -LogLevel SUCCESS -Message "Running on $($OS.Productname)" -Runlog $Runlog
	}Else{
		Invoke-Logging -LogLevel ALERT -Message "We must run on Windows 10 or credentials saving will not work" -Runlog $Runlog
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
		Invoke-Logging -LogLevel INFO -Message "Opening M365 Security & Compliance Center WebPage" -Runlog $Runlog
		Start-Process -FilePath https://protection.office.com
		Invoke-Logging -LogLevel TEXT -Message "1. Open the Data Governance"
		Invoke-Logging -LogLevel TEXT -Message "2. Select the Import Menu"
		Invoke-Logging -LogLevel TEXT -Message "3. Click on the + New import job button"
		Invoke-Logging -LogLevel TEXT -Message "4. Name the job"
		Invoke-Logging -LogLevel TEXT -Message "5. Hit Next"
		Invoke-Logging -LogLevel TEXT -Message "6. Select Upload your Data"
		Invoke-Logging -LogLevel TEXT -Message "7. Hit Next"
		Invoke-Logging -LogLevel TEXT -Message "8. Click on Show network Upload SAS URL & save it"
	}

	## Opening Help WebPage
	If($help){
		Invoke-Logging -LogLevel INFO -Message "Opening help WebPage"
		Start-Process -FilePath "ADD LINK LATER"
	}

	## AZ Copy Options Array
	if($UploadPST){
		
		# Arrays
		$OptionArray = @()
		$MappingArray = @()

		### Go through all subfolders 
		$PSTFiles = Get-ChildItem -Path $Path -Filter "*.PST"
		If($Recurse){
			Invoke-Logging -LogLevel INFO -Message "Adding Recursive Switch" -Runlog $runlog
			$OptionArray += "/S"
			$PSTFiles = Get-ChildItem -Path $Path -Filter "*.PST" -Recurse
		}

		### Use the .pst Pattern
		If($Pattern){
			Invoke-Logging -LogLevel INFO -Message "Adding Pattern Switch" -Runlog $runlog
			$OptionArray += "/Pattern=*.PST"
		}

		### AZ Copy logging option
		If($AZCopyVerbose){
			Invoke-Logging -LogLevel INFO -Message "Adding Verbose Switch" -Runlog $runlog
			$OptionArray += "/V"
		}Else{
			Invoke-Logging -LogLevel INFO -Message "AZCopy output will be logged to $(Get-Date -uformat %Y%m%d-%H%M)-AZCopyVerbose.log" -Runlog $runlog
			$OptionArray += "/V: $ScriptDir\$(Get-Date -uformat %Y%m%d-%H%M)-AZCopyVerbose.log"
		}

		### Allowing Write only Tokens
		$OptionArray += '/Y'

		### Joining array
		$Options = [system.string]::Join(" ",$OptionArray)
	
		### Starting AZCopy Process
		Invoke-Logging -LogLevel INFO -Message "Starting AZCopy..." -Runlog $runlog
		Start-Process -FilePath "$ProgramFiles\Microsoft SDKs\Azure\AZCopy\AZCopy.exe" -ArgumentList "/Source:`"$Path`" /Dest:$SASURL $Options" -PassThru

		### Creating Mapping file if it doesn't Exist
		If(Test-Path $MappingFile){
			Invoke-Logging -LogLevel SUCCESS -Message "Mapping file found! We'll be adding the information to this file..." -runlog $Runlog
		}Else{
			Invoke-Logging -LogLevel ALERT -Message "Mapping file not found! Creating new mapping file" -runlog $Runlog
			New-Item -ItemType File -Path $MappingFile | Out-Null

			#### Creating Headers
			$Workload = "Workload"
			$Destination = 'FilePath'
			$Name = 'Name'
			$mailbox = 'Mailbox'
			$IsArchive = 'IsArchive'
			$TargetRootFolder = 'TargetRootFolder'
			$SPFileContainer = 'SPFileContainer'
			$SPManifestContainer = 'SPManifestContainer'
			$SPSiteUrl = 'SPSiteURL'

			#### Creating Line
			$NewLine =  "{0},{1},{2},{3},{4},{5},{6},{7},{8}" -f $Workload,$Destination,$Name,$mailbox,$IsArchive,$TargetRootFolder,$SPFileContainer,$SPManifestContainer,$SPSiteUrl
			
			#### Adding line to file
			$NewLine | Add-Content -Path $MappingFile 
		}

		## Doing PST File Stuff
		Foreach($PST in $PSTFiles){
			#### Populating
			$Workload = "EXCHANGE"
			$Destination = ''
			$Name = $PST.Name
			$mailbox = $EmailAddress
			$IsArchive = 'FALSE'
			$TargetRootFolder = '/'
			$SPFileContainer = ''
			$SPManifestContainer = ''
			$SPSiteUrl = ''

			#### Creating Line
			$NewLine =  "{0},{1},{2},{3},{4},{5},{6},{7},{8}" -f $Workload,$Destination,$Name,$mailbox,$IsArchive,$TargetRootFolder,$SPFileContainer,$SPManifestContainer,$SPSiteUrl
			
			#### Adding line to file
			$NewLine | Add-Content -Path $MappingFile 
		}
	}
#endregion