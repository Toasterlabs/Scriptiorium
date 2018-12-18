Param(
	#region Storage Blob creation
	[Parameter(Mandatory = $false, HelpMessage = 'Calls code to make a new Azure storage blob', ParameterSetName='StorageBlob')]
	[SWITCH]$ConfigureBlob,
	[Parameter(Mandatory = $false, HelpMessage = 'Storage Account Name: <contoso>.blob.core.windows.net', ParameterSetName='StorageBlob')]
	$StorageAccountName,
	[Parameter(Mandatory = $false, HelpMessage = 'Date the SAS token expires', ParameterSetName='StorageBlob')]
	$TokenExpireDate,
	[Parameter(Mandatory = $false, HelpMessage = 'Part after the Storage account name: <contoso>.blob.core.windows.net/<StorageAccountContainer>', ParameterSetName='StorageBlob')]
	$StorageAccountContainer,
	[Parameter(Mandatory = $false, HelpMessage = 'Location for the Blob Storage', ParameterSetName='StorageBlob')]
	[ValidateSet("West Europe","North Europe","East US 2","Central US","South Central US","West US","East US","Southeast Asia","East Asia","Japan West")]
	$StorageLocation,
	#endregion
	#region Upload PST
	[Parameter(Mandatory = $false, HelpMessage = 'Upload PST Files', ParameterSetName='UploadPST')]
	[SWITCH]$UploadPST,
	[Parameter(Mandatory = $false, HelpMessage = 'Azure blob URI', ParameterSetName='UploadPST')]
	[parameter(ParameterSetName = 'ImportPST')]
	$AzureBlobStorageAccountUri,
	[Parameter(Mandatory = $false, HelpMessage = 'Signature Token', ParameterSetName='UploadPST')]
	[parameter(ParameterSetName = 'ImportPST')]
	$AzureSharedAccessSignatureToken,
	[Parameter(Mandatory = $false, HelpMessage = 'Look for PST files in subfolders', ParameterSetName='UploadPST')]
	[SWITCH]$Recurse,
	[Parameter(Mandatory = $false, HelpMessage = 'Pattern (Set to *.PST)', ParameterSetName='UploadPST')]
	[SWITCH]$Pattern,
	[Parameter(Mandatory = $false, HelpMessage = 'Output AZCopy logging to the screen', ParameterSetName='UploadPST')]
	[SWITCH]$AZCopyVerbose,
	[Parameter(Mandatory = $True, HelpMessage = 'Path to iterate through looking for PST Files', ParameterSetName='UploadPST')]
	[Alias("PSTPath")]
	[ValidateNotNullOrEmpty()]
	$Path,
	[Parameter(Mandatory = $True, HelpMessage = 'Export path for the CSV Mapping', ParameterSetName='UploadPST')]
	[ValidateNotNullOrEmpty()]
	$CSVPath,
	[Parameter(Mandatory = $false, HelpMessage = 'User we are processing', ParameterSetName='UploadPST')]
	[Alias("Email")]
	[Alias("mail")]
	[Alias("UPN")]
	[SWITCH]$EmailAddress,
	#endregion
	#region Import PST
	[Parameter(Mandatory = $false, HelpMessage = 'Import PST Files', ParameterSetName='ImportPST')]
	[SWITCH]$ImportPST,
	[Parameter(Mandatory = $false, HelpMessage = 'Where the PST should be imported', ParameterSetName='ImportPST')]
	$TargetRootFolder,
	[Parameter(Mandatory = $false, HelpMessage = 'How many items can be skipped', ParameterSetName='ImportPST')]
	$BadItemLimit,
	[Parameter(Mandatory = $false, HelpMessage = 'Name for the import', ParameterSetName='ImportPST')]
	$BatchName,
	[Parameter(Mandatory = $false, HelpMessage = 'Folders to exclude', ParameterSetName='ImportPST')]
	$ExcludeFolders,
	#endregion
	[Parameter(Mandatory = $false, HelpMessage = 'Runlog')]
	$runlog
)

# Variables
$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path
$Module = -join($ScriptDir,"\Module\Invoke-M365PSTImport.psm1")

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

# Header
## Clearing Screen
cls
## Output
Invoke-Logging -LogLevel Title -Message "Microsoft 365 PST Uploader & Importer starting!" -Runlog $Runlog
Write-Output ""
Invoke-Logging -LogLevel INFO -Message "RUnlog will be saved in $runlog"

#region Testing Prerequisites
	Invoke-Logging -LogLevel INFO -Message "Prerequisite testing starting..." -Runlog $Runlog
	
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
		0 {Install-Module AzureAD}
		1 {Exit}
		}
	}

	# Test for Azure module
	If (!(Get-ChildItem -Path 'C:\Program Files\WindowsPowerShell\Modules\AzureAD\' -Recurse -Filter "AzureAD.psd1" -ErrorAction SilentlyContinue)){
		Invoke-Logging -LogLevel ALERT -message "Microsoft Azure AD Module is not installed" -Runlog $Runlog
		$Title = "Install Module?"
		$Message = "Would you like to install the Azure AD module?"
		$Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Installs Azure AD Module"
		$No = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Exit this script"
		$Options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No)
		$Result = $Host.UI.PromptForChoice($Title, $Message, $Options, 0)
		Switch ($Result){
			0 {Install-Module AzureAD -Force}
			1 {Exit}
		}
	}Else{
		Invoke-Logging -LogLevel SUCCESS -message "Microsoft Azure AD Module is installed" -Runlog $Runlog
	}
	
	# Test for Azure RM module
	If (!(Get-ChildItem -Path 'C:\Program Files\WindowsPowerShell\Modules\Azure.Storage\' -Recurse -Filter "Azure.Storage.psd1" -ErrorAction SilentlyContinue)){
		Invoke-Logging -LogLevel ALERT -message "Microsoft Azure RM Module is not installed" -Runlog $Runlog
		$Title = "Install Module?"
		$Message = "Would you like to install the Azure RM module?"
		$Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Installs Azure RM Module"
		$No = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Exit this script"
		$Options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No)
		$Result = $Host.UI.PromptForChoice($Title, $Message, $Options, 0)
		Switch ($Result){
			0 {Install-Module AzureRM -Force}
			1 {Exit}
		}
	}Else{
		Invoke-Logging -LogLevel SUCCESS -message "Microsoft Azure RM Module is installed" -Runlog $Runlog
	}

#endregion

#region Upload PST files
If($UploadPST){
	
	# Header
	Write-Output ""
	Invoke-Logging -LogLevel INFO -Message "Starting PST Upload Process"

	# Go through all subfolders 
	$PSTFiles = Get-ChildItem -Path $Path -Filter "*.PST"

	Foreach($PST in $PSTFiles){
		$CSV = [PSCustomObject]
		$CSV = "" | Select email,PSTName

		$CSV.Email = $EmailAddress
		$CSV.PSTName = $PST.Name

		$CSV | Export-Csv -Path "$CSVPath\Mapping.csv" -NoTypeInformation -Append
	}
	
	If($Recurse){
		Invoke-Logging -LogLevel INFO -Message "Adding Recursive Switch" -Runlog $runlog
		$OptionArray += "/S"
		$PSTFiles = Get-ChildItem -Path $Path -Filter "*.PST" -Recurse
	}

	# Use the .pst Pattern
	If($Pattern){
		Invoke-Logging -LogLevel INFO -Message "Adding Pattern Switch" -Runlog $runlog
		$OptionArray += "/Pattern=*.PST"
	}

	# AZ Copy logging option
	If($AZCopyVerbose){
		Invoke-Logging -LogLevel INFO -Message "Adding Verbose Switch" -Runlog $runlog
		$OptionArray += " /V"
	}Else{
		Invoke-Logging -LogLevel INFO -Message "AZCopy output will be logged to $(Get-Date -uformat %Y%m%d-%H%M)-AZCopyVerbose.log" -Runlog $runlog
		$OptionArray += " /V: $ScriptDir\$(Get-Date -uformat %Y%m%d-%H%M)-AZCopyVerbose.log"
		}

	# Allowing Write only Tokens
	$OptionArray += ' /Y'

	# Joining array
	$Options = [system.string]::Join(" ",$OptionArray)
	
	# Starting AZCopy Process
	Invoke-Logging -LogLevel INFO -Message "Starting AZCopy..." -Runlog $runlog
	Start-Process -FilePath "$ProgramFiles\Microsoft SDKs\Azure\AZCopy\AZCopy.exe" -ArgumentList "/Source:`"$Path`" /Dest:$AzureBlobStorageAccountUri /DestKey:$AzureSharedAccessSignatureToken $Options" -PassThru -NoNewWindow -Wait
}

#endregion

#region Azure Storage Blob
If($ConfigureBlob){
	# Header
	Write-Output ""
	Invoke-Logging -LogLevel INFO -Message "Starting Azure Storage Blob configuration Process"

	# Names must be in lowercase or Azure RM commands throw a hissyfit
	$StorageAccountName = $StorageAccountName.ToLower()
	$StorageAccountContainer = $StorageAccountContainer.ToLower()

	# Importing Azure RM module
	Invoke-Logging -LogLevel INFO -Message "Importing Azure RM module"
	Import-Module AzureRM

	# Connecting to Azure
	Invoke-Logging -LogLevel INFO -Message "Connecting to Azure"
	$Connection = connect-azureRmAccount

	# Getting subscription name
	$SubscriptionName = $connection.Context.Subscription.Name

	# Creating a new resourceGroup
	$ResourceGroup = "PSTUploads"
	invoke-logging -LogLevel INFO -Message "Creating resource group: PSTUploads"
	New-AzureRmResourceGroup -Name $resourceGroup -Location $StorageLocation  | Out-Null

	# Creating Storage account
	Invoke-Logging -LogLevel INFO -Message "Creating new storage Account"
	New-AzureRMStorageAccount -ResourceGroupName $resourceGroup -Name $StorageAccountName -Location $StorageLocation -SkuName "Standard_LRS" | Out-Null
	
	# Sleeping for 3 seconds to allow things to catch up
	Start-Sleep -Seconds 3

	# Setting subscription
	Invoke-Logging -LogLevel INFO -Message "Setting storage account context"
	Set-AzureRmCurrentStorageAccount -ResourceGroupName $ResourceGroup -Name $StorageAccountName

	# Creating storage containter
	Invoke-Logging -LogLevel INFO -Message "Creating new storage container"
	New-AzureStorageContainer -Name $StorageAccountContainer -Permission Off | Out-null

	# Creating Token For importing PST
	Invoke-Logging -LogLevel INFO -Message "Creating Token for storage account container."
	$token = New-AzureStorageContainerSASToken -Name $StorageAccountContainer -Permission rwl -ExpiryTime $TokenExpireDate

	# Getting storage key for AZCopy
	$StorageKey = Get-AzureRmStorageAccountKey -ResourceGroupName $resourceGroup -AccountName $StorageAccountName

	# Reporting Back
	Invoke-Logging -LogLevel ALERT -Message "Save the below information! You will need it when running the PST Upload section of this script!"
	Write-Output "`nStorage Account URI`n-------------------`nhttps://$StorageAccountName.blob.core.windows.net/$StorageAccountContainer`n`n"
	Write-Output "SAS Token (Importing PST)`n-------------------------`n$token`n`n"
	Write-Output "Storage Key (AZCopy)`n--------------------`n$($storagekey[0].value)`n`n"
	Write-Output "SAS URI`n-------`nhttps://$StorageAccountName.blob.core.windows.net/$StorageAccountContainer$token`n"
}

#endregion

#region Import PSTs
#endregion