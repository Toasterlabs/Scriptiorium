Param(
	[Parameter(Mandatory = $false, HelpMessage = 'Path to save the runlog')]
	$runlog
)

# Variables
	# Checking if Runlog is empty
	If(!($Runlog)){
		$Runlog = "\$(Get-Date -uformat %Y%m%d-%H%M)-PSTImporter.log"
	}Else{
		$Runlog += "\$(Get-Date -uformat %Y%m%d-%H%M)-PSTImporter.log"
	}

	#region Synchronized Collections
	$Global:SyncHash = [hashtable]::Synchronized(@{})
	$Global:VariableProxy = [hashtable]::Synchronized(@{})

	$Global:VariableProxy.ScriptDir = $PSScriptRoot
	$Global:VariableProxy.Module = -join($PSScriptRoot,"\Module\Invoke-M365PSTImport.psm1")
	$Global:VariableProxy.Runlog = -join($PSScriptRoot,$runlog)
	$VariableProxy.ResourceGroup = "pstuploads"
	$VariableProxy.MappingFile = -join($PSScriptRoot,"\Mapping.csv")
	$VariableProxy.SettingsFile = -join($PSScriptRoot,"\Settings.xml")


# Importing Module
If(Test-Path $VariableProxy.Module){
	Import-Module $VariableProxy.Module -DisableNameChecking -Force
}Else{
	Write-Warning -Message "Unable to import script module. Now Exiting on the right!"
	exit
}

#region Header
	## Clearing Screen
	cls

	## Output
	Invoke-Logging -LogLevel Title -Message "Microsoft 365 PST Uploader & Importer (Graphical Edition) starting!" -Runlog $VariableProxy.Runlog
	Write-Output ""
	Invoke-Logging -LogLevel INFO -Message "RUnlog will be saved in $($VariableProxy.Runlog)"
#endregion

#region Testing Prerequisites
	Invoke-Logging -LogLevel INFO -Message "Prerequisite testing starting..." -Runlog $VariableProxy.Runlog
	
	## Internet connectivity
	$INETCONN = invoke-connectivityCheck -Internet
	If($INETCONN.Internet -eq $true){
		Invoke-Logging -LogLevel SUCCESS -Message 'Internet connectivity check passed!' -runlog $VariableProxy.Runlog
	}Else{
		Invoke-Logging -LogLevel ALERT -Message 'Internet connectivity check failed!' -runlog $VariableProxy.Runlog
		exit
	}

	## OS info
	$VariableProxy.OS = Get-OSVersion
	If($VariableProxy.OS.Productname -like "*10*"){
		Invoke-Logging -LogLevel SUCCESS -Message "Running on $($VariableProxy.OS.Productname)" -Runlog $VariableProxy.Runlog
	}Else{
		Invoke-Logging -LogLevel ALERT -Message "We must run on Windows 10 or credentials saving will not work" -Runlog $VariableProxy.Runlog
		Exit
	}

	## Program Files location setting
	If($VariableProxy.OS.'64Bit' -like "True"){
		$VariableProxy.Programfiles = [Environment]::GetEnvironmentVariable("ProgramFiles(x86)")
	}Else{
		$VariableProxy.Programfiles = [Environment]::GetEnvironmentVariable("ProgramFiles")
	}

	## AZCopy 
	$AZCopyPresent = Test-Path "$($VariableProxy.ProgramFiles)\Microsoft SDKs\Azure\AZCopy\AZCopy.exe"
	If ($AZCopyPresent -ne "True") {
	Invoke-Logging -LogLevel ALERT -Message "Unable to find AZCopy.exe in the default path.." -Runlog $VariableProxy.Runlog
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
	}Else{
		Invoke-Logging -LogLevel SUCCESS -Message "AZCopy installation found!" -Runlog $VariableProxy.Runlog
	}

	# Test for Azure module
	If (!(Get-ChildItem -Path 'C:\Program Files\WindowsPowerShell\Modules\AzureAD\' -Recurse -Filter "AzureAD.psd1" -ErrorAction SilentlyContinue)){
		Invoke-Logging -LogLevel ALERT -message "Microsoft Azure AD Module is not installed" -Runlog $VariableProxy.Runlog
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
		Invoke-Logging -LogLevel SUCCESS -message "Microsoft Azure AD Module is installed" -Runlog $VariableProxy.Runlog
	}
	
	# Test for Azure RM module
	If (!(Get-ChildItem -Path 'C:\Program Files\WindowsPowerShell\Modules\Azure.Storage\' -Recurse -Filter "Azure.Storage.psd1" -ErrorAction SilentlyContinue)){
		Invoke-Logging -LogLevel ALERT -message "Microsoft Azure RM Module is not installed" -Runlog $VariableProxy.Runlog
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
		Invoke-Logging -LogLevel SUCCESS -message "Microsoft Azure RM Module is installed" -Runlog $VariableProxy.Runlog
	}

	# Testing Mapping file
	If(!(Test-path $VariableProxy.MappingFile)){
		$CSV = [PSCustomObject]
		$CSV = "" | Select email,PSTName, Status
		$CSV | Export-Csv -Path $VariableProxy.MappingFile -NoTypeInformation
	}
#endregion

#region Main
	# setting up
	$syncHash.host = $Host
	$newRunspace =[runspacefactory]::CreateRunspace()
	$newRunspace.ApartmentState = "STA"
	$newRunspace.ThreadOptions = "ReuseThread"
	$newRunspace.Open()
	# Syncing variables
	$newRunspace.SessionStateProxy.SetVariable("syncHash",$syncHash)
	$newRunspace.SessionStateProxy.SetVariable("VariableProxy",$VariableProxy)
	#endregion

	# Scriptblock to execute in a runspace
	$psCmd = [PowerShell]::Create().AddScript({
	
		#region Background runspace to clean up jobs
		$jobCleanup.Flag = $True
		$newRunspace =[runspacefactory]::CreateRunspace()
		$newRunspace.ApartmentState = "STA"
		$newRunspace.ThreadOptions = "ReuseThread"
		$newRunspace.Open()
		$newRunspace.SessionStateProxy.SetVariable("synchash",$synchash)
		$newRunspace.SessionStateProxy.SetVariable("jobCleanup",$jobCleanup)
		$newRunspace.SessionStateProxy.SetVariable("jobs",$jobs)		
		#endregion

		# Import Module

		#region GUI Prep
			# Load Required Assemblies
			Add-Type –assemblyName PresentationFramework

			# Loading XAML code	
			## Reading XAML file	
			[xml]$xaml = Get-Content -Path "$($VariableProxy.ScriptDir)\Forms\Main.xaml"
			## Loading in to XML Node reader
			$reader = (New-Object System.Xml.XmlNodeReader $xaml)
			## Loading XML Node reader in to $SyncHash window property to launch later
			$SyncHash.Window = [Windows.Markup.XamlReader]::Load($reader)
		#endregion

		#region Connecting Controls
			# Configure Blob
			$synchash.TXT_StorageAccountName = $SyncHash.Window.FindName("TXT_StorageAccountName")
			$synchash.TXT_StorageAccountContainer = $SyncHash.Window.FindName("txt_StorageAccountContainer")
			$synchash.DatePicker_TokenExpireTime = $SyncHash.Window.FindName("DatePicker_TokenExpireTime")
			$synchash.DD_StorageLocation = $SyncHash.Window.FindName("DD_StorageLocation")
			$synchash.BTN_Create = $SyncHash.Window.FindName("BTN_Create")
			$synchash.OutputWindow = $SyncHash.Window.FindName("OutputWindow")

			# Upload PST
			$synchash.TXT_AccountURI = $SyncHash.Window.FindName("txt_AccountURI")
			$synchash.TXT_StorageKey = $SyncHash.Window.FindName("txt_StorageKey")
			$synchash.TXT_PSTPath = $SyncHash.Window.FindName("txt_PSTPath")
			$synchash.TXT_Email = $SyncHash.Window.FindName("txt_Email")
			$synchash.BTN_Browse = $SyncHash.Window.FindName("BTN_Browse")
			$synchash.BTN_Upload = $SyncHash.Window.FindName("BTN_Upload")
			$synchash.BTN_Import = $SyncHash.Window.FindName("BTN_Import")
			$synchash.BTN_Save = $SyncHash.Window.FindName("BTN_Save")

			# PST import
			$synchash.TXT_SASToken = $SyncHash.Window.FindName("txt_SASToken")
			$synchash.TXT_MappingFile = $SyncHash.Window.FindName("txt_MappingFile")
			$synchash.BTN_Import_PST = $SyncHash.Window.FindName("BTN_Import_PST")
		#endregion

		# Import module
		import-module $VariableProxy.Module -force -DisableNameChecking

		# Adding items to the ComboBox (Storage Locations)
		$synchash.DD_StorageLocation.items.Add("West Europe")
		$synchash.DD_StorageLocation.items.Add("North Europe")
		$synchash.DD_StorageLocation.items.Add("Central US")
		$synchash.DD_StorageLocation.items.Add("East US")
		$synchash.DD_StorageLocation.items.Add("East US2")
		$synchash.DD_StorageLocation.items.Add("North Central US")
		$synchash.DD_StorageLocation.items.Add("Soutch Central US")
		$synchash.DD_StorageLocation.items.Add("West US")
		$synchash.DD_StorageLocation.items.Add("East Asia")
		$synchash.DD_StorageLocation.items.Add("Southeast Asia")
		$synchash.DD_StorageLocation.items.Add("Japan West")
		$synchash.DD_StorageLocation.items.Add("Japan East")
		$synchash.DD_StorageLocation.items.Add("Brazi South")
		$synchash.DD_StorageLocation.items.Add("Australia East")
		$synchash.DD_StorageLocation.items.Add("Australia Southeast")
		$synchash.DD_StorageLocation.items.Add("Central India")
		$synchash.DD_StorageLocation.items.Add("South India")
		$synchash.DD_StorageLocation.items.Add("West India")

		# Loaded!
		$Timestamp = (get-date -Format HH:mm:ss)
		$Message = "$TimeStamp - [INFO]: User Interface Loaded"
		$Synchash.OutputWindow.AppendText("$Message")

		#region Event Handlers

			# Create Button
			$Synchash.BTN_Create.Add_Click({
				# Capturing Values
				$VariableProxy.TokenExpiry = Get-Date($Synchash.DatePicker_TokenExpireTime.Text) -Format 'yyyy-MM-dd'
				$VariableProxy.StorageLocation = ($Synchash.DD_StorageLocation.Items.CurrentItem).ToLower()
				$VariableProxy.StorageAccountName = ($synchash.TXT_StorageAccountName.Text).ToLower()
				$VariableProxy.StorageAccountContainer = ($synchash.TXT_StorageAccountContainer.Text).ToLower()
				
				# Reporting
				$Timestamp = (get-date -Format HH:mm:ss)
				$Message = "$TimeStamp - [INFO]: Initiating creation of storage space`nStorage Region: $($VariableProxy.StorageLocation)`nToken Expires: $($VariableProxy.TokenExpiry)`nStorage URI: https://$($VariableProxy.StorageAccountName).blob.core.windows.net/$($VariableProxy.StorageAccountContainer)"
				$Synchash.OutputWindow.AppendText("`n")
				$Synchash.OutputWindow.AppendText("$Message")

				# Runspace building (If we don't run the code in a seperate runspace the GUI will lock up. Nobody like a GUI that locks up!)
                $newRunspace =[runspacefactory]::CreateRunspace()
                $newRunspace.ApartmentState = "STA"
                $newRunspace.ThreadOptions = "ReuseThread"
                $newRunspace.Open()
                $newRunspace.SessionStateProxy.SetVariable("synchash",$synchash)
                $newRunspace.SessionStateProxy.SetVariable("VariableProxy",$VariableProxy)

				# Scriptblock
                $powershell = [PowerShell]::Create().AddScript({

					# Import module
					$Timestamp = (get-date -Format HH:mm:ss)
					Update-Control -Synchash $Synchash -Control "OutputWindow" -value " $Timestamp - [INFO]: Configuring Session" -AppendContent
					import-module $VariableProxy.Module -force -DisableNameChecking

					# Header
					$Timestamp = (get-date -Format HH:mm:ss)
					Update-Control -Synchash $Synchash -Control "OutputWindow" -value "$Timestamp - [INFO]: Starting Azure Storage Blob configuration Process" -AppendContent

					# Importing Azure RM module
					$Timestamp = (get-date -Format HH:mm:ss)
					Update-Control -Synchash $Synchash -Control "OutputWindow" -value "$Timestamp - [INFO]: Importing Azure RM module" -AppendContent
					Import-Module AzureRM

					# Connecting to Azure
					$Timestamp = (get-date -Format HH:mm:ss)
					Update-Control -Synchash $Synchash -Control "OutputWindow" -value "$Timestamp - [INFO]: Connecting to Azure Resource Manager" -AppendContent
					$Connection = connect-azureRmAccount
					
					# Getting subscription name
					$Timestamp = (get-date -Format HH:mm:ss)
					Update-Control -Synchash $Synchash -Control "OutputWindow" -value "$Timestamp - [INFO]: Retrieving Subscription Name" -AppendContent
					$SubscriptionName = $connection.Context.Subscription.Name

					# Creating a new resourceGroup
					$Timestamp = (get-date -Format HH:mm:ss)
					Update-Control -Synchash $Synchash -Control "OutputWindow" -value "$Timestamp - [INFO]: Creating resource group: PSTUploads" -AppendContent
					New-AzureRmResourceGroup -Name $VariableProxy.resourceGroup -Location $VariableProxy.StorageLocation  | Out-Null

					# Creating Storage account
					$Timestamp = (get-date -Format HH:mm:ss)
					Update-Control -Synchash $Synchash -Control "OutputWindow" -value "$Timestamp - [INFO]: Creating new storage account" -AppendContent
					New-AzureRMStorageAccount -ResourceGroupName $VariableProxy.resourceGroup -Name $VariableProxy.StorageAccountName -Location $VariableProxy.StorageLocation -SkuName "Standard_LRS" | Out-Null
	
					# Sleeping for 3 seconds to allow things to catch up
					$Timestamp = (get-date -Format HH:mm:ss)
					Update-Control -Synchash $Synchash -Control "OutputWindow" -value "$Timestamp - [INFO]: Wait while Azure configures the storage account" -AppendContent
					Start-Sleep -Seconds 3

					# Setting subscription
					$Timestamp = (get-date -Format HH:mm:ss)
					Update-Control -Synchash $Synchash -Control "OutputWindow" -value "$Timestamp - [INFO]: Setting storage account context" -AppendContent
					Set-AzureRmCurrentStorageAccount -ResourceGroupName $VariableProxy.ResourceGroup -Name $VariableProxy.StorageAccountName

					# Creating storage containter
					$Timestamp = (get-date -Format HH:mm:ss)
					Update-Control -Synchash $Synchash -Control "OutputWindow" -value "$Timestamp - [INFO]: Creating new storage container" -AppendContent
					New-AzureStorageContainer -Name $VariableProxy.StorageAccountContainer -Permission Off | Out-null

					# Creating Token For importing PST
					$Timestamp = (get-date -Format HH:mm:ss)
					Update-Control -Synchash $Synchash -Control "OutputWindow" -value "$Timestamp - [INFO]: Creating Token for storage account container" -AppendContent
					$token = New-AzureStorageContainerSASToken -Name $VariableProxy.StorageAccountContainer -Permission rwl -ExpiryTime $VariableProxy.TokenExpiry

					# Getting storage key for AZCopy
					$Timestamp = (get-date -Format HH:mm:ss)
					Update-Control -Synchash $Synchash -Control "OutputWindow" -value "$Timestamp - [INFO]: Retrieving Storage Account Key" -AppendContent
					$StorageKey = Get-AzureRmStorageAccountKey -ResourceGroupName $VariableProxy.resourceGroup -AccountName $VariableProxy.StorageAccountName
					
					# Adding Data
					$VariableProxy.StorageKey = $storagekey[0].value
					$VariableProxy.SASURI = "https://$($VariableProxy.StorageAccountName).blob.core.windows.net/$($VariableProxy.StorageAccountContainer)$token"
					$VariableProxy.AccountURI = "https://$($VariableProxy.StorageAccountName).blob.core.windows.net/$($VariableProxy.StorageAccountContainer)"
					$VariableProxy.SASToken = $token

					# Adding to GUI
					## Upload PST Tab
					Update-Control -Synchash $Synchash -Control txt_AccountURI -Property 'Text' -Value $VariableProxy.AccountURI
					Update-Control -Synchash $Synchash -Control txt_StorageKey -Property 'Text' -Value $VariableProxy.StorageKey

					## Import PST Tab
					Update-Control -Synchash $Synchash -Control txt_SASToken -Property 'Text' -Value $VariableProxy.SASToken
					Update-Control -Synchash $Synchash -Control txt_MappingFile -Property 'Text' -Value $VariableProxy.MappingFile

					# Reporting end
					Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Done" -AppendContent
				
				})

				# Invoking Scriptblock
				$powershell.Runspace = $newRunspace
                $data = $powershell.BeginInvoke()
								
			})
			
			$synchash.BTN_Browse.Add_Click({
				 # Location dialog for selecting folder
				$FolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
				$null = $FolderDialog.ShowDialog()

				# Setting the outputpath
				$SyncHash.TXT_PSTPath.text = $FolderDialog.SelectedPath

				# Reporting
				$Timestamp = (get-date -Format HH:mm:ss)
				$Message = "$TimeStamp - [INFO]: PST Path changed"
				$Synchash.OutputWindow.AppendText("`n")
				$Synchash.OutputWindow.AppendText("$Message")
			})
	
			# Import Settings File
			$synchash.BTN_Import.Add_Click({
				# Retrieving data
				$VariableProxy = Import-Clixml $VariableProxy.settingsFile

				# Adding to GUI
				## Configure Blob Tab
				$Synchash.TXT_StorageAccountName.Text = $VariableProxy.StorageAccountName
				$Synchash.TXT_StorageAccountContainer.Text = $VariableProxy.StorageAccountContainer
				$Synchash.DatePicker_TokenExpireTime.Text = $VariableProxy.TokenExpiry
				$Synchash.DD_StorageLocation.Text = $VariableProxy.StorageLocation

				## Upload PST Tab
				$Synchash.txt_AccountURI.Text = $VariableProxy.AccountURI
				$Synchash.txt_StorageKey.Text = $VariableProxy.StorageKey

				## Import PST Tab
				$Synchash.txt_SASToken.Text = $VariableProxy.SASToken
				$Synchash.txt_MappingFile.Text = $VariableProxy.MappingFile
				
				# Reporting end
				$Timestamp = (get-date -Format HH:mm:ss)
				update-control -synchash $Synchash -Control "OutputWindow" -value "$Timestamp - [INFO]: Import Completed" -AppendContent
			})
			
			# Export Settings File
			$synchash.BTN_Save.Add_Click({
				# Popup warning
				$Acceptance = new-popup -Title "Attention! Sensitive data warning!" -Message "Saving these settings potentially exposes sensitive data.`n It is your responsibility to ensure this file is secured appropriately!" -Buttons "OKCancel" -Icon "Exclamation" -DefaultButton "Second"

				If($Acceptance -eq "1"){
					# Exporting Settings
					$VariableProxy | Export-Clixml $VariableProxy.SettingsFile -Force

					# Reporting
					$Timestamp = (get-date -Format HH:mm:ss)
					$Message = "$TimeStamp - [INFO]: Settings saved."
					$Synchash.OutputWindow.AppendText("`n")
					$Synchash.OutputWindow.AppendText("$Message")
				}Else{
					# Reporting
					$Timestamp = (get-date -Format HH:mm:ss)
					$Message = "$TimeStamp - [INFO]: Cancelled saving the settings."
					$Synchash.OutputWindow.AppendText("`n")
					$Synchash.OutputWindow.AppendText("$Message")
				}
			})

			# Window closing
			$synchash.Window.add_Closing({
				# Cleaning Variables
				Clear-Variable Synchash
				Clear-Variable VariableProxy
			})

			# Upload PST files
			$synchash.BTN_Upload.Add_Click({
				
				# Capturing Values
				$VariableProxy.TokenExpiry = Get-Date($Synchash.DatePicker_TokenExpireTime.Text) -Format 'yyyy-MM-dd'
				$VariableProxy.StorageLocation = ($Synchash.DD_StorageLocation.Items.CurrentItem).ToLower()
				$VariableProxy.StorageAccountName = ($synchash.TXT_StorageAccountName.Text).ToLower()
				$VariableProxy.StorageAccountContainer = ($synchash.TXT_StorageAccountContainer.Text).ToLower()
				$VariableProxy.AccountURI = $Synchash.txt_AccountURI.Text
				$VariableProxy.StorageKey = $Synchash.txt_StorageKey.Text
				$VariableProxy.PSTPath = $SyncHash.TXT_PSTPath.text
				$VariableProxy.EmailAddress = $SyncHash.TXT_Email.Text

				# Runspace building (If we don't run the code in a seperate runspace the GUI will lock up. Nobody like a GUI that locks up!)
                $newRunspace =[runspacefactory]::CreateRunspace()
                $newRunspace.ApartmentState = "STA"
                $newRunspace.ThreadOptions = "ReuseThread"
                $newRunspace.Open()
                $newRunspace.SessionStateProxy.SetVariable("synchash",$synchash)
                $newRunspace.SessionStateProxy.SetVariable("VariableProxy",$VariableProxy)

				# Scriptblock
                $powershell = [PowerShell]::Create().AddScript({
					# Variables
					$Programfiles = [Environment]::GetEnvironmentVariable("ProgramFiles")

					# Import module
					$Timestamp = (get-date -Format HH:mm:ss)
					Update-Control -Synchash $Synchash -Control "OutputWindow" -value "$Timestamp - [INFO]: Configuring PST Upload session" -AppendContent
					import-module $VariableProxy.Module -force -DisableNameChecking

					# AZCopy Settings
					$Timestamp = (get-date -Format HH:mm:ss)
					Update-Control -Synchash $Synchash -Control "OutputWindow" -value "$Timestamp - [INFO]: Setting AZCopy options (Recurse, Verbose output logging, write only tokens), PST files only" -AppendContent
					$Timestamp = (get-date -Format HH:mm:ss)
					Update-Control -Synchash $Synchash -Control "OutputWindow" -value "$Timestamp - [INFO]: AZ Copy log is located in %LocalAppData%\Microsoft\Azure\AzCopy" -AppendContent
					$AZCOptions = " /S /V /Y /Pattern:*.pst"
					
					# Header
					$Timestamp = (get-date -Format HH:mm:ss)
					Update-Control -Synchash $Synchash -Control "OutputWindow" -value "$Timestamp - [INFO]: Starting PST upload process" -AppendContent

					# Go through all subfolders 
					$PSTFiles = Get-ChildItem -Path $VariableProxy.PSTPath -Filter "*.PST" -Recurse

					Foreach($PST in $PSTFiles){
						$Timestamp = (get-date -Format HH:mm:ss)
						Update-Control -Synchash $Synchash -Control "OutputWindow" -value "$Timestamp - [INFO]: Processing $($PST.Name)" -AppendContent

						$CSV = [PSCustomObject]
						$CSV = "" | Select email,PSTName, Status

						$CSV.Email = $VariableProxy.EmailAddress
						$CSV.PSTName = $PST.Name
						$CSV.Status = "Not Started"

						$CSV | Export-Csv -Path $VariableProxy.MappingFile -NoTypeInformation -Append
					}

					# Starting AZCopy Process
					$Timestamp = (get-date -Format HH:mm:ss)
					Update-Control -Synchash $Synchash -Control "OutputWindow" -value "$Timestamp - [INFO]: Starting PST upload (AZCopy)" -AppendContent
					Start-Process -FilePath "$($VariableProxy.ProgramFiles)\Microsoft SDKs\Azure\AZCopy\AZCopy.exe" -ArgumentList "/Source:`"$($VariableProxy.PSTPath)`" /Dest:$($VariableProxy.AccountURI) /DestKey:$($VariableProxy.StorageKey) $AZCOptions" -Wait

					# Report completed
					$Timestamp = (get-date -Format HH:mm:ss)
					Update-Control -Synchash $Synchash -Control "OutputWindow" -value "$Timestamp - [INFO]: PST upload completed (AZCopy)" -AppendContent
				})

				# Invoking Scriptblock
				$powershell.Runspace = $newRunspace
                $data = $powershell.BeginInvoke()
			})

			$synchash.BTN_Save.Add_Click({
				$VariableProxy.SASToken = $synchash.TXT_SASToken.Text
				$VariableProxy.MappingFile = $synchash.txt_MappingFile.Text

				Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Starting PST Import process" -AppendContent

				# Runspace building (If we don't run the code in a seperate runspace the GUI will lock up. Nobody like a GUI that locks up!)
                $PSTImportRunSpace =[runspacefactory]::CreateRunspace()
                $PSTImportRunSpace.ApartmentState = "STA"
                $PSTImportRunSpace.ThreadOptions = "ReuseThread"
                $PSTImportRunSpace.Open()
                $PSTImportRunSpace.SessionStateProxy.SetVariable("synchash",$synchash)
                $PSTImportRunSpace.SessionStateProxy.SetVariable("VariableProxy",$VariableProxy)

				$powershell = [PowerShell]::Create().AddScript({

					# Importing Mapping file
					$CSV = Import-Csv $VariableProxy.MappingFile

					# Reporting
					Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Connecting to Exchange Online" -AppendContent

					# Connecting to Exchange Online
					$UserCredential = Get-Credential
					$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
					Import-PSSession $Session -DisableNameChecking

					# iterate through the mapping file and start the import
					Foreach($i in $CSV){
						
						# Reporting
						Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Processing $($i.PSTName)" -AppendContent
						
						# Some checks
						If($i.status -eq "Started"){
							Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: $($i.PSTName) for $($i.email) has previously been processed." -AppendContent
						}Else{
							# Reporting
							Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Starting import creation of $($i.PSTName)." -AppendContent

							# Import request creation
							Try{
								New-MailboxImportRequest -Name $batchname -TargetRootFolder "Imported PST" -Mailbox $($i.email) -AzureBlobStorageAccountUri $VariableProxy.AccountUri/$i.PSTname -AzureSharedAccessSignatureToken $VariableProxy.SASToken
								($csv |where {$_.PSTName -eq $i.PSTName}).Status = "Started"
							}Catch{
								# Reporting
								Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[ERROR]: Something went wrong creating the import request for $($i.PSTName)." -AppendContent
								($csv |where {$_.PSTName -eq $i.PSTName}).Status = "ERROR"
							}
						}
					}
				})

				# Invoking Scriptblock
				$powershell.Runspace = $newRunspace
                $data = $powershell.BeginInvoke()
								

			})
		#endregion

		# Show GUI
		$SyncHash.Window.ShowDialog() | Out-Null
		$VariableProxy.Error = $Error
	
	})

#endregion


$psCmd.Runspace = $newRunspace
$data = $psCmd.BeginInvoke()