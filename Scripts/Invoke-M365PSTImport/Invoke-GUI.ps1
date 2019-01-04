Param(
	[Parameter(Mandatory = $false, HelpMessage = 'Path to save the runlog')]
	$runlog
)

#region Variables
	# Checking if Runlog is empty
	If(!($Runlog)){
		$Runlog = "\$(Get-Date -uformat %Y%m%d-%H%M)-PSTImporter.log"
	}Else{
		$Runlog += "\$(Get-Date -uformat %Y%m%d-%H%M)-PSTImporter.log"
	}

	#region Synchronized Collections
	$Global:SyncHash = [hashtable]::Synchronized(@{})
	$Global:VariableProxy = [hashtable]::Synchronized(@{})
	
	# Setting variables
	$VariableProxy.ScriptDir = $PSScriptRoot
	$VariableProxy.FormsDir = -join($PSScriptRoot,"\Resources\Forms\")
	$VariableProxy.IconsDir = -join($PSScriptRoot,"\Resources\Icons\")
	$VariableProxy.Module = -join($PSScriptRoot,"\Resources\Module\Invoke-M365PSTImport.psm1")
	$VariableProxy.Runlog = -join($PSScriptRoot,$runlog)
	$VariableProxy.ResourceGroup = "pstuploads"
	$VariableProxy.MappingFile = -join($PSScriptRoot,"\Mapping.csv")
	$VariableProxy.SettingsFile = -join($PSScriptRoot,"\Settings.xml")
	$VariableProxy.TranscriptsDir = -join($PSScriptRoot,"\Transcripts\")
	#endregion Synchronized Collections
#endregion Variables

#region Prerequisites
	# Importing Module
	If(Test-Path $VariableProxy.Module){
		Import-Module $VariableProxy.Module -DisableNameChecking -Force
	}Else{
		Write-Warning -Message "Unable to import script module. Now Exiting on the right!"
		exit
	}

	## Clearing Screen
	cls

	## Output
	Invoke-Logging -LogLevel Title -Message "Microsoft 365 PST Uploader & Importer (Graphical Edition) starting!" -Runlog $VariableProxy.Runlog
	Write-Output ""
	Invoke-Logging -LogLevel INFO -Message "RUnlog will be saved in $($VariableProxy.Runlog)"

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

	# Transcripts Dir
	If(!(Test-path $VariableProxy.TranscriptsDir)){
		New-Item -ItemType Directory -Path $VariableProxy.TranscriptsDir
	}
#endregion

#region Configuring GUI Runspace
	
	# Runspace creation
	$syncHash.host = $Host
	$GUI_Runspace =[runspacefactory]::CreateRunspace()
	$GUI_Runspace.ApartmentState = "STA"
	$GUI_Runspace.ThreadOptions = "ReuseThread"
	$GUI_Runspace.Open()

	# Passing variables
	$GUI_Runspace.SessionStateProxy.SetVariable("syncHash",$syncHash)
	$GUI_Runspace.SessionStateProxy.SetVariable("VariableProxy",$VariableProxy)
	
#endregion Configuring GUI Runspace

#region Main
	# Scriptblock to execute in a runspace
	$psCmd = [PowerShell]::Create().AddScript({
	
		# Importing Module (Force to overwrite functions which may have already been imported, disabling name checking)
		Import-Module $VariableProxy.module -Force -DisableNameChecking

		# Load Required Assemblies
		Add-Type –assemblyName PresentationFramework # Required to show the GUI
		Add-Type -AssemblyName System.Windows.Forms # Required to use the folder browser dialog
		
		# Loading XAML code	
		[xml]$xaml = Get-Content -Path "$($VariableProxy.FormsDir)\Main.xaml"
		
		# Loading in to XML Node reader
		$reader = (New-Object System.Xml.XmlNodeReader $xaml)
		
		# Loading XML Node reader in to $SyncHash window property to launch later
		$SyncHash.Window = [Windows.Markup.XamlReader]::Load($reader)

		#region Connecting Controls
			# Main
			$synchash.BTN_Close = $SyncHash.Window.FindName("BTN_Close")
			$synchash.IMG_Close = $SyncHash.Window.FindName("IMG_Close")	
			$synchash.TitleBar = $SyncHash.Window.FindName("TitleBar")
			$synchash.OutputWindow = $SyncHash.Window.FindName("OutputWindow")

			# Configure Blob
			$synchash.TXT_StorageAccountName = $SyncHash.Window.FindName("TXT_StorageAccountName")
			$synchash.TXT_StorageAccountContainer = $SyncHash.Window.FindName("txt_StorageAccountContainer")
			$synchash.DatePicker_TokenExpireTime = $SyncHash.Window.FindName("DatePicker_TokenExpireTime")
			$synchash.DD_StorageLocation = $SyncHash.Window.FindName("DD_StorageLocation")
			$synchash.BTN_Create = $SyncHash.Window.FindName("BTN_Create")

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
		#endregion Connecting Controls

		# Configuring Images
		$SyncHash.IMG_Close.Source = "$($VariableProxy.IconsDir)\appbar.Close.png"
		$SyncHash.TitleBar.Text = "PST Upload and Import tool"

		# Adding items to drop down (Storage Location)
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
		
		# Reporting script loaded
		Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Tool loaded" -AppendContent  -logfile $VariableProxy.Runlog

		#region Event handlers
			# Enable window to move when dragged
			$Synchash.Window.FindName('Grid').Add_MouseLeftButtonDown({
				$Synchash.Window.DragMove()
			})

			# Close button
			$SyncHash.BTN_Close.Add_Click({
				$Synchash.Window.Close()
			})

			#region Create blob
			$Synchash.BTN_Create.Add_Click({
				# Capturing Values
				$VariableProxy.TokenExpiry = Get-Date($Synchash.DatePicker_TokenExpireTime.Text) -Format 'yyyy-MM-dd'
				$VariableProxy.StorageLocation = ($Synchash.DD_StorageLocation.Items.CurrentItem).ToLower()
				$VariableProxy.StorageAccountName = ($synchash.TXT_StorageAccountName.Text).ToLower()
				$VariableProxy.StorageAccountContainer = ($synchash.TXT_StorageAccountContainer.Text).ToLower()

				# Reporting
				Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Initiating creation of storage space`nStorage Region: $($VariableProxy.StorageLocation)`nToken Expires: $($VariableProxy.TokenExpiry)`nStorage URI: https://$($VariableProxy.StorageAccountName).blob.core.windows.net/$($VariableProxy.StorageAccountContainer)" -AppendContent -logfile $VariableProxy.Runlog			

				# Creating runspace to execute creation of storage blob
				$Blob_Runspace =[runspacefactory]::CreateRunspace()
                $Blob_Runspace.ApartmentState = "STA"
                $Blob_Runspace.ThreadOptions = "ReuseThread"
                $Blob_Runspace.Open()
                $Blob_Runspace.SessionStateProxy.SetVariable("synchash",$synchash)
                $Blob_Runspace.SessionStateProxy.SetVariable("VariableProxy",$VariableProxy)

				# Scriptblock
                $powershell = [PowerShell]::Create().AddScript({
					# Transcript for debugging
					Start-Transcript -Path "$($VariableProxy.TranscriptsDir)\$(Get-Date -uformat %Y%m%d-%H%M)-AzureBlobCreation.log"

					# Importing Module (Force to overwrite functions which may have already been imported, disabling name checking)
					Import-Module $VariableProxy.module -Force -DisableNameChecking

					Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Configuring Session" -AppendContent -logfile $VariableProxy.Runlog

					# Importing Azure RM module
					Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Importing Azure RM module" -AppendContent -logfile $VariableProxy.Runlog
					Import-Module AzureRM

					# Connecting to Azure
					Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Connecting to Azure Resource Manager" -AppendContent -logfile $VariableProxy.Runlog
					$Connection = connect-azureRmAccount

					# Header
					Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Starting Azure Storage Blob configuration Process" -AppendContent -logfile $VariableProxy.Runlog

					# Getting subscription name
					Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Retrieving Subscription Name" -AppendContent -logfile $VariableProxy.Runlog
					$SubscriptionName = $connection.Context.Subscription.Name

					# Creating a new resourceGroup
					Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Creating resource group: PSTUploads" -AppendContent -logfile $VariableProxy.Runlog
					New-AzureRmResourceGroup -Name $VariableProxy.resourceGroup -Location $VariableProxy.StorageLocation  | Out-Null

					# Creating Storage account
					Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Creating new storage account" -AppendContent -logfile $VariableProxy.Runlog
					New-AzureRMStorageAccount -ResourceGroupName $VariableProxy.resourceGroup -Name $VariableProxy.StorageAccountName -Location $VariableProxy.StorageLocation -SkuName "Standard_LRS" | Out-Null
	
					# Sleeping for 3 seconds to allow things to catch up
					Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Wait while Azure configures the storage account" -AppendContent -logfile $VariableProxy.Runlog
					Start-Sleep -Seconds 3

					# Setting subscription
					Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Setting storage account context" -AppendContent -logfile $VariableProxy.Runlog
					Set-AzureRmCurrentStorageAccount -ResourceGroupName $VariableProxy.ResourceGroup -Name $VariableProxy.StorageAccountName

					# Creating storage containter
					Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Creating new storage container" -AppendContent -logfile $VariableProxy.Runlog
					New-AzureStorageContainer -Name $VariableProxy.StorageAccountContainer -Permission Off | Out-null

					# Creating Token For importing PST
					Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Creating Token for storage account container" -AppendContent -logfile $VariableProxy.Runlog
					$token = New-AzureStorageContainerSASToken -Name $VariableProxy.StorageAccountContainer -Permission rwl -ExpiryTime $VariableProxy.TokenExpiry

					# Getting storage key for AZCopy
					Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Retrieving Storage Account Key" -AppendContent -logfile $VariableProxy.Runlog
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
					Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Azure storage blob creation completed!" -AppendContent -logfile $VariableProxy.Runlog
				
					Stop-Transcript
				})

				# Invoking Scriptblock
				$powershell.Runspace = $Blob_Runspace
                $data = $powershell.BeginInvoke()

			})
			#endregion

			#region Upload PST Tab
				# Folder browser for PST path
				$Synchash.BTN_Browse.Add_Click({
					# Location dialog for selecting folder
					$FolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
					$null = $FolderDialog.ShowDialog()

					# Setting the outputpath
					$SyncHash.TXT_PSTPath.text = $FolderDialog.SelectedPath

					# Reporting
					Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Path containing PST files: $($SyncHash.TXT_PSTPath.text)" -AppendContent -logfile $VariableProxy.Runlog
				})
			
				# Save settings
				$Synchash.BTN_Save.Add_Click({
				
					# Popup warning
					$MsgParams = @{
						TitleBackground = "Red"
						TitleTextForeground = "Yellow"
						TitleFontWeight = "UltraBold"
						TitleFontSize = 28
						ContentBackground = 'Red'
						ContentFontSize = 18
						ContentTextForeground = 'White'
						ButtonTextForeground = 'White'
					}

					Invoke-WPFMessageBox @MsgParams -Title "Attention! Sensitive data warning!" -Content "Saving these settings potentially exposes sensitive data.`n It is your responsibility to ensure this file is secured appropriately!" -ButtonType 'OK-Cancel' -Window $synchash.Window
				
					If($VariableProxy.Acceptance -like "OK"){
						# Exporting Settings
						$VariableProxy | Export-Clixml $VariableProxy.SettingsFile -Force

						# Reporting
						Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Settings saved" -AppendContent -logfile $VariableProxy.Runlog
					}Else{
						# Reporting
						Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Cancelled saving of settings" -AppendContent -logfile $VariableProxy.Runlog
					}
				})

				# Import settings
				$Synchash.BTN_Import.Add_Click({
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
					Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Import Completed" -AppendContent -logfile $VariableProxy.Runlog
				})

				# Upload PST Files
				$Synchash.BTN_Upload.Add_Click({
					
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
					$UploadPST_Runspace =[runspacefactory]::CreateRunspace()
					$UploadPST_Runspace.ApartmentState = "STA"
					$UploadPST_Runspace.ThreadOptions = "ReuseThread"
					$UploadPST_Runspace.Open()
					$UploadPST_Runspace.SessionStateProxy.SetVariable("synchash",$synchash)
					$UploadPST_Runspace.SessionStateProxy.SetVariable("VariableProxy",$VariableProxy)

					# Scriptblock
					$powershell = [PowerShell]::Create().AddScript({
						# Transcript for debugging
						Start-Transcript -Path "$($VariableProxy.TranscriptsDir)\$(Get-Date -uformat %Y%m%d-%H%M)-PSTUpload.log"

						# Importing Module (Force to overwrite functions which may have already been imported, disabling name checking)
						Import-Module $VariableProxy.module -Force -DisableNameChecking

						# Variables
						$Programfiles = [Environment]::GetEnvironmentVariable("ProgramFiles")
				
						# Reporting
						Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Starting Upload Process" -AppendContent -logfile $VariableProxy.Runlog

						# AZ Copy options
						Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Setting AZCopy options (Recurse, Verbose output logging, write only tokens), PST files only" -AppendContent -logfile $VariableProxy.Runlog
						Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: AZ Copy log is located in %LocalAppData%\Microsoft\Azure\AzCopy" -AppendContent -logfile $VariableProxy.Runlog
						$AZCOptions = " /S /V /Y /Pattern:*.pst"

						# Go through all subfolders 
						$PSTFiles = Get-ChildItem -Path $VariableProxy.PSTPath -Filter "*.PST" -Recurse

						Foreach($PST in $PSTFiles){
							$Timestamp = (get-date -Format HH:mm:ss)
							Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Processing $($PST.Name)" -AppendContent -logfile $VariableProxy.Runlog

							$CSV = [PSCustomObject]
							$CSV = "" | Select email,PSTName, Status

							$CSV.Email = $VariableProxy.EmailAddress
							$CSV.PSTName = $PST.Name
							$CSV.Status = "Not Started"

							$CSV | Export-Csv -Path $VariableProxy.MappingFile -NoTypeInformation -Append
						}

						# Starting AZCopy Process
						Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Starting PST upload (AZCopy)" -AppendContent -logfile $VariableProxy.Runlog
						Start-Process -FilePath "$($VariableProxy.ProgramFiles)\Microsoft SDKs\Azure\AZCopy\AZCopy.exe" -ArgumentList "/Source:`"$($VariableProxy.PSTPath)`" /Dest:$($VariableProxy.AccountURI) /DestKey:$($VariableProxy.StorageKey) $AZCOptions" -Wait

						# Report completed
						Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: PST upload completed (AZCopy)" -AppendContent -logfile $VariableProxy.Runlog

						Stop-Transcript
					})

					# Invoking Scriptblock
					$powershell.Runspace = $UploadPST_Runspace
					$data = $powershell.BeginInvoke()
				})
			#endregion

			#region Import PST
				# Start PST Import
				$Synchash.BTN_Import_PST.Add_Click({

					# Runspace building (If we don't run the code in a seperate runspace the GUI will lock up. Nobody like a GUI that locks up!)
					$ImportPST_Runspace =[runspacefactory]::CreateRunspace()
					$ImportPST_Runspace.ApartmentState = "STA"
					$ImportPST_Runspace.ThreadOptions = "ReuseThread"
					$ImportPST_Runspace.Open()
					$ImportPST_Runspace.SessionStateProxy.SetVariable("synchash",$synchash)
					$ImportPST_Runspace.SessionStateProxy.SetVariable("VariableProxy",$VariableProxy)

					# Scriptblock
					$powershell = [PowerShell]::Create().AddScript({

						# Transcript for debugging
					    Start-Transcript -Path "$($VariableProxy.TranscriptsDir)\$(Get-Date -uformat %Y%m%d-%H%M)-PSTImport.log"

						# Importing Module (Force to overwrite functions which may have already been imported, disabling name checking)
						Import-Module $VariableProxy.module -Force -DisableNameChecking

						Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Starting PST import process" -AppendContent -logfile $VariableProxy.Runlog

						Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Importing Mapping file" -AppendContent -logfile $VariableProxy.Runlog
						$CSV = Import-Csv $VariableProxy.MappingFile

						Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Connecting to Exchange Online" -AppendContent -logfile $VariableProxy.Runlog
						$UserCredential = Get-Credential
						$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
						Import-PSSession $Session -DisableNameChecking

						# Iterating through the CSV File
						Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Iterating through Mapping file and creating PST Import Requests" -AppendContent -logfile $VariableProxy.Runlog
						Foreach($i in $CSV){
							# If this has been started before, skip
							If($i.status -eq "Started"){
								Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[ALERT]: PST import previously started: $($I.PSTName) for $($i.email)" -AppendContent -logfile $VariableProxy.Runlog
							}Else{
								# Processing users
								Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Processing: $($i.email)" -AppendContent -logfile $VariableProxy.Runlog

								# Import request
								Try{
									New-MailboxImportRequest -TargetRootFolder "Imported PST" -Mailbox $($i.email) -AzureBlobStorageAccountUri $VariableProxy.AccountUri/$i.PSTname -AzureSharedAccessSignatureToken $VariableProxy.SASToken
									($csv |where {$_.PSTName -eq $i.PSTName}).Status = "Started"
									Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Created PST request for $($i.email), $($i.PSTName)" -AppendContent -logfile $VariableProxy.Runlog
								}Catch{
									Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[ALERT]: Unable to create PST import: $($i.email), $($i.PSTName)" -AppendContent -logfile $VariableProxy.Runlog
								}
							}
						}

						# Exporting updated Mapping File
						Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Exporting updated Mapping file" -AppendContent -logfile $VariableProxy.Runlog
						$CSV | Export-Csv $VariableProxy.MappingFile -NoTypeInformation -Force

						# Removing Exchange online session
						Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: Closing connection to Exchange Online" -AppendContent -logfile $VariableProxy.Runlog
						Remove-PSSession $Session

						# Reporting Done
						Update-Control -Synchash $Synchash -Control "OutputWindow" -value "[INFO]: PST Import process completed" -AppendContent -logfile $VariableProxy.Runlog

						Stop-Transcript
					})

					# Invoking Scriptblock
					$powershell.Runspace = $ImportPST_Runspace
					$data = $powershell.BeginInvoke()

				})
			#endregion
		#endregion

		# Show GUI
		$SyncHash.Window.ShowDialog() | Out-Null
		$VariableProxy.Error = $Error
	})

	# Invoking GUI
	$psCmd.Runspace = $GUI_Runspace
	$data = $psCmd.BeginInvoke()
#endregion Main