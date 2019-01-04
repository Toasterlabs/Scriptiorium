function Invoke-ColorOutput{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$False,Position=1,ValueFromPipeline=$True,ValueFromPipelinebyPropertyName=$True)]
		[alias('message')]
		[alias('msg')]
		[Object]$Object,
        [Parameter(Mandatory=$False,Position=2,ValueFromPipeline=$True,ValueFromPipelinebyPropertyName=$True)]
		[ValidateSet('Black', 'DarkBlue', 'DarkGreen', 'DarkCyan', 'DarkRed', 'DarkMagenta', 'DarkYellow', 'Gray', 'DarkGray', 'Blue', 'Green', 'Cyan', 'Red', 'Magenta', 'Yellow', 'White')]
		[alias('fore')]
		[ConsoleColor] $ForegroundColor,
        [Parameter(Mandatory=$False,Position=3,ValueFromPipeline=$True,ValueFromPipelinebyPropertyName=$True)]
		[ValidateSet('Black', 'DarkBlue', 'DarkGreen', 'DarkCyan', 'DarkRed', 'DarkMagenta', 'DarkYellow', 'Gray', 'DarkGray', 'Blue', 'Green', 'Cyan', 'Red', 'Magenta', 'Yellow', 'White')]
		[alias('back')]
		[alias('BGR')]
		[ConsoleColor] $BackgroundColor,
        [Switch]$NoNewline
    )    

    # Save previous colors
    $previousForegroundColor = $host.UI.RawUI.ForegroundColor
    $previousBackgroundColor = $host.UI.RawUI.BackgroundColor

    # Set BackgroundColor if available
    if($BackgroundColor -ne $null)
    { 
       $host.UI.RawUI.BackgroundColor = $BackgroundColor
    }

    # Set $ForegroundColor if available
    if($ForegroundColor -ne $null)
    {
        $host.UI.RawUI.ForegroundColor = $ForegroundColor
    }

    # Always write (if we want just a NewLine)
    if($Object -eq $null)
    {
        $Object = ""
    }

    if($NoNewline)
    {
        [Console]::Write($Object)
    }
    else
    {
        Write-Output $Object
    }

    # Restore previous colors
    $host.UI.RawUI.ForegroundColor = $previousForegroundColor
    $host.UI.RawUI.BackgroundColor = $previousBackgroundColor
}

Function invoke-Logging{
    Param(
        [Parameter(Mandatory=$True)]
        [STRING]$Message,
		[Parameter(Mandatory=$True)]
		[Validateset('TEXT','TITLE','STATUS','AUDIT','INFO','SUCCESS','ALERT','WARNING','ERROR','CRITICAL','VERBOSE','DEBUG')]
		[STRING]$LogLevel,
        [Parameter(Mandatory=$False)]
        [System.IO.FileInfo]$RunLog
	)

		Switch($LogLevel){
			"TEXT"     {
				Invoke-ColorOutput $Message -ForegroundColor White
				Try{
					$Message | Out-File $RunLog -Append -ErrorAction SilentlyContinue
				}Catch{
					<#No Error Handling... This is horrible!#>
				}
			}
			"TITLE"    {
				Invoke-ColorOutput $Message -ForegroundColor Magenta
				Try{
					$Message | Out-File $RunLog -Append -ErrorAction SilentlyContinue
				}Catch{
					<#No Error Handling... This is horrible!#>
				}
			}
			"STATUS"   {
				$Message = (get-date -Format HH:mm:ss) + " - [STATUS]: " + $Message
				Invoke-ColorOutput $Message -ForegroundColor Magenta
				Try{
					$Message | Out-File $RunLog -Append -ErrorAction SilentlyContinue
				}Catch{
					<#No Error Handling... This is horrible!#>
				}
			}
			"AUDIT"    {
				$Message = (get-date -Format HH:mm:ss) + " - [AUDIT]: " + $Message 
				Invoke-ColorOutput $Message -ForegroundColor DarkGrey
				Try{
					$Message | Out-File $RunLog -Append -ErrorAction SilentlyContinue
				}Catch
				{<#No Error Handling... This is horrible!#>
				}
			}
			"INFO"     {
				$Message = (get-date -Format HH:mm:ss) + " - [INFO]: " + $Message 
				Invoke-ColorOutput $Message -ForegroundColor Cyan
				Try{
					$Message | Out-File $RunLog -Append -ErrorAction SilentlyContinue
				}Catch{
					<#No Error Handling... This is horrible!#>
				}
			}
			"SUCCESS"     {
				$Message = (get-date -Format HH:mm:ss) + " - [SUCCESS]: " + $Message 
				Invoke-ColorOutput $Message -ForegroundColor Green
				Try{
					$Message | Out-File $RunLog -Append -ErrorAction SilentlyContinue
				}Catch{
					<#No Error Handling... This is horrible!#>
				}
			}
			"ALERT"    {
				$Message = (get-date -Format HH:mm:ss) + " - [ALERT]: " + $Message
				Write-Host $Message -ForegroundColor Yellow
				Try{
					$Message | Out-File $RunLog -Append -ErrorAction SilentlyContinue
				}Catch{
					<#No Error Handling... This is horrible!#>
				}
			}
			"NEWLINE"  {Write-Host "";try{Write-Output "`n" | Out-File $RunLog -Append -ErrorAction SilentlyContinue}Catch{}}
			"WARNING"  {
				Write-Warning -Message $Message
				Try{
					$Message | Out-File $RunLog -Append -ErrorAction SilentlyContinue
				}Catch{
					<#No Error Handling... This is horrible!#>
				}
			}
			"ERROR"    {
				Write-Error -Message $Message
				Try{
					$Message | Out-File $RunLog -Append -ErrorAction SilentlyContinue
				}Catch{
					<#No Error Handling... This is horrible!#>
				}
			}
			"CRITICAL" {
				$Message = (get-date -Format HH:mm:ss) + " - [CRITICAL]: " + $Message
				Invoke-colorOutput $Message -ForegroundColor White -BGR Red
				Try{
					$Message | Out-File $RunLog -Append -ErrorAction SilentlyContinue
				}Catch{
					<#No Error Handling... This is horrible!#>
				}
			}
			"VERBOSE" {
				$Message = (get-date -Format HH:mm:ss) + " - [VERBOSE]: " + $Message
				Write-Verbose -Message $Message
				Try{
					$Message | Out-File $RunLog -Append -ErrorAction SilentlyContinue
				}Catch{
					<#No Error Handling... This is horrible!#>
				}
			}
			"DEBUG" {
				$Message = (get-date -Format HH:mm:ss) + " - [CRITICAL]: " + $Message
				Invoke-colorOutput $Message -ForegroundColor White -BGR Red
				Try{
					$Message | Out-File $RunLog -Append -ErrorAction SilentlyContinue
				}Catch{
					<#No Error Handling... This is horrible!#>
				}
			}
		}
}

Function Get-ElapsedTime{
	<#
		.SYNOPSIS
			Calculates a time interval between two DateTime objects.

		.DESCRIPTION
			Calculates a time interval between two DateTime objects.
			Use $ScriptStartTime = Get-Date & $ScriptStopTime = Get-Date variables to measure the time it took to run the script

		.PARAMETER Start
			Specifies the start time.

		.PARAMETER End
			Specifies the end time.

		.EXAMPLE
			PS> GetElapsedTime -Start "1/1/2011 12:00:00 AM" -End "1/2/2011 2:00:00 PM"

		.EXAMPLE
			PS> GetElapsedTime -Start $scriptStartTime -End $ScriptStopTime

		.INPUTS
			System.String

		.OUTPUTS
			System.Management.Automation.PSObject

		.NOTES

	#>

	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $True, Position = 0)]
		[ValidateNotNullOrEmpty()]
		[DateTime]$Start,
		
		[Parameter(Mandatory = $True, Position = 1)]
		[ValidateNotNullOrEmpty()]
		[DateTime]$End
	)
	
	$TotalSeconds = ($End).Subtract($Start).TotalSeconds
	$objElapsedTime = New-Object PSObject
	
	# less than 1 minute
	If ($TotalSeconds -lt 60)
	{
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Days -Value 0
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Hours -Value 0
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Minutes -Value 0
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Seconds -Value $($TotalSeconds)
	}

	# more than 1 minute, less than 1 hour
	If (($TotalSeconds -ge 60) -and ($TotalSeconds -lt 3600))
	{
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Days -Value 0
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Hours -Value 0
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Minutes -Value $([Math]::Truncate($TotalSeconds / 60))
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Seconds -Value $([Math]::Truncate($TotalSeconds % 60))
	}

	# more than 1 hour, less than 1 day
	If (($TotalSeconds -ge 3600) -and ($TotalSeconds -lt 86400))
	{
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Days -Value 0
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Hours -Value $([Math]::Truncate($TotalSeconds / 3600))
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Minutes -Value $([Math]::Truncate(($TotalSeconds % 3600) / 60))
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Seconds -Value $([Math]::Truncate($TotalSeconds % 60))
	}

	# more than 1 day, less than 1 year
	If (($TotalSeconds -ge 86400) -and ($TotalSeconds -lt 31536000))
	{
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Days -Value $([Math]::Truncate($TotalSeconds / 86400))
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Hours -Value $([Math]::Truncate(($TotalSeconds % 86400) / 3600))
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Minutes -Value $([Math]::Truncate((($TotalSeconds - 86400) % 3600) / 60))
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Seconds -Value $([Math]::Truncate($TotalSeconds % 60))
	}
	
	Return $objElapsedTime
}

Function Get-OSVersion{
	$OSVersion = [PSCustomObject]
	$OSVersion = '' | Select ProductName, 64Bit
	# Retrieving
	Write-Verbose "Retrieving OS Version"
	$OSVersion.ProductName = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ProductName
	$OsVersion."64Bit" = [System.Environment]::Is64BitOperatingSystem

	# Return
	return $OSVersion
}

Function Invoke-ConnectivityCheck{
	[CmdletBinding()]
	Param(
		[SWITCH]$Internet
	)

	# Object
	$Result = [PSCustomObject]
	$Result = "" | Select Internet

	# Region Internet
	If($Internet){
		Write-Verbose -Message "Getting Internet connectivity state"
		$HasInternetAccess = ([Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}')).IsConnectedToInternet)
	}
	
	If($HasInternetAccess){
		Write-Verbose -Message "System is connected to the internet"
		$Result.Internet = $true
	}Else{
		Write-Verbose -Message "System is not connected to the internet"
		$Result.Internet = $false
	}
	#endregion

	# Returning object
	return $Result
}

Function Invoke-IsAdmin{
	[CmdletBinding()]
	Param()

	# Retrieving
	Write-Verbose -Message "Retrieving Windows Security Principal"
	$ThisPrincipal = new-object System.Security.principal.windowsprincipal( [System.Security.Principal.WindowsIdentity]::GetCurrent())
	Write-Verbose -Message "Checking if the user in in the Administrator Role"
	$IsAdmin = $ThisPrincipal.IsInRole("Administrators")

	If($IsAdmin){
		Write-Verbose -Message "User is in administrator role"
		$Result = $True
	}Else{
		Write-Verbose -Message "User is in not administrator role"
		$Result = $False
	}

	Return $Result
}

Function invoke-ModuleVerify{
	[CmdletBinding()]
	Param(
		[Parameter(HelpMessage = "Verify the MSOnline module is installed")]
		[SWITCH]$MSOnline,
		[Parameter(HelpMessage = "Verify the Azure AD module is installed", ParameterSetName = 'AzureAD')]
		[SWITCH]$AzureAD,
		[Parameter(HelpMessage = "Verify the Azure AD Preview module is installed", ParameterSetName = 'AzureADPreview')]
		[SWITCH]$AzureADPreview,
		[Parameter(HelpMessage = "Verify the SharePoint Online module is installed")]
		[SWITCH]$SharePointOnline,
		[Parameter(HelpMessage = "Verify the Skype for Business Online module is installed")]
		[SWITCH]$SkypeForBusinessOnline,
		[Parameter(HelpMessage = "Verify the Microsoft Teams module is installed")]
		[SWITCH]$MicrosoftTeams,
		[SWITCH]$Install
	)

	# New object
	$Return = [PSCustomObject]
	$Return = "" | Select MSOnline, AzureAD, AzureADPreview, SharePointOnline, SkypeForBusinessOnline, MicrosoftTeams

	#region MSOnline module
    If($MSOnline){
		# verify that the MSOnline module is installed and import into current powershell session
		If (!(get-childitem -path 'C:\Program Files\WindowsPowerShell\Modules\MSOnline' -Recurse -Filter "MSOnline.psd1")){
			If($install){
				try{
					Write-Verbose -Message "Installing Microsoft Online Services Module"
					Install-Module MSOnline 
				}Catch{
					Write-Warning -Message "Unable to install the Microsoft Online Services module. Please download and install manually..."
					Write-Error -Message $_
					$Return.MSOnline = $false
				}
			}Else{
				write-warning -Message ("Please download and install the Microsoft Online Services Module. Exiting...")
				$Return.MSOnline = $false
			}
		}

		$getModuleResults = Get-Module

		If (!$getModuleResults) {
			write-verbose -Message "Importing MSOnline module"
			Try{
				Import-Module MSOnline -ErrorAction SilentlyContinue
				$Return.MSOnline = $true
			}Catch{
				Write-Warning -Message "Unable to import Microsoft Online Services Module"
				$Return.MSOnline = $false
			}

			Import-Module MSOnline -ErrorAction SilentlyContinue}
		Else {$getModuleResults | ForEach-Object {If (!($_.Name -eq "MSOnline")){
			write-verbose -Message "Importing MSOnline module"
			Try{
				Import-Module MSOnline -ErrorAction SilentlyContinue
				$Return.MSOnline = $true
			}Catch{
				Write-Warning -Message "Unable to import Microsoft Online Services Module"
				$Return.MSOnline = $false
			} # End Catch
		}# End If
	}# End Foreach
	}# End Else
	}Else{$Return.MSOnline = "NotChecked"}
	#endregion

	#region Azure AD module
    If($AzureAD){
		# verify that the MSOnline module is installed and import into current powershell session
		If (!(Get-ChildItem -Path 'C:\Program Files\WindowsPowerShell\Modules\AzureAD\' -Recurse -Filter "AzureAD.psd1" -ErrorAction SilentlyContinue)){
			If($install){
				try{
					Write-Verbose -Message "Installing Microsoft Azure AD Module"
					Install-Module AzureAD 
				}Catch{
					Write-Warning -Message "Unable to install the Microsoft Azure AD module. Please download and install manually..."
					Write-Error -Message $_
					$Return.AzureAD = $false
				}
			}Else{
				write-warning -Message ("Please download and install the Microsoft Azure AD Module. Exiting...")
				$Return.AzureAD = $false
			}
		}

		$getModuleResults = Get-Module

		If (!$getModuleResults) {write-verbose -Message "Importing Microsoft Azure AD module"; Import-Module AzureAD -ErrorAction SilentlyContinue;}
		Else {$getModuleResults | ForEach-Object {If (!($_.Name -eq "AzureAD")){
			write-verbose -Message "Importing MSOnline module"
			Try{
				Import-Module AzureAD -ErrorAction SilentlyContinue
				$Return.AzureAD = $true
			}Catch{
				Write-Warning -Message "Unable to import Microsoft Online Services Module"
				$Return.AzureAD = $false
			} # End Catch
		} # end if
		}#end foreach
		}# End Else
	}Else{$Return.AzureAD = "NotChecked"}
	#endregion

	#region Azure AD Preview Module
	If($AzureADPreview){
		# verify that the MSOnline module is installed and import into current powershell session
		If (!(Get-ChildItem -Path 'C:\Program Files\WindowsPowerShell\Modules\AzureADPreview\' -Recurse -Filter "AzureADPreview.psd1")){
			If($install){
				try{
					Write-Verbose -Message "Installing Microsoft Azure AD Module"
					Install-Module AzureADPreview
				}Catch{
					Write-Warning -Message "Unable to install the Microsoft Azure AD Preview module. Please download and install manually..."
					Write-Error -Message $_
					$Return.AzureADPreview = $false
				}
			}Else{
				write-warning -Message ("Please download and install the Microsoft Azure AD Preview Module.")
				$Return.AzureADPreview = $false
			}
		}

		$getModuleResults = Get-Module

		If (!$getModuleResults) {write-verbose -Message "Importing Microsoft Azure AD Preview module"; Import-Module AzureADPreview -ErrorAction SilentlyContinue;}
		Else {$getModuleResults | ForEach-Object {If (!($_.Name -eq "AzureADPreview")){
			write-verbose -Message "Importing MSOnline module"
			Try{
				Import-Module AzureADPreview -ErrorAction SilentlyContinue
				$Return.AzureADPreview = $true
			}Catch{
				Write-Warning -Message "Unable to import Microsoft Azure AD Preview Module"
				$Return.AzureADPreview = $false
			} # End Catch
		} # end if
		}#end foreach
		}# End Else
	}Else{$Return.AzureADPreview = "NotChecked"}
	#endregion

	#region SharePoint Online Module
	If($SharePointOnline){
		# verify that the MSOnline module is installed and import into current powershell session
		If (!(Get-ChildItem -Path 'C:\Program Files\WindowsPowerShell\Modules\Microsoft.Online.SharePoint.PowerShell\' -Recurse -Filter "Microsoft.Online.SharePoint.PowerShell.psd1")){
			If($install){
				try{
					Write-Verbose -Message "Installing Microsoft SharePoint Online Module"
					Install-Module Microsoft.Online.SharePoint.PowerShell
				}Catch{
					Write-Warning -Message "Unable to install the Microsoft SharePoint Online module. Please download and install manually..."
					Write-Error -Message $_
					$Return.SharePointOnline = $false
				}
			}Else{
				write-warning -Message ("Please download and install the Microsoft SharePoint Online Preview Module. Exiting...")
				$Return.SharePointOnline = $false
			}
		}

		$getModuleResults = Get-Module

		If (!$getModuleResults) {write-verbose -Message "Importing Microsoft SharePoint Online Preview module"; Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction SilentlyContinue;}
		Else {$getModuleResults | ForEach-Object {If (!($_.Name -eq "Microsoft.Online.SharePoint.PowerShell")){
			write-verbose -Message "Importing MSOnline module"
			Try{
				Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction SilentlyContinue
				$Return.SharePointOnline = $true
			}Catch{
				Write-Warning -Message "Unable to import Microsoft SharePoint Online Preview Module"
				$Return.SharePointOnline = $false
			} # End Catch
		} # end if
		}#end foreach
		}# End Else
	}Else{$Return.SharePointOnline = "NotChecked"}
	#endregion

	#region Skype for Business Online Module
	If($SkypeForBusinessOnline){
		# verify that the MSOnline module is installed and import into current powershell session
		If (!(Get-ChildItem -Path 'C:\Program Files\WindowsPowerShell\Modules\SkypeOnlineConnector\' -Recurse -Filter "SkypeOnlineConnector.psd1")){
			If($install){
				rite-warning -Message ("Unable to install the Skype for Bussiness Module. Please download and install the Microsoft SharePoint Online Preview Module. Exiting...")
				$Return.SkypeForBusinessOnline = $false
			}Else{
				write-warning -Message ("Please download and install the Microsoft SharePoint Online Preview Module. Exiting...")
				$Return.SkypeForBusinessOnline = $false
			}
		}

		$getModuleResults = Get-Module

		If (!$getModuleResults) {write-verbose -Message "Importing Microsoft SharePoint Online Preview module"; Import-Module SkypeOnlineConnector -ErrorAction SilentlyContinue;}
		Else {$getModuleResults | ForEach-Object {If (!($_.Name -eq "SkypeOnlineConnector")){
			write-verbose -Message "Importing MSOnline module"
			Try{
				Import-Module SkypeOnlineConnector -ErrorAction SilentlyContinue
				$Return.SkypeForBusinessOnline = $true
			}Catch{
				Write-Warning -Message "Unable to import Microsoft SharePoint Online Preview Module"
				$Return.SkypeForBusinessOnline = $false
			} # End Catch
		} # end if
		}#end foreach
		}# End Else
	}Else{$Return.SkypeForBusinessOnline = "NotChecked"}
	#endregion

	#region MSOnline module
    If($MicrosoftTeams){
		# verify that the MSOnline module is installed and import into current powershell session
		If (!(get-childitem -path 'C:\Program Files\WindowsPowerShell\Modules\MicrosoftTeams' -Recurse -Filter "MicrosoftTeams.psd1")){
			If($install){
				try{
					Write-Verbose -Message "Installing Microsoft Teams Module"
					Install-Module MicrosoftTeams 
				}Catch{
					Write-Warning -Message "Unable to install the Microsoft Teams module. Please download and install manually..."
					Write-Error -Message $_
					$Return.MicrosoftTeams = $false
				}
			}Else{
				write-warning -Message ("Please download and install the Microsoft Teams Module. Exiting...")
				$Return.MicrosoftTeams = $false
			}
		}

		$getModuleResults = Get-Module

		If (!$getModuleResults) {
			write-verbose -Message "Importing Microsoft Teams module"
			Try{
				Import-Module MicrosoftTeams -ErrorAction SilentlyContinue
				$Return.MicrosoftTeams = $true
			}Catch{
				Write-Warning -Message "Unable to import Microsoft Teams Module"
				$Return.MicrosoftTeams = $false
			}

			Import-Module MicrosoftTeams -ErrorAction SilentlyContinue}
		Else {$getModuleResults | ForEach-Object {If (!($_.Name -eq "MicrosoftTeams")){
			write-verbose -Message "Importing Microsoft Teams module"
			Try{
				Import-Module MicrosoftTeams -ErrorAction SilentlyContinue
				$Return.MicrosoftTeams = $true
			}Catch{
				Write-Warning -Message "Unable to import Microsoft Teams Module"
				$Return.MicrosoftTeams = $false
			} # End Catch
		}# End If
	}# End Foreach
	}# End Else
	}Else{$Return.MicrosoftTeams = "NotChecked"}
	#endregion

	Return $Return
}

function write-window{
<#
.Synopsis
  Writes to GUI

.Description
  Used for logging. Will preface any message written through this function with a timestamp

.PARAMETER  Synchash
  Needed to write back to the Output text box

.PARAMETER MESSAGE
  content to add

.OUTPUTS
  Content to log file

#>
    Param(
        $message,
        $Synchash
    )

    $timestamp = (get-date -Format HH:mm:ss)
    $entry = $timestamp + " - " + "$message`n"
    $SyncHash.OutputWindow.AppendText("`n")
    $SyncHash.OutputWindow.AppendText($entry)

}

Function Update-Control {
        Param (
            $syncHash,
            $Control,
            $Property,
            $Value,
			$logFile,
            [switch]$AppendContent,
            [Switch]$Add
        )

        # This is kind of a hack, there may be a better way to do this
        If ($Property -eq "Close") {
            $syncHash.Window.Dispatcher.invoke([action]{$syncHash.Window.Close()},"Normal")
            Return
        }

        # This updates the control based on the parameters passed to the function
        $syncHash.$Control.Dispatcher.Invoke([action]{
            # This bit is only really meaningful for the TextBox control, which might be useful for logging progress steps
            If ($PSBoundParameters['AppendContent']) {
				
				# Adding time stamp
				$Timestamp = (get-date -Format HH:mm:ss)
				$value = "$Timestamp - $value"
				
				# Updating Control
				$syncHash.$Control.AppendText("`n")
                $syncHash.$Control.AppendText($Value)
				$syncHash.$Control.ScrollToEnd()
				
				# Logging to file
				$Value | Out-File $LogFile -Append -ErrorAction SilentlyContinue
            }
            If ($PSBoundParameters['Add']) {
                $current = $syncHash.$Control.$Property
                $Value = $current + $Value
                $syncHash.$Control.$Property = $Value
            } Else {
                $syncHash.$Control.$Property = $Value
            }
        }, "Normal")
    }

Function New-Popup {
	Param (
	[Parameter(Position=0,Mandatory=$True,HelpMessage="Enter a message for the popup")]
	[ValidateNotNullorEmpty()]
	[string]$Message,
	[Parameter(Position=1,Mandatory=$True,HelpMessage="Enter a title for the popup")]
	[ValidateNotNullorEmpty()]
	[string]$Title,
	[Parameter(Position=2,HelpMessage="How many seconds to display? Use 0 require a button click.")]
	[ValidateScript({$_ -ge 0})]
	[int]$Time=0,
	[Parameter(Position=3,HelpMessage="Enter a button group")]
	[ValidateNotNullorEmpty()]
	[ValidateSet("OK","OKCancel","AbortRetryIgnore","YesNo","YesNoCancel","RetryCancel","CancelTryAgainCOntinue")]
	[string]$Buttons="OK",
	[Parameter(Position=4,HelpMessage="Enter an icon set")]
	[ValidateNotNullorEmpty()]
	[ValidateSet("Stop","Question","Exclamation","Information" )]
	[string]$Icon="Information",
	[Parameter(HelpMessage="Select which button will be selected as the default")]
	[ValidateSet("Second","Third")]
	[string]$DefaultButton

	)

	#convert buttons to their integer equivalents
	Switch ($Buttons) {
		"OK"               {$ButtonValue = 0}
		"OKCancel"         {$ButtonValue = 1}
		"AbortRetryIgnore" {$ButtonValue = 2}
		"YesNo"            {$ButtonValue = 4}
		"YesNoCancel"      {$ButtonValue = 3}
		"RetryCancel"      {$ButtonValue = 5}
		"CancelTryAgainCOntinue" {$ButtonValue = 6}
	}

	#set an integer value for Icon type
	Switch ($Icon) {
		"Stop"        {$iconValue = 16}
		"Question"    {$iconValue = 32}
		"Exclamation" {$iconValue = 48}
		"Information" {$iconValue = 64}
	}

	Switch($DefaultButton){
		"Second" {$Selectedbtn = 256}
		"Third"  {$Selectedbtn = 512}

	}

	#create the COM Object
	Try {
		$wshell = New-Object -ComObject Wscript.Shell -ErrorAction Stop
		#Button and icon type values are added together to create an integer value
		$wshell.Popup($Message,$Time,$Title,$ButtonValue + $iconValue + $Selectedbtn)
	}
	Catch {
		#You should never really run into an exception in normal usage
		Write-Warning "Failed to create Wscript.Shell COM object"
		Write-Warning $_.exception.message
	}

}

Function Invoke-M365PSTImport {
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
	[parameter(ParameterSetName = 'ImportPST')]
	$CSVPath,
	[Parameter(Mandatory = $false, HelpMessage = 'User we are processing', ParameterSetName='UploadPST')]
	[Alias("Email")]
	[Alias("mail")]
	[Alias("UPN")]
	$EmailAddress,
	#endregion
	#region Import PST
	[Parameter(Mandatory = $false, HelpMessage = 'Import PST Files', ParameterSetName='ImportPST')]
	[SWITCH]$ImportPST,
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
		$CSV = "" | Select email,PSTName, Status

		$CSV.Email = $EmailAddress
		$CSV.PSTName = $PST.Name
		$CSV.Status = "Not Started"

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
If($ImportPST){

	# Header
	Write-Output ""
	Invoke-Logging -LogLevel INFO -Message "Starting PST Import Process"

	# Importing CSV
	Invoke-Logging -LogLevel INFO -Message "Importing CSV File"
	$CSV = Import-Csv $CSVPath

	# Connecting to Exchange Online
	Invoke-Logging -LogLevel INFO -Message "Connecting to Exchange Online"
	$UserCredential = Get-Credential
	$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
	Import-PSSession $Session -DisableNameChecking

	# Iterating through list of items in CSV and starting upload process
	Invoke-Logging -LogLevel INFO -Message "Creating PST Import requests"
	Foreach($i in $CSV){

		If($i.status -eq "Started"){
			Invoke-Logging -LogLevel ALERT -Message "PST import previously started: $($I.PSTName) for $($i.email)"
		}Else{
			# Processing
			Invoke-Logging -LogLevel INFO -Message "Processing: $($i.email)"

			# Import request
			Try{
				New-MailboxImportRequest -TargetRootFolder "Imported PST" -Mailbox $($i.email) -AzureBlobStorageAccountUri $AzureBlobStorageAccountUri/$i.PSTname -AzureSharedAccessSignatureToken $AzureBlobStorageAccountUri
				($csv |where {$_.PSTName -eq $i.PSTName}).Status = "Started"
			}Catch{
				Invoke-Logging -LogLevel ALERT -Message "Unable to start PST import: $($I.PSTName) for $($i.email)"
			}
		}

	}

	# Exporting CSV
	$CSV | Export-Csv $CSVPath -NoTypeInformation -Force
	
	# Exiting session with Exchange Online
	Remove-PSSession $Session


}
#endregion
}

Function Invoke-WPFMessageBox {

    <#
		Function by SMS Agent
		https://smsagent.blog/2017/08/24/a-customisable-wpf-messagebox-for-powershell/
	#>
    
    # Define Parameters
    [CmdletBinding()]
    Param
    (
        # The popup Content
        [Parameter(Mandatory=$True,Position=0)]
        [Object]$Content,

        # The window title
        [Parameter(Mandatory=$false,Position=1)]
        [string]$Title,

        # The buttons to add
        [Parameter(Mandatory=$false,Position=2)]
        [ValidateSet('OK','OK-Cancel','Abort-Retry-Ignore','Yes-No-Cancel','Yes-No','Retry-Cancel','Cancel-TryAgain-Continue','None')]
        [array]$ButtonType = 'OK',

        # The buttons to add
        [Parameter(Mandatory=$false,Position=3)]
        [array]$CustomButtons,

        # Content font size
        [Parameter(Mandatory=$false,Position=4)]
        [int]$ContentFontSize = 14,

        # Title font size
        [Parameter(Mandatory=$false,Position=5)]
        [int]$TitleFontSize = 14,

        # BorderThickness
        [Parameter(Mandatory=$false,Position=6)]
        [int]$BorderThickness = 0,

        # CornerRadius
        [Parameter(Mandatory=$false,Position=7)]
        [int]$CornerRadius = 8,

        # ShadowDepth
        [Parameter(Mandatory=$false,Position=8)]
        [int]$ShadowDepth = 3,

        # BlurRadius
        [Parameter(Mandatory=$false,Position=9)]
        [int]$BlurRadius = 20,

        # WindowHost
        [Parameter(Mandatory=$false,Position=10)]
        [object]$WindowHost,

        # Timeout in seconds,
        [Parameter(Mandatory=$false,Position=11)]
        [int]$Timeout,

        # Code for Window Loaded event,
        [Parameter(Mandatory=$false,Position=12)]
        [scriptblock]$OnLoaded,

        # Code for Window Closed event,
        [Parameter(Mandatory=$false,Position=13)]
        [scriptblock]$OnClosed

    )

    # Dynamically Populated parameters
    DynamicParam {
        
        # Add assemblies for use in PS Console 
        Add-Type -AssemblyName System.Drawing, PresentationCore
        
        # ContentBackground
        $ContentBackground = 'ContentBackground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $RuntimeParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.ContentBackground = "White"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ContentBackground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($ContentBackground, $RuntimeParameter)
        

        # FontFamily
        $FontFamily = 'FontFamily'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute)  
        $arrSet = [System.Drawing.FontFamily]::Families | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
        $AttributeCollection.Add($ValidateSetAttribute)
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($FontFamily, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($FontFamily, $RuntimeParameter)
        $PSBoundParameters.FontFamily = "Segui"

        # TitleFontWeight
        $TitleFontWeight = 'TitleFontWeight'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Windows.FontWeights] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.TitleFontWeight = "Normal"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TitleFontWeight, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($TitleFontWeight, $RuntimeParameter)

        # ContentFontWeight
        $ContentFontWeight = 'ContentFontWeight'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Windows.FontWeights] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.ContentFontWeight = "Normal"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ContentFontWeight, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($ContentFontWeight, $RuntimeParameter)
        

        # ContentTextForeground
        $ContentTextForeground = 'ContentTextForeground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.ContentTextForeground = "Black"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ContentTextForeground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($ContentTextForeground, $RuntimeParameter)

        # TitleTextForeground
        $TitleTextForeground = 'TitleTextForeground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.TitleTextForeground = "Black"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TitleTextForeground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($TitleTextForeground, $RuntimeParameter)

        # BorderBrush
        $BorderBrush = 'BorderBrush'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.BorderBrush = "Black"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($BorderBrush, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($BorderBrush, $RuntimeParameter)


        # TitleBackground
        $TitleBackground = 'TitleBackground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.TitleBackground = "White"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TitleBackground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($TitleBackground, $RuntimeParameter)

        # ButtonTextForeground
        $ButtonTextForeground = 'ButtonTextForeground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.ButtonTextForeground = "Black"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ButtonTextForeground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($ButtonTextForeground, $RuntimeParameter)

        # Sound
        $Sound = 'Sound'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        #$ParameterAttribute.Position = 14
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = (Get-ChildItem "$env:SystemDrive\Windows\Media" -Filter Windows* | Select -ExpandProperty Name).Replace('.wav','')
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($Sound, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($Sound, $RuntimeParameter)

        return $RuntimeParameterDictionary
    }

    Begin {
        Add-Type -AssemblyName PresentationFramework
    }
    
    Process {

# Define the XAML markup
[XML]$Xaml = @"
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        x:Name="Window" Title="" SizeToContent="WidthAndHeight" WindowStartupLocation="CenterScreen" WindowStyle="None" ResizeMode="NoResize" AllowsTransparency="True" Background="Transparent" Opacity="1">
    <Window.Resources>
        <Style TargetType="{x:Type Button}">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border>
                            <Grid Background="{TemplateBinding Background}">
                                <ContentPresenter />
                            </Grid>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Border x:Name="MainBorder" Margin="10" CornerRadius="$CornerRadius" BorderThickness="$BorderThickness" BorderBrush="$($PSBoundParameters.BorderBrush)" Padding="0" >
        <Border.Effect>
            <DropShadowEffect x:Name="DSE" Color="Black" Direction="270" BlurRadius="$BlurRadius" ShadowDepth="$ShadowDepth" Opacity="0.6" />
        </Border.Effect>
        <Border.Triggers>
            <EventTrigger RoutedEvent="Window.Loaded">
                <BeginStoryboard>
                    <Storyboard>
                        <DoubleAnimation Storyboard.TargetName="DSE" Storyboard.TargetProperty="ShadowDepth" From="0" To="$ShadowDepth" Duration="0:0:1" AutoReverse="False" />
                        <DoubleAnimation Storyboard.TargetName="DSE" Storyboard.TargetProperty="BlurRadius" From="0" To="$BlurRadius" Duration="0:0:1" AutoReverse="False" />
                    </Storyboard>
                </BeginStoryboard>
            </EventTrigger>
        </Border.Triggers>
        <Grid >
            <Border Name="Mask" CornerRadius="$CornerRadius" Background="$($PSBoundParameters.ContentBackground)" />
            <Grid x:Name="Grid" Background="$($PSBoundParameters.ContentBackground)">
                <Grid.OpacityMask>
                    <VisualBrush Visual="{Binding ElementName=Mask}"/>
                </Grid.OpacityMask>
                <StackPanel Name="StackPanel" >                   
                    <TextBox Name="TitleBar" IsReadOnly="True" IsHitTestVisible="False" Text="$Title" Padding="10" FontFamily="$($PSBoundParameters.FontFamily)" FontSize="$TitleFontSize" Foreground="$($PSBoundParameters.TitleTextForeground)" FontWeight="$($PSBoundParameters.TitleFontWeight)" Background="$($PSBoundParameters.TitleBackground)" HorizontalAlignment="Stretch" VerticalAlignment="Center" Width="Auto" HorizontalContentAlignment="Center" BorderThickness="0"/>
                    <DockPanel Name="ContentHost" Margin="0,10,0,10"  >
                    </DockPanel>
                    <DockPanel Name="ButtonHost" LastChildFill="False" HorizontalAlignment="Center" >
                    </DockPanel>
                </StackPanel>
            </Grid>
        </Grid>
    </Border>
</Window>
"@

[XML]$ButtonXaml = @"
<Button xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Width="Auto" Height="30" FontFamily="Segui" FontSize="16" Background="Transparent" Foreground="White" BorderThickness="1" Margin="10" Padding="20,0,20,0" HorizontalAlignment="Right" Cursor="Hand"/>
"@

[XML]$ButtonTextXaml = @"
<TextBlock xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" FontFamily="$($PSBoundParameters.FontFamily)" FontSize="16" Background="Transparent" Foreground="$($PSBoundParameters.ButtonTextForeground)" Padding="20,5,20,5" HorizontalAlignment="Center" VerticalAlignment="Center"/>
"@

[XML]$ContentTextXaml = @"
<TextBlock xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Text="$Content" Foreground="$($PSBoundParameters.ContentTextForeground)" DockPanel.Dock="Right" HorizontalAlignment="Center" VerticalAlignment="Center" FontFamily="$($PSBoundParameters.FontFamily)" FontSize="$ContentFontSize" FontWeight="$($PSBoundParameters.ContentFontWeight)" TextWrapping="Wrap" Height="Auto" MaxWidth="500" MinWidth="50" Padding="10"/>
"@

    # Load the window from XAML
    $Window = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $xaml))

    # Custom function to add a button
    Function Add-Button {
        Param($Content)
        $Button = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $ButtonXaml))
        $ButtonText = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $ButtonTextXaml))
        $ButtonText.Text = "$Content"
        $Button.Content = $ButtonText
        $Button.Add_MouseEnter({
            $This.Content.FontSize = "17"
        })
        $Button.Add_MouseLeave({
            $This.Content.FontSize = "16"
        })
        $Button.Add_Click({
			$VariableProxy.Acceptance = $($This.Content.Text)
			New-Variable -Name WPFMessageBoxOutput -Value $($This.Content.Text) -Option ReadOnly -Force
            $Window.Close()
        })
        $Window.FindName('ButtonHost').AddChild($Button)
    }

    # Add buttons
    If ($ButtonType -eq "OK")
    {
        Add-Button -Content "OK"
    }

    If ($ButtonType -eq "OK-Cancel")
    {
        Add-Button -Content "OK"
        Add-Button -Content "Cancel"
    }

    If ($ButtonType -eq "Abort-Retry-Ignore")
    {
        Add-Button -Content "Abort"
        Add-Button -Content "Retry"
        Add-Button -Content "Ignore"
    }

    If ($ButtonType -eq "Yes-No-Cancel")
    {
        Add-Button -Content "Yes"
        Add-Button -Content "No"
        Add-Button -Content "Cancel"
    }

    If ($ButtonType -eq "Yes-No")
    {
        Add-Button -Content "Yes"
        Add-Button -Content "No"
    }

    If ($ButtonType -eq "Retry-Cancel")
    {
        Add-Button -Content "Retry"
        Add-Button -Content "Cancel"
    }

    If ($ButtonType -eq "Cancel-TryAgain-Continue")
    {
        Add-Button -Content "Cancel"
        Add-Button -Content "TryAgain"
        Add-Button -Content "Continue"
    }

    If ($ButtonType -eq "None" -and $CustomButtons)
    {
        Foreach ($CustomButton in $CustomButtons)
        {
            Add-Button -Content "$CustomButton"
        }
    }

    # Remove the title bar if no title is provided
    If ($Title -eq "")
    {
        $TitleBar = $Window.FindName('TitleBar')
        $Window.FindName('StackPanel').Children.Remove($TitleBar)
    }

    # Add the Content
    If ($Content -is [String])
    {
        # Replace double quotes with single to avoid quote issues in strings
        If ($Content -match '"')
        {
            $Content = $Content.Replace('"',"'")
        }
        
        # Use a text box for a string value...
        $ContentTextBox = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $ContentTextXaml))
        $Window.FindName('ContentHost').AddChild($ContentTextBox)
    }
    Else
    {
        # ...or add a WPF element as a child
        Try
        {
            $Window.FindName('ContentHost').AddChild($Content) 
        }
        Catch
        {
            $_
        }        
    }

    # Enable window to move when dragged
    $Window.FindName('Grid').Add_MouseLeftButtonDown({
        $Window.DragMove()
    })

    # Activate the window on loading
    If ($OnLoaded)
    {
        $Window.Add_Loaded({
            $This.Activate()
            Invoke-Command $OnLoaded
        })
    }
    Else
    {
        $Window.Add_Loaded({
            $This.Activate()
        })
    }
    

    # Stop the dispatcher timer if exists
    If ($OnClosed)
    {
        $Window.Add_Closed({
            If ($DispatcherTimer)
            {
                $DispatcherTimer.Stop()
            }
            Invoke-Command $OnClosed
        })
    }
    Else
    {
        $Window.Add_Closed({
            If ($DispatcherTimer)
            {
                $DispatcherTimer.Stop()
            }
        })
    }
    

    # If a window host is provided assign it as the owner
    If ($WindowHost)
    {
        $Window.Owner = $WindowHost
        $Window.WindowStartupLocation = "CenterOwner"
    }

    # If a timeout value is provided, use a dispatcher timer to close the window when timeout is reached
    If ($Timeout)
    {
        $Stopwatch = New-object System.Diagnostics.Stopwatch
        $TimerCode = {
            If ($Stopwatch.Elapsed.TotalSeconds -ge $Timeout)
            {
                $Stopwatch.Stop()
                $Window.Close()
            }
        }
        $DispatcherTimer = New-Object -TypeName System.Windows.Threading.DispatcherTimer
        $DispatcherTimer.Interval = [TimeSpan]::FromSeconds(1)
        $DispatcherTimer.Add_Tick($TimerCode)
        $Stopwatch.Start()
        $DispatcherTimer.Start()
    }

    # Play a sound
    If ($($PSBoundParameters.Sound))
    {
        $SoundFile = "$env:SystemDrive\Windows\Media\$($PSBoundParameters.Sound).wav"
        $SoundPlayer = New-Object System.Media.SoundPlayer -ArgumentList $SoundFile
        $SoundPlayer.Add_LoadCompleted({
            $This.Play()
            $This.Dispose()
        })
        $SoundPlayer.LoadAsync()
    }

    # Display the window
    $null = $window.Dispatcher.InvokeAsync{$window.ShowDialog()}.Wait()

    }
}