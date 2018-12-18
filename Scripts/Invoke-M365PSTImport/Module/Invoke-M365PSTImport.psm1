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

Function Invoke-AccessPwdVault{
    [CmdletBinding(DefaultParameterSetName='Retrieve')]
    [Alias()]
    Param
    (
        # Used for removing credentials from the Vault
		[parameter(Mandatory=1,ParameterSetName='Remove')]
		[Switch]
		$remove,

		# Used for getting credentials from the Vault
        [parameter(Mandatory=1,ParameterSetName='Retrieve')]
        [switch]
        $Fetch,

        # Used for setting credentials from the Vault
        [parameter(Mandatory=1,ParameterSetName='Add')]
        [parameter(Mandatory=1,ParameterSetName='AddPass')]
        [switch]
        $Add,

        # Username/Identity
        [Parameter(Mandatory=0,ParameterSetName='Retrieve')]
        [Parameter(Mandatory=1,ParameterSetName='AddPass')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [alias('ID')]
        [string]$Identity = [string]::Empty,

        # Username/Identity
        [Parameter(Mandatory=1,ParameterSetName='AddPass')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string]$Pwd = [string]::Empty,

        # Resource tag
        [Parameter(Mandatory=0,ParameterSetName='Retrieve')]
        [parameter(Mandatory=1,ParameterSetName='Add')]
        [parameter(Mandatory=1,ParameterSetName='AddPass')]
		[parameter(Mandatory=1,ParameterSetName='Remove')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [alias('Res')]
        [string]$Resource,

        # Credentials for storing in Password Vault
        [parameter(Mandatory=1,ParameterSetName='Add')]
		[parameter(Mandatory=1,ParameterSetName='Remove')]
        [PSCredential[]]$Credential

    )

    Begin
    {
        Write-Verbose ("Loading PasswordVault Class.")
        [void][Windows.Security.Credentials.PasswordVault,Windows.Security.Credentials,ContentType=WindowsRuntime]

        function Convert-PSCreds
        {
            param
            (
            [parameter(Mandatory = 1, Position = 0, ValueFromPipeline = 1, ValueFromPipelineByPropertyName = 1)]
            [Windows.Security.Credentials.PasswordCredential[]]$Credential
            )
            Write-Verbose ("Converting WindowsCredential to PSCredential")
            if ($Credential.UserName -eq [string]::Empty){ throw New-Object System.NullReferenceException }
            New-Object System.Management.Automation.PSCredential -ArgumentList ($Credential.UserName, (ConvertTo-SecureString $Credential.Password -AsPlainText -Force))

        }
        function Convert-WinCreds
        {
            param
            (
            [parameter(Mandatory = 1, Position = 0, ValueFromPipeline = 1, ValueFromPipelineByPropertyName = 1)]
            [PSCredential[]]$Credential,
            [string]$Resource
            )
            Write-Verbose ("Converting PSCredential to WindowsCredential")
            New-Object Windows.Security.Credentials.PasswordCredential -ArgumentList ($Resource, $Credential.UserName, $Credential.GetNetworkCredential().Password)
        }

    }
    Process
    {

        if ($Fetch)
        {

        try
        {
            if ($Identity -ne [string]::Empty)
            {
                Write-Verbose ("Retrieving WindowsCredential from ResourceName : '{0}', UserName : '{1}'" -f $Resource, $Identity)
                (New-Object Windows.Security.Credentials.PasswordVault).Retrieve($Resource, $Identity) | ForEach-Object { $_.RetrievePassword(); $_ } | Convert-PSCreds
            }
            else
            {
                Write-Verbose ("Retrieving All Windows Credential for ResourceName : '{0}'" -f $Resource)
                (New-Object Windows.Security.Credentials.PasswordVault).FindAllByResource($Resource) | ForEach-Object { $_.RetrievePassword(); $_ } | Convert-PSCreds
            }
        }

        catch
        {
            throw $_
        }
        }
        if($Add)
        {

            try
            {
                if ($Identity -ne [string]::Empty)
                {
                    Write-Verbose ("Set Windows Credential for UserName : '{0}'" -f $Identity)
                    $Credential = New-Object System.Management.Automation.PSCredential -ArgumentList ($Identity, (ConvertTo-SecureString $Pwd -AsPlainText -Force))
                    $winCred = $Credential | Convert-WinCreds -Resource $Resource
                    (New-Object Windows.Security.Credentials.PasswordVault).Add($winCred)
                }
                else
                {
                Write-Verbose ("Set Windows Credential for UserName : '{0}'" -f $Credential.UserName)
                $winCred = $Credential | Convert-WinCreds -Resource $Resource
                (New-Object Windows.Security.Credentials.PasswordVault).Add($winCred)
                }
            }
            catch
            {
                throw $_
            }

        }
		if($remove)
		{
		Try
		{
			Write-Verbose ("Removing All Windows Credential for ResourceName : '{0}'" -f $Resource)
			Convert-WinCreds -Resource $Resource -Credential $Credential | % {
				Write-Verbose ("Removing Windows Password Vault for ResourceName : '{0}', UserName : '{1}'" -f $ResourceName, $_.UserName)
                (New-Object Windows.Security.Credentials.PasswordVault).Remove($_)
			}
            
		}
		Catch
		{
			throw $_
		}
		}
    }
    End
    {
    }
}

Function Invoke-RemoveCredentials{
	Param(
		[Parameter(Mandatory=$True)]
		[STRING]$ResourceName
	)

	# Retrieving existing credentials (We can't delete it if we don't have the username... Password vault is weird in that way)
	Write-Verbose -Message "Retrieving $resourceName"
	Try{
		$ExistentialCrisis = invoke-AccessPwdVault -Fetch -Resource $ResourceName 
	}Catch{
		Write-Warning -Message "$ResourceName not found in Credential Manager!"
		$Result = "ResourceNameNotPresent"
	}

	# Removing credentials
	Write-Verbose -Message "Starting removal of credentials"
	Try{
		invoke-AccessPwdVault -Remove -Credential $ExistentialCrisis -Resource $ResourceName
		Write-Verbose -Message "Removed credentials for $ResourceName"
		$Result = "Removed"
	}Catch{
		$result = "ErrorRemoving"
		Write-Warning "Failed to remove $ResourceName"
		Throw $_
	}
}

Function Invoke-SaveCredentials{
	Param(
		[Parameter(Mandatory=$True)]
		[STRING]$ResourceName
	)

	# Setting Check to null
	$ExistentialCrisis = $null
	## Checking we already have credentials saved
	Write-Verbose "Checking if the credentials have already been saved previously"
	$ExistentialCrisis = invoke-AccessPwdVault -Fetch -Resource $ResourceName -erroraction SilentlyContinue
	
	## If statement based on above result.
	if(!($ExistentialCrisis)){
		Try{
			$Credentials = $host.ui.PromptForCredential("$resourceName Credentials", 'Please enter your credentials', '', '')
			invoke-AccessPwdVault -Add -Credential $Credentials -resource $ResourceName
			Write-Verbose "Saved Credentials"
		}Catch{
			# Oops something went wrong
			Write-Warning "Unable to save credentials"
			# Error Output
			throw $_
		}
	}Else{
		Write-Output -Message "Credentials already exist. Overwriting..."
		
		# Removing credentials
		invoke-AccessPwdVault -Remove -Credential $ExistentialCrisis -Resource $ResourceName

		# Adding credentials
		Try{
			$Credentials = $host.ui.PromptForCredential("$resourceName Credentials", 'Please enter your credentials', '', '')
			invoke-AccessPwdVault -Add -Credential $Credentials -resource $ResourceName
			Write-Verbose "Saved Credentials"
		}Catch{
			# Oops something went wrong
			Write-Warning "Unable to save credentials"
			# Error Output
			throw $_
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
