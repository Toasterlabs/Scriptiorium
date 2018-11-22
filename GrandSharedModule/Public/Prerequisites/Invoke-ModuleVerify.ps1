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
		If (!(Get-ChildItem -Path 'C:\Program Files\WindowsPowerShell\Modules\AzureAD\' -Recurse -Filter "AzureAD.psd1")){
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
				write-warning -Message ("Please download and install the Microsoft Azure AD Preview Module. Exiting...")
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
