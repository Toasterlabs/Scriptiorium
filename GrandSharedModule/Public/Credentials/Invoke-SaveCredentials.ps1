Function Invoke-SaveCredentials{
	Param(
		[Parameter(Mandatory=$True)]
		[STRING]$ResourceName
	)

	# Setting Check to null
	$ExistentialCrisis = $null
	## Checking we already have credentials saved
	Write-Verbose -Message "Checking if the credentials have already been saved previously"
	$ExistentialCrisis = invoke-AccessPwdVault -Fetch -Resource $ResourceName -erroraction SilentlyContinue
	
	## If statement based on above result.
	if(!($ExistentialCrisis)){
		Try{
			$Credentials = $host.ui.PromptForCredential('Credentials', 'Please enter your credentials', '', '')
			invoke-AccessPwdVault -Add -Credential $Credentials -resource $ResourceName
			Write-Verbose -Message "Saved Credentials"
		}Catch{
			# Oops something went wrong
			Write-Verbose -Message "Unable to save credentials"
			# Error Output
			throw $_
		}
	}Else{
		Write-Verbose -Message "Credentials already exist. Overwriting..."
		
		# Removing credentials
		invoke-AccessPwdVault -Remove -Credential $ExistentialCrisis -Resource $ResourceName

		# Adding credentials
		Try{
			$Credentials = $host.ui.PromptForCredential('Credentials', 'Please enter your credentials', '', '')
			invoke-AccessPwdVault -Add -Credential $Credentials -resource $ResourceName
			Write-Verbose -Message "Saved Credentials"
		}Catch{
			# Oops something went wrong
			Write-Verbose -Message "Unable to save credentials"
			# Error Output
			throw $_
		}
	}
}