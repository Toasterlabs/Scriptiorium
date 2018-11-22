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