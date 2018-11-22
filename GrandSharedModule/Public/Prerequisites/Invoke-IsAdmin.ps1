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