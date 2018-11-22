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