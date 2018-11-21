Function Get-ADForest
{
	<#
		.SYNOPSIS
			Returns an object representing an Active Directory forest.

		.DESCRIPTION
			Returns an object representing an Active Directory forest.  If 
			a forest name is not specified, the current	user context is used.

		.PARAMETER ForestName
			Specifies the DNS name of the Active Directory forest.

		.PARAMETER Credential
			Specifies the username and password required to perform the operation.

		.EXAMPLE
			PS> GetADForest

		.EXAMPLE
			PS> GetADForest -ForestName "example.contoso.com"

		.EXAMPLE
			PS> GetADForest -Credential $cred

		.EXAMPLE
			PS> GetADForest -ForestName "example.contoso.com" -Credential $cred

		.INPUTS
			System.String
			System.Management.Automation.PsCredential
	
		.OUTPUTS
			System.DirectoryService.ActiveDirectory.Forest

		.NOTES

	#>
	
	
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $False, Position = 0)]
		[string]$ForestName,
		
		[Parameter(Mandatory = $False)]
		[System.Management.Automation.PsCredential]$Credential
	)
	
	If (!$ForestName) {$ForestName = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().Name.ToString()}
	If ($Credential) {$directoryContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("forest", $ForestName, $Credential.UserName.ToString(), $Credential.GetNetworkCredential().Password.ToString())}
	Else {$directoryContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("forest", $ForestName)}
	$objForest = ([System.DirectoryServices.ActiveDirectory.Forest]::GetForest($directoryContext))
	Return $objForest
}