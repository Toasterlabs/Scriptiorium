Function Get-ADDomain
{
	<#
		.SYNOPSIS
			Returns an object representing an Active Directory domain.  If 
			a domain name is not specified, the current	user context is used.

		.DESCRIPTION
			Returns an object representing an Active Directory domain.  If 
			a domain name is not specified, the current	user context is used.

		.PARAMETER DomainName
			Specifies the DNS name of the Active Directory domain.

		.PARAMETER Credential
			Specifies the username and password required to perform the operation.

		.EXAMPLE
			PS> GetADDomain
		
		.EXAMPLE
			PS> GetADDomain -DomainName "example.contoso.com"

		.EXAMPLE
			PS> GetADDomain -Credential $cred

		.EXAMPLE
			PS> GetADDomain -DomainName "example.contoso.com" -Credential $cred

		.INPUTS
			System.String
			System.Management.Automation.PsCredential
	
		.OUTPUTS
			System.DirectoryService.ActiveDirectory.Domain

		.NOTES

	#>
	
	
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $False)]
		[string]$DomainName,
		
		[Parameter(Mandatory = $False)]
		[System.Management.Automation.PsCredential]$Credential
	)
	
	If (!$DomainName) {$DomainName = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name.ToString()}
	If ($Credential) {$directoryContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("domain", $DomainName, $Credential.UserName.ToString(), $Credential.GetNetworkCredential().Password.ToString())}
	Else {$directoryContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("domain", $DomainName)}
	$objDomain = ([System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($directoryContext))
	
	Return $objDomain
}