Function Get-ADObject
{
	<#
		.SYNOPSIS
			Returns an object that represents an Active Directory object.

		.DESCRIPTION
			Returns an object that represents an Active Directory object.

		.PARAMETER DomainController
			Specifies the DNS name of the Active Directory domain controller to
			query for the search.

		.PARAMETER SearchRoot
			Specifies the distinguished name of the directory service location
			from which the search will begin.

		.PARAMETER SearchScope
			Specifies the scope for a directory search.
			
			Default value: subtree.

			base      :   Limits the search to only the base object.
			onelevel  :   Search is restricted to the immediate children
			              of a base object, but excludes the base object itself.
			subtree   :   Includes all of the objects beneath the base
			              object, excluding the base object itself.

		.PARAMETER Filter
			Specifies an LDAP filter to use for the search.
			Example: (&(objectcategory=person)(objectclass=user)(proxyaddresses=smtp:*))

		.PARAMETER PropertiesToLoad
			Specifies a collection of Active Directory properties to retrieve
			about the object. Separate multiple values with commas.

		.PARAMETER Credential
			Specifies the username and password required to perform the operation.

		.EXAMPLE
			PS> GetADObject -DomainController "servername.example.contoso.com" -SearchRoot "ou=organizational unit,dc=example,dc=contoso,dc=com" -SearchScope "subtree" -Filter "(&(objectcategory=person)(objectclass=user))" -PropertiesToLoad "cn, distinguishedname, userprincipalname" -Credential (Get-Credential)

		.INPUTS
			System.String
			System.Management.Automation.PsCredential
	
		.OUTPUTS
			System.DirectoryServices.DirectorySearcher

		.NOTES

	#>
	

	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $False, Position = 0, ParameterSetName = "DomainController")]
		[string]$DomainController,
		
		[Parameter(Mandatory = $False)]
		[string]$SearchRoot,
		
		[Parameter(Mandatory = $False)]
		[string]$SearchScope,
		
		[Parameter(Mandatory = $False)]
		[string]$Filter,
		
		[Parameter(Mandatory = $False)]
		$PropertiesToLoad,
		
		[Parameter(Mandatory = $False)]
		[System.Management.Automation.PsCredential]$Credential
	)

	$DirectoryEntryUserName = [string]$Credential.UserName
    $DirectoryEntryPassword = [string]$Credential.GetNetworkCredential().Password
	$AuthenticationType = [System.DirectoryServices.AuthenticationTypes]::Signing -bor [System.DirectoryServices.AuthenticationTypes]::Sealing -bor [System.DirectoryServices.AuthenticationTypes]::Secure
    
	$SearchRoot = "LDAP://{0}/{1}" -f ($DomainController, $SearchRoot)

	$objDirectoryEntry = New-Object System.DirectoryServices.DirectoryEntry($SearchRoot, `
		$DirectoryEntryUserName, `
		$DirectoryEntryPassword, `
		$AuthenticationType)

	$objDirectorySearcher = New-Object System.DirectoryServices.DirectorySearcher
	$objDirectorySearcher.SearchRoot = $objDirectoryEntry
	$objDirectorySearcher.SearchScope = $SearchScope
	$objDirectorySearcher.PageSize = 1000
	$objDirectorySearcher.ReferralChasing = "All"
	$objDirectorySearcher.CacheResults = $False
	$colPropertiesToLoad | ForEach-Object -Process {[Void]$objDirectorySearcher.PropertiesToLoad.Add($_)}
	$objDirectorySearcher.Filter = $Filter
	$colADObject = $objDirectorySearcher.FindAll()
    
    Return $colADObject
}