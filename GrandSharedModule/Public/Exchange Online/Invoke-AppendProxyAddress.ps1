Function Invoke-AppendProxyAddress{
<#
	.SYNOPSIS
		Based upon: SetMailbox_AppendProxyAddress_FromCSV.ps1 By Thomas Ashworth
		This script will append a proxy address to an existing object.

	.DESCRIPTION
		This script will append a proxy address to an existing object.

#>
	[CmdletBinding()]
	Param(
		[Parameter(Mandatory = $true)]
		$userPrincipalName,
		[Parameter(Mandatory = $true)]
		$AdditionalProxyAddress
	)

	# Validation
	If($userPrincipalName -notmatch "^([\w-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([\w-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$"){
		$Validation = "Failed"
		Write-Warning -Message "Invalid format: $userPrincipalName. Expected format: e.g. John@Contoso.com"
	}

	If($AdditionalProxyAddress -notmatch "^([\w-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([\w-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$"){
		$Validation = "Failed"
		Write-Warning -Message "Invalid format: $AdditionalProxyAddress. Expected format: e.g. John.Doe@Contoso.com"
	}

	If($validation -eq "Failed"){
		Write-Verbose "Validation failed. Skipping adding $AdditionalProxyAddress to $userPrincipalName"
		$result = $false
	}Else{
		Write-Verbose "Adding $AdditionalProxyAddress to $userPrincipalName"
		$proxyAddresses = Get-Mailbox $user.UserPrincipalName | Select-Object emailaddresses
		[Void]$proxyAddresses.EmailAddresses.Add($user.AdditionalProxyAddress)
		Set-Mailbox $user.UserPrincipalName -EmailAddresses ($proxyAddresses.emailaddresses)
		$result = $true
	}

	return $result 
}