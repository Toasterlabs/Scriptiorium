<#
	.SYNOPSIS
		Based upon: ConnectProvisioningWebServiceAPI.ps1 By Thomas Ashworth
		This script will append a proxy address to an existing object.

	.DESCRIPTION
		This script will append a proxy address to an existing object.

#>

Function Connect-ProvisioningWebServiceAPI
{
	<#
		.SYNOPSIS
			Connects to the Office 365 provisioning web service API.

		.DESCRIPTION
			Connects to the Office 365 provisioning web service API.
			
			If a credential is specified, it will be used to establish a connection with the provisioning
			web service API.
			
			If a credential is not specified, an attempt is made to identify an existing connection to
			the provisioning web service API.  If an existing connection is identified, the existing
			connection is used.  If an existing connection is not identified, the user is prompted for
			credentials so that a new connection can be established.

		.PARAMETER Credential
			Specifies the credential to use when connecting to the provisioning web service API
			using Connect-MsolService.

		.EXAMPLE
			PS> ConnectProvisioningWebServiceAPI

		.EXAMPLE
			PS> ConnectProvisioningWebServiceAPI -Credential
			
		.INPUTS
			[System.Management.Automation.PsCredential]

		.OUTPUTS

		.NOTES

	#>
	
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $False)]
		[System.Management.Automation.PsCredential]$Credential
	)
	
	# if a credential was supplied, assume a new connection is intended and create a new
	# connection using specified credential
	If ($Credential)
	{
		If ((!$Credential) -or (!$Credential.Username) -or ($Credential.Password.Length -eq 0))
		{
			Write-warning -Message ("Invalid credential.  Please verify the credential and try again.")
			Exit
		}
		
		# connect to provisioning web service api
		Write-Verbose -Message "Connecting to the Office 365 provisioning web service API.  Please wait..."
		Connect-MsolService -Credential $Credential
		If($? -eq $False){WriteConsoleMessage -Message "Error while connecting to the Office 365 provisioning web service API.  Quiting..." -MessageType "Error";Exit}
	}
	Else
	{
		Write-Verbose -Message "Attempting to identify an open connection to the Office 365 provisioning web service API.  Please wait..." 
		$getMsolCompanyInformationResults = Get-MsolCompanyInformation -ErrorAction SilentlyContinue
		If (!$getMsolCompanyInformationResults)
		{
			Write-Verbose -Message "Could not identify an open connection to the Office 365 provisioning web service API." 			If (!$Credential)
			{
				$Credential = $Host.UI.PromptForCredential("Enter Credential",
					"Enter the username and password of an Office 365 administrator account.",
					"",
					"userCreds")
			}
			If ((!$Credential) -or (!$Credential.Username) -or ($Credential.Password.Length -eq 0))
			{
				Write-Verbose -Message ("Invalid credential.  Please verify the credential and try again.")
				Exit
			}
			
			# connect to provisioning web service api
			Write-Verbose -Message "Connecting to the Office 365 provisioning web service API.  Please wait..."
			Connect-MsolService -Credential $Credential
			If($? -eq $False){WriteConsoleMessage -Message "Error while connecting to the Office 365 provisioning web service API.  Quiting..." -MessageType "Error";Exit}
			$getMsolCompanyInformationResults = Get-MsolCompanyInformation -ErrorAction SilentlyContinue
			WriteConsoleMessage -Message ("Connected to Office 365 tenant named: `"{0}`"." -f $getMsolCompanyInformationResults.DisplayName) -MessageType "Information"
		}
		Else
		{
			Write-Warning -Message ("Connected to Office 365 tenant named: `"{0}`"." -f $getMsolCompanyInformationResults.DisplayName) 
		}
	}
	If (!$Script:Credential) {$Script:Credential = $Credential}
}