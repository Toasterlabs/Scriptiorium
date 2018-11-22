Function Invoke-AccessPwdVault{
    [CmdletBinding(DefaultParameterSetName='Retrieve')]
    [Alias()]
    Param
    (
        # Used for removing credentials from the Vault
		[parameter(Mandatory=1,ParameterSetName='Remove')]
		[Switch]
		$remove,

		# Used for getting credentials from the Vault
        [parameter(Mandatory=1,ParameterSetName='Retrieve')]
        [switch]
        $Fetch,

        # Used for setting credentials from the Vault
        [parameter(Mandatory=1,ParameterSetName='Add')]
        [parameter(Mandatory=1,ParameterSetName='AddPass')]
        [switch]
        $Add,

        # Username/Identity
        [Parameter(Mandatory=0,ParameterSetName='Retrieve')]
        [Parameter(Mandatory=1,ParameterSetName='AddPass')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [alias('ID')]
        [string]$Identity = [string]::Empty,

        # Username/Identity
        [Parameter(Mandatory=1,ParameterSetName='AddPass')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string]$Pwd = [string]::Empty,

        # Resource tag
        [Parameter(Mandatory=0,ParameterSetName='Retrieve')]
        [parameter(Mandatory=1,ParameterSetName='Add')]
        [parameter(Mandatory=1,ParameterSetName='AddPass')]
		[parameter(Mandatory=1,ParameterSetName='Remove')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [alias('Res')]
        [string]$Resource,

        # Credentials for storing in Password Vault
        [parameter(Mandatory=1,ParameterSetName='Add')]
		[parameter(Mandatory=1,ParameterSetName='Remove')]
        [PSCredential[]]$Credential

    )

    Begin
    {
        Write-Verbose ("Loading PasswordVault Class.")
        [void][Windows.Security.Credentials.PasswordVault,Windows.Security.Credentials,ContentType=WindowsRuntime]

        function Convert-PSCreds
        {
            param
            (
            [parameter(Mandatory = 1, Position = 0, ValueFromPipeline = 1, ValueFromPipelineByPropertyName = 1)]
            [Windows.Security.Credentials.PasswordCredential[]]$Credential
            )
            Write-Verbose ("Converting WindowsCredential to PSCredential")
            if ($Credential.UserName -eq [string]::Empty){ throw New-Object System.NullReferenceException }
            New-Object System.Management.Automation.PSCredential -ArgumentList ($Credential.UserName, (ConvertTo-SecureString $Credential.Password -AsPlainText -Force))

        }
        function Convert-WinCreds
        {
            param
            (
            [parameter(Mandatory = 1, Position = 0, ValueFromPipeline = 1, ValueFromPipelineByPropertyName = 1)]
            [PSCredential[]]$Credential,
            [string]$Resource
            )
            Write-Verbose ("Converting PSCredential to WindowsCredential")
            New-Object Windows.Security.Credentials.PasswordCredential -ArgumentList ($Resource, $Credential.UserName, $Credential.GetNetworkCredential().Password)
        }

    }
    Process
    {

        if ($Fetch)
        {

        try
        {
            if ($Identity -ne [string]::Empty)
            {
                Write-Verbose ("Retrieving WindowsCredential from ResourceName : '{0}', UserName : '{1}'" -f $Resource, $Identity)
                (New-Object Windows.Security.Credentials.PasswordVault).Retrieve($Resource, $Identity) | ForEach-Object { $_.RetrievePassword(); $_ } | Convert-PSCreds
            }
            else
            {
                Write-Verbose ("Retrieving All Windows Credential for ResourceName : '{0}'" -f $Resource)
                (New-Object Windows.Security.Credentials.PasswordVault).FindAllByResource($Resource) | ForEach-Object { $_.RetrievePassword(); $_ } | Convert-PSCreds
            }
        }

        catch
        {
            throw $_
        }
        }
        if($Add)
        {

            try
            {
                if ($Identity -ne [string]::Empty)
                {
                    Write-Verbose ("Set Windows Credential for UserName : '{0}'" -f $Identity)
                    $Credential = New-Object System.Management.Automation.PSCredential -ArgumentList ($Identity, (ConvertTo-SecureString $Pwd -AsPlainText -Force))
                    $winCred = $Credential | Convert-WinCreds -Resource $Resource
                    (New-Object Windows.Security.Credentials.PasswordVault).Add($winCred)
                }
                else
                {
                Write-Verbose ("Set Windows Credential for UserName : '{0}'" -f $Credential.UserName)
                $winCred = $Credential | Convert-WinCreds -Resource $Resource
                (New-Object Windows.Security.Credentials.PasswordVault).Add($winCred)
                }
            }
            catch
            {
                throw $_
            }

        }
		if($remove)
		{
		Try
		{
			Write-Verbose ("Removing All Windows Credential for ResourceName : '{0}'" -f $Resource)
			Convert-WinCreds -Resource $Resource -Credential $Credential | % {
				Write-Verbose ("Removing Windows Password Vault for ResourceName : '{0}', UserName : '{1}'" -f $ResourceName, $_.UserName)
                (New-Object Windows.Security.Credentials.PasswordVault).Remove($_)
			}
            
		}
		Catch
		{
			throw $_
		}
		}
    }
    End
    {
    }
}
