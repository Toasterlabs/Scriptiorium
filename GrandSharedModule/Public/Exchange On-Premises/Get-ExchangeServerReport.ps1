Function Get-ExchangeServerReport{
    Param($ExchangeServer)   
    # Set Basic Variables
	$MailboxCount = 0
	$RollupLevel = 0
	$RollupVersion = ""
    $ExtNames = @()
    $IntNames = @()
    $CASArrayName = ""

    # Get WMI Information
	invoke-WriteLog -Message "INFO: $($ExchangeServer.Name) - Retrieving WMI Information"
	$tWMI = Get-WmiObject Win32_OperatingSystem -ComputerName $ExchangeServer.Name -ErrorAction SilentlyContinue
	if ($tWMI)
	{
		$OSVersion = $tWMI.Caption.Replace("(R)","").Replace("Microsoft ","").Replace("Enterprise","Ent").Replace("Standard","Std").Replace(" Edition","")
		$OSServicePack = $tWMI.CSDVersion
		$RealName = $tWMI.CSName.ToUpper()
	} else {
		invoke-WriteLog -Message "ALERT: Cannot detect OS information via WMI for $($ExchangeServer.Name)"
		$OSVersion = "N/A"
		$OSServicePack = "N/A"
		$RealName = $ExchangeServer.Name.ToUpper()
	}

    # Get Exchange Version
	invoke-WriteLog -Message "INFO: $($ExchangeServer.Name) - Retrieving Version"
	if ($ExchangeServer.AdminDisplayVersion.Major -eq 6)
	{
		$ExchangeMajorVersion = "$($ExchangeServer.AdminDisplayVersion.Major).$($ExchangeServer.AdminDisplayVersion.Minor)"
		$ExchangeSPLevel = $ExchangeServer.AdminDisplayVersion.FilePatchLevelDescription.Replace("Service Pack ","")
	} elseif ($ExchangeServer.AdminDisplayVersion.Major -eq 15 -and $ExchangeServer.AdminDisplayVersion.Minor -eq 1) {
        $ExchangeMajorVersion = [double]"$($ExchangeServer.AdminDisplayVersion.Major).$($ExchangeServer.AdminDisplayVersion.Minor)"
        $ExchangeSPLevel = 0
    } else {
		$ExchangeMajorVersion = $ExchangeServer.AdminDisplayVersion.Major
		$ExchangeSPLevel = $ExchangeServer.AdminDisplayVersion.Minor
	}
	# Exchange 2007+
	invoke-WriteLog -Message "INFO: $($ExchangeServer.Name) - Retrieving Roles, URI, Rollups/CU level"
	if ($ExchangeMajorVersion -ge 8)
	{
		# Get Roles
        $Roles = @()
        Foreach ($i in ($ExchangeServer.ServerRole.ToString().Replace(" ","").Split(","))){$roles += $i}
		
        # Get HTTPS Names (Exchange 2010 only due to time taken to retrieve data)
        if ($Roles -contains "ClientAccess")
        {
            
            Get-OWAVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | %{ $ExtNames+=[STRING]$_.ExternalURL.Host; $IntNames+=[STRING]$_.InternalURL.Host; }
            Get-WebServicesVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | %{ $ExtNames+=[STRING]$_.ExternalURL.Host; $IntNames+=[STRING]$_.InternalURL.Host; }
            Get-OABVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | %{ $ExtNames+=[STRING]$_.ExternalURL.Host; $IntNames+=[STRING]$_.InternalURL.Host; }
            Get-ActiveSyncVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | %{ $ExtNames+=[STRING]$_.ExternalURL.Host; $IntNames+=[STRING]$_.InternalURL.Host; }
            if (Get-Command Get-MAPIVirtualDirectory -ErrorAction SilentlyContinue)
            {
                Get-MAPIVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | %{ $ExtNames+=[STRING]$_.ExternalURL.Host; $IntNames+=[STRING]$_.InternalURL.Host; }
            }
            if (Get-Command Get-ClientAccessService -ErrorAction SilentlyContinue)
            {
                $IntNames+=[STRING](Get-ClientAccessService -Identity $ExchangeServer.Name).AutoDiscoverServiceInternalURI.Host
            } else {
                $IntNames+=[STRING](Get-ClientAccessServer -Identity $ExchangeServer.Name).AutoDiscoverServiceInternalURI.Host
            }
            
            if ($ExchangeMajorVersion -ge 14)
            {
                Get-ECPVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | %{ $ExtNames+=$_.ExternalURL.Host; $IntNames+=$_.InternalURL.Host; }
            }
            $IntNames = $IntNames|Sort-Object -Unique
            $ExtNames = $ExtNames|Sort-Object -Unique
            $CASArray = Get-ClientAccessArray -Site $ExchangeServer.Site.Name
            $CASArrayName = [STRING]$CASArray.Fqdn
        }

		# Rollup Level / Versions (Thanks to Bhargav Shukla http://bit.ly/msxGIJ)
		if ($ExchangeMajorVersion -ge 14) {
            $RegKey="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\AE1D439464EB1B8488741FFA028E291C\\Patches"
        
        }else{
			$RegKey="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\461C2B4266EDEF444B864AD6D9E5B613\\Patches"
		}
		$RemoteRegistry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ExchangeServer.Name);
		if ($RemoteRegistry)
		{
			$RUKeys = $RemoteRegistry.OpenSubKey($RegKey).GetSubKeyNames() | ForEach {"$RegKey\\$_"}
			if ($RUKeys)
			{
				[array]($RUKeys | %{$RemoteRegistry.OpenSubKey($_).getvalue("DisplayName")}) | %{
					if ($_ -like "Update Rollup *")
					{
						$tRU = $_.Split(" ")[2]
						if ($tRU -like "*-*") { $tRUV=$tRU.Split("-")[1]; $tRU=$tRU.Split("-")[0] } else { $tRUV="" }
						if ([int]$tRU -ge [int]$RollupLevel) { $RollupLevel=$tRU; $RollupVersion=$tRUV }
					}
				}
			}
        } else {
			Write-Warning "Cannot detect Rollup Version via Remote Registry for $($ExchangeServer.Name)"
		}
        # Exchange 2013 CU or SP Level
        if ($ExchangeMajorVersion -ge 15)
		{
			$RegKey="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\Microsoft Exchange v15"
		    $RemoteRegistry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ExchangeServer.Name);
		    if ($RemoteRegistry)
		    {
			    $ExchangeSPLevel = $RemoteRegistry.OpenSubKey($RegKey).getvalue("DisplayName")
                if ($ExchangeSPLevel -like "*Service Pack*" -or $ExchangeSPLevel -like "*Cumulative Update*")
                {
			        $ExchangeSPLevel = $ExchangeSPLevel.Replace("Microsoft Exchange Server 2013 ","");
                    $ExchangeSPLevel = $ExchangeSPLevel.Replace("Microsoft Exchange Server 2016 ","");
                    $ExchangeSPLevel = $ExchangeSPLevel.Replace("Service Pack ","SP");
                    $ExchangeSPLevel = $ExchangeSPLevel.Replace("Cumulative Update ","CU"); 
                } else {
                    $ExchangeSPLevel = 0;
                }
            } else {
			    Write-Warning "Cannot detect CU/SP via Remote Registry for $($ExchangeServer.Name)"
		    }
        }
	}
    
    $VersionNumbers = $ExchangeServer.AdminDisplayVersion
    $versionNumber = ($VersionNumbers.Major).ToString("00") + "." + ($VersionNumbers.Minor).ToString("00") + "." + ($VersionNumbers.Build).ToString("0000")  + "." + ($VersionNumbers.Revision).ToString("000")

    $EXInfoObject = New-Object system.collections.arraylist
    $ExInfoObject = "" | Select Name,RealName,AdminDisplayVersion,ExchangeMajorVersion,ExchangeSPLevel,Edition,OSVersion,OSServicePack,Roles,RollupLevel,RollupVersion,Site,IntNames,ExtNames,CASArrayName
	# Return Hashtable
	$ExInfoObject.Name					= $ExchangeServer.Name.ToUpper()
	$ExInfoObject.RealName				= $RealName
	$ExInfoObject.ExchangeMajorVersion 	= $ExchangeMajorVersion
	$ExInfoObject.ExchangeSPLevel		= $ExchangeSPLevel
	$ExInfoObject.Edition				= $ExchangeServer.Edition
	$ExInfoObject.OSVersion				= $OSVersion
	$ExInfoObject.OSServicePack			= $OSServicePack
	$ExInfoObject.Roles					= ($Roles | Out-String)
	$ExInfoObject.RollupLevel			= $RollupLevel
	$ExInfoObject.RollupVersion			= $RollupVersion
	$ExInfoObject.Site					= $ExchangeServer.Site.Name
	$ExInfoObject.IntNames				= ($IntNames  | Out-String)
    $ExInfoObject.ExtNames				= ($ExtNames  | Out-String)
    $ExInfoObject.CASArrayName			= $CASArrayName
    $EXInfoObject.AdminDisplayVersion   = $versionNumber
		
    return $EXInfoObject
}