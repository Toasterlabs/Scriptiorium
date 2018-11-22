Function Get-OSVersion{
	# Retrieving
	Write-Verbose "Retrieving OS Version"
	$OSVersion = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ProductName

	# Return
	return $OSVersion
}


	