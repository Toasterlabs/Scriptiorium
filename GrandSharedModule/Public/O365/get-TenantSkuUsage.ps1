Funtion get-TenantSkuUsage{
	Param(
		[Parameter(Mandatory = $true)]
		$Credential
	)

	Connect-ProvisioningWebServiceAPI -Credential $Credential

	# get Office 365 SKU info
	WriteConsoleMessage -Message "Getting SKU information.  Please wait..." -MessageType "Information"
	$getMsolAccountSkuResults = Get-MsolAccountSku

	# iterate through the sku results
	WriteConsoleMessage -Message "Processing SKU results.  Please wait..." -MessageType "Information"
	$arrSkuData = @()
	foreach($sku in $getMsolAccountSkuResults)
	{
		$objSkuData = New-Object PSObject
		Add-Member -InputObject $objSkuData -MemberType NoteProperty -Name "AccountSkuId" -Value $sku.accountskuid
		Add-Member -InputObject $objSkuData -MemberType NoteProperty -Name "ActiveUnits" -Value $sku.activeunits
		Add-Member -InputObject $objSkuData -MemberType NoteProperty -Name "ConsumedUnits" -Value $sku.consumedunits
		Add-Member -InputObject $objSkuData -MemberType NoteProperty -Name "AvailableUnits" -Value $($sku.activeunits - $sku.consumedunits)
		Add-Member -InputObject $objSkuData -MemberType NoteProperty -Name "WarningUnits" -Value $sku.warningunits
		Add-Member -InputObject $objSkuData -MemberType NoteProperty -Name "SuspendedUnits" -Value $sku.suspendedunits
		$arrSkuData += $objSkuData
	}

	return $arrSkuData
}