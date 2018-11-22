Function Get-PerfCounter{
	[CmdletBinding()]
	Param(
		[Parameter(Mandatory=$true, HelpMessage = "Name of the counter to retrieve")]
		[ValidateNotNullOrEmpty()]
		[STRING]$CounterName,
		[Parameter(Mandatory=$False)]
		[STRING]$ComputerName = $env:COMPUTERNAME,
		[Parameter(Mandatory=$False)]
		[STRING]$Threshold
	)

	# Obj for storing information
	$objTemp = [PSCustomObject]
	$objTemp = "" | select Counter, Value, Threshold
	
	# Setting counter name
	$objTemp.Counter = $counter

	# Setting threshold
	if([STRING]::IsNullOrEmpty($Threshold)){
		$objTemp.Threshold = "N/A"
	}Else{
		$objTemp.Threshold = $threshold
	}

	# Getting counter value
	Try{
		$objTemp.Value = [MATH]::Round((Get-Counter $counter -ComputerName $ComputerName).CounterSamples[0].Cookedvalue)
	}Catch{
		$objTemp.Value = "N/A"
	}

	# Returning Obj
	Return $objTemp
}