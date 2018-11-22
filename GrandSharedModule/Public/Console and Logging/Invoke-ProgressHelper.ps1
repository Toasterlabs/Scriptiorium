function Invoke-ProgressHelper{
   param (
	   [Parameter(Mandatory=$true, HelpMessage = "Step number")]
	   [ValidateNotNullOrEmpty()]
	   [INT]$StepNumber,
	   [Parameter(Mandatory=$true, HelpMessage = "Status Message")]
	   [ValidateNotNullOrEmpty()]
	   [STRING]$StatusMessage,
	   [Parameter(Mandatory=$true, HelpMessage = "Identification number (activity number)")]
	   [ValidateNotNullOrEmpty()]
	   [INT]$ID = 1,
	   [Parameter(Mandatory=$true, HelpMessage = "Activity Description")]
	   [ValidateNotNullOrEmpty()]
	   [STRING]$Activity,
	   [Parameter(Mandatory=$true, HelpMessage = "Total number of steps")]
	   [ValidateNotNullOrEmpty()]
	   [INT]$Steps
   )

	# Calculating how far along we are
	$PercentComplete = (($StepNumber / $steps) * 100)

	# Catching the rare instance where percent complete is larger than 100, thus avoiding error messages
	if($PercentComplete -gt 100){
		Write-Progress -Id $ID -Activity $Activity -Status $StatusMessage -PercentComplete 100
	}Else{
		Write-Progress -Id $ID -Activity $Activity -Status $StatusMessage -PercentComplete $PercentComplete
	}
}
