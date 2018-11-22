Function Invoke-Countdown{
	Param(
		[Parameter(ParameterSetName = "Timer")]
		[SWITCH]$Timer,
		[Parameter(ParameterSetName = "Timer")]
		[int]$Days,
		[Parameter(ParameterSetName = "Timer")]
		[int]$Hours,
		[Parameter(ParameterSetName = "Timer")]
		[int]$Minutes,
		[Parameter(ParameterSetName = "Timer")]
		[int]$Seconds,
		[Parameter(ParameterSetName = "Timer")]
		[Parameter(ParameterSetName = "Format")]
		[SWITCH]$Format,
		[Parameter(ParameterSetName = "Format")]
		[Parameter(ParameterSetName = "Timer")]
		[ValidateSet('Transparent', 'Black', 'DarkBlue', 'DarkGreen', 'DarkCyan', 'DarkRed', 'DarkMagenta', 'DarkYellow', 'Gray', 'DarkGray', 'Blue', 'Green', 'Cyan', 'Red', 'Magenta', 'Yellow', 'White')]
		[STRING]$ProgressBackground,
		[Parameter(ParameterSetName = "Format")]
		[Parameter(ParameterSetName = "Timer")]
		[ValidateSet('Black', 'DarkBlue', 'DarkGreen', 'DarkCyan', 'DarkRed', 'DarkMagenta', 'DarkYellow', 'Gray', 'DarkGray', 'Blue', 'Green', 'Cyan', 'Red', 'Magenta', 'Yellow', 'White')]
		[STRING]$ProgressForeground,
		[Parameter(ParameterSetName = "Timer")]
		[SWITCH]$RandomFormat,
		[Parameter(ParameterSetName = "Timer")]
		[SWITCH]$RandomMessage
	)

	# Hash Variable for random colors
	$Colors = @(
			'Black', 
			'DarkBlue', 
			'DarkGreen', 
			'DarkCyan', 
			'DarkRed', 
			'DarkMagenta', 
			'DarkYellow', 
			'Gray', 
			'DarkGray', 
			'Blue', 
			'Green', 
			'Cyan', 
			'Red', 
			'Magenta', 
			'Yellow', 
			'White'
			)

	# Hash Variable for random messages
	$loading = @(
			"I'm about to drop the hammer...",
			'May I take your order?',
			"I'm listening...",
			'Strap yourselves in, boys.',
			'When removing your overhead luggage, please be careful.',
			'In case of a water landing, you may be used as a flotation device.',
			'Keep your arms and legs inside until this ride comes to a full and complete stop.',
			'E=mc . . . doh, let me get my notepad!',
			'Who set all these lab monkeys free?',
			'I think we may have a gas leak...',
			'Hailing frequencies open...',
			'Battlecruiser operational!',
			"I can't believe they put me in one of these things",
			"If it weren't for these neural implants...",
			'Is something burning?',
			'Milsat EE209 on.'
			'Waiting for someone to hit enter',
            'Warming up processors', 
            'Downloading the Internet', 
            'Trying common passwords', 
            'Commencing infinite loop', 
            'Injecting double negatives', 
            'Breeding bits', 
            'Capturing escaped bits', 
            'Dreaming of electric sheep', 
            'Calculating gravitational constant', 
            'Adding Hidden Agendas', 
            'Adjusting Bell Curves', 
            'Aligning Covariance Matrices', 
            'Attempting to Lock Back-Buffer', 
            'Building Data Trees', 
            'Calculating Inverse Probability Matrices', 
            'Calculating Llama Expectoration Trajectory', 
            'Compounding Inert Tessellations', 
            'Concatenating Sub-Contractors', 
            'Containing Existential Buffer', 
            'Deciding What Message to Display Next', 
            'Increasing Accuracy of RCI Simulators', 
            'Perturbing Matrices',
            'Initializing flux capacitors',
            'Brushing up on my Dothraki',
            'Preparing second breakfast',
            'Preparing the jump to lightspeed',
            'Initiating self-destruct sequence',
            'Mining cryptocurrency',
            'Aligning Heisenberg compensators',
            'Setting phasers to stun',
            'Deciding...blue pill or yellow?',
            'Bringing Skynet online',
            'Learning PowerShell',
            'On hold with Comcast customer service',
            'Waiting for Godot',
            'Folding proteins',
            'Searching for infinity stones',
            'Restarting the ARC reactor',
            'Learning regular expressions',
            'Trying to quit vi',
            'Waiting for the last Game_of_Thrones book',
            'Watching paint dry',
            'Aligning warp coils',
			'Drop your weapon. You have fifteen seconds to comply. . . .  5, 4, 3, 2, 1, [ZAP]',
			'Your thoughts betray you.',
			'NO. I am your father!',
			"Nelson Mandela died in prison in the 1980’s"
        )

	# Saving Default color scheme
	$SavedBgrColor = $host.PrivateData.ProgressBackgroundColor 
	$SavedFgrColor = $Host.PrivateData.ProgressForegroundColor

	# Region Timer
	If($Timer){

		# Setting StartTime
		$startTime = get-date
		
		# Setting EndTime By adding
		$EndTime = ($startTime).AddDays($Days) 
		$EndTime = $EndTime.AddHours($Hours)
		$EndTime = $EndTime.AddSeconds($Seconds)

		# Timespan
		$TimeSpan = (New-TimeSpan -Start $startTime -End $endTime)

		If($Format){
			# Setting Colors
			Try{
				If($ProgressBackground -eq "Transparent"){
					$Host.PrivateData.ProgressBackgroundColor = $host.ui.RawUI.BackgroundColor
				}Else{
					$Host.PrivateData.ProgressBackgroundColor = $ProgressBackground
				}
				$Host.PrivateData.ProgressForegroundColor = $ProgressForeground
			}Catch{
				Write-Verbose "Format specified but no colors defined. Sticking to the originals..."
			}
		}

		Do{
			# Random colors
			If($RandomFormat){

				# Picking random colors
				$ProgressBackground = $colors | Get-Random
				$ProgressForeground = $colors | Get-Random

				# Avoiding Identical colors
				If($ProgressBackground -eq $ProgressForeground){$ProgressForeground = $colors | Get-Random}

				# Setting Colors
				$Host.PrivateData.ProgressBackgroundColor = $ProgressBackground
				$Host.PrivateData.ProgressForegroundColor = $ProgressForeground
			}

			# Random message
			If($RandomMessage){
				$RandomloadingMessage = $loading[(Get-Random -Minimum 0 -Maximum ($loading.Length - 1))]
			}

			# Getting current time
			$Now = Get-Date
			
			# Setting elapsed timespan
			$elapsed = (New-TimeSpan -Start $startTime -End $now)
			
			# Remaining Time
			$remaining = $TimeSpan - $elapsed

			# How far in to the timespan we are
			$InPercent = ($elapsed.TotalSeconds/$TimeSpan.TotalSeconds) * 100
			
			# Display in Progress bar
			## Catching error that InPercent is bigger than 100
			Try{
				If($RandomMessage){
					Write-Progress -id 0 -Activity $RandomloadingMessage -Status "Remaining: $($remaining.Days) Days, $($remaining.Hours) Hours, $($remaining.Seconds) Seconds"  -PercentComplete $InPercent
				}Else{
					Write-Progress -id 0 -Activity "Countdown" -Status "Remaining: $($remaining.Days) Days, $($remaining.Hours) Hours, $($remaining.Seconds) Seconds"  -PercentComplete $InPercent
				}
			}Catch{
				If($RandomMessage){
					Write-Progress -id 0 -Activity $RandomloadingMessage -Status "Remaining: $($remaining.Days) Days, $($remaining.Hours) Hours, $($remaining.Seconds) Seconds"  -PercentComplete 100
				}Else{
					Write-Progress -id 0 -Activity "Countdown" -Status "Remaining: $($remaining.Days) Days, $($remaining.Hours) Hours, $($remaining.Seconds) Seconds"  -PercentComplete 100
				}
			}
			# Sleeping 1 second
			Start-Sleep -Seconds 1
		}Until($now -gt $EndTime)

		# Returning default color scheme
		$host.PrivateData.ProgressBackgroundColor = $SavedBgrColor
		$Host.PrivateData.ProgressForegroundColor = $SavedFgrColor
	}
	#endregion
}