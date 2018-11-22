Function Invoke-Speak{
	Param(
		[Parameter(Mandatory = $true)]
		[STRING]$Message,
		[Parameter(Mandatory = $false)]
		[ValidateSet('David','Zira')]
		[STRING]$Voice = "Zira"
	)

	# Loading speach assembly
	Add-Type -AssemblyName System.speech
    $speakup = New-Object System.Speech.Synthesis.SpeechSynthesizer
	
	# Selecting Voice
	Switch($Voice){
        "David" {$speakup.SelectVoice('Microsoft David Desktop')}
        "Zira"  {$speakup.SelectVoice('Microsoft Zira Desktop')}
    }

	$speakup.speak($Message)
}