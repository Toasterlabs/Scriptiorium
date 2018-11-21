function invoke-SpeakingPing{
    <#
        .SYNOPSIS
          Talks to you and tells you if a host is responding to ping... or not...
        .DESCRIPTION
          Uses test-connection to see if a host is up. It will speak to rell you if the host is responding to ping or not. It will ping indefinitely, untill cancelled out 
        .PARAMETER TARGET
            device that will be pinged
        .PARAMETER Voice
            The voice that will be used to speak. Currently you have a choice between David (Male) and Zira (Female)
        .INPUTS
          none
        .OUTPUTS
          none
        .NOTES
          Version:        1.0
          Author:         Marc Dekeyser
          Creation Date:  Juli 7th, 2018
          Purpose/Change: Just having some fun
  
        .EXAMPLE
          invoke-SpeakingPing -target www.microsoft.com
    #>
    
    Param(
        [parameter(Mandatory=$true,HelpMessage="Device to Ping")]$target,
        [ValidateSet("David","Zira")]
        [parameter(Mandatory=$false,HelpMessage="Device to Ping")]$Voice = "David" 
    )
    
    Add-Type -AssemblyName System.speech
    $speakup = New-Object System.Speech.Synthesis.SpeechSynthesizer

    Switch($Voice){
    
        "David" {$speakup.SelectVoice('Microsoft David Desktop')}

        "Zira"  {$speakup.SelectVoice('Microsoft Zira Desktop')}
    
    }

    Switch($target){

    Rohirrim    {
                    $speakup.speak("Arise! Arise! Riders of Théoden! Spears shall be shaken, shields shall be splintered! A sword-day! A red day, ere the sun rises!")
                    $speakup.Speak("Ride now, ride now, ride! Ride for ruin and the world's ending!")
                    $speakup.Speak("Death!")
                }

    Default     {
                    Do{
                        if(Test-Connection $target -Count 1 -Quiet){
                            $speakup.speak("$target is responding to ping!")
                        } Else {
                            $speakup.speak("$target is not responding to ping!")
                        }
                    }while($true)
            
                }

    }
}