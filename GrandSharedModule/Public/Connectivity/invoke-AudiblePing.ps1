function invoke-AudiblePing{
    <#
        .SYNOPSIS
          Issues a audible beep if the host is down
        .DESCRIPTION
          Uses test-connection to see if a host is up. If the host is down, it will issue a console beep to alert the user. It will ping indefinitely, untill cancelled out 
        .PARAMETER TARGET
            device that will be pinged
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
          invoke-AudiblePing -target www.microsoft.com
    #>
    
    Param(
        [parameter(Mandatory=$true,HelpMessage="Device to Ping")]$target 
    )
    
    Do{
        Try{
            Test-Connection -ComputerName $target
        }Catch{ 
             Write-Output "Unable to ping $target"    
            [console]::Beep(500,300)
            start-sleep -Seconds 1
        } 
    }while($true)
}
