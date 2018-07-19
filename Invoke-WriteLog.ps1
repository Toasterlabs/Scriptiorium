function Invoke-WriteLog{
    <#
        .SYNOPSIS
          Writes a message to log file
        .DESCRIPTION
          Uses add-content to append a message to a log file, prefacing it with a timestamp in the "ddMMyy-HHmm" format
        .PARAMETER LOGFILE
            Path to the logfile 
        .PARAMETER MESSAGE
            Message to be added to the logfile
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
          Invoke-writelog -logfile C:\temp\log.log -Message "Just testing this out"
    #>

    Param(
        [parameter(mandatory=$true,HelpMessage="Path to the logfile")]
        $logfile,
        [parameter(mandatory=$true,HelpMessage="text to be added to the logfile")]
        $message
    )

    $timestamp = Get-Date -Format "ddMMyy-HHmm"
    $entry = $timestamp + " - " + $message
    add-content -Path $logfile -value $entry

}
