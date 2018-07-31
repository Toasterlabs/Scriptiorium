function Invoke-NewRunspace{
    <#
        .SYNOPSIS
          Runs code in a runspace

        .DESCRIPTION
          This function will create a new runspace with a handlename and execute the code specified

        .PARAMETER DEBUGWITHGLOBALVARIABLES
            Allows you to have the uihash variable defined as global. Handy in case you need to debug what is going on

        .PARAMETER CODEBLOCK
            The code that you want executed

        .PARAMETER RUNSPACEHANDLENAME
            The name for the runspace

        .PARAMETER PROXYVARIABLE
            In case you want a variable passed through to the runspace

        .NOTES
          Version:        1.0
          Author:         Marc Dekeyser
          Creation Date:  Juli 30th, 2018
          Purpose/Change: Just having some fun
  
        .EXAMPLE
          invoke-NewRunspace -codeblock $code -runspacehandlename "JustAnotherTest" -proxyvariable $variablehash
    #>

    [cmdletbinding()]
    param(
        [parameter(Mandatory=$false)][switch]$DebugWithGlobalVariables,
        [parameter(Mandatory=$true)][scriptblock]$codeBlock,
        [parameter(Mandatory=$true)][string]$RunspaceHandleName,
        [Parameter(Mandatory=$false)]$proxyvariable
    )

    if ($DebugWithGlobalVariables) {
        $Global:uiHash = [hashtable]::Synchronized(@{})
        $Global:proxyvariable = [hashtable]::Synchronized(@{})
    } else {
        $uiHash = [hashtable]::Synchronized(@{})
        $proxyvariable = [hashtable]::Synchronized(@{})
    }

    $newRunspace = [runspacefactory]::CreateRunspace()
    $newRunspace.ApartmentState = "STA"
    $newRunspace.ThreadOptions = "ReuseThread"
    $newRunspace.Open()
    $newRunspace.SessionStateProxy.SetVariable("uiHash",$uiHash)
    $newRunspace.SessionStateProxy.SetVariable("variableHash",$variableHash)
    $newRunspace.Name = $RunspaceHandleName

    $pscmd = [PowerShell]::Create().AddScript($codeBlock)
    
    $pscmd.Runspace = $newRunspace
    
    $data = $psCmd.BeginInvoke()
}
