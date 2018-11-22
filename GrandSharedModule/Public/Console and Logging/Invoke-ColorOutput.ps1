<#
.SYNOPSIS
  Custom write-output to enable colorization without using Write-Host (A puppy dies everytime you use Write-host... 
  Plus it's text based instead of the glorious object based <thonk>)

.DESCRIPTION
  Takes input in the 'DESCRIPTION: Message' format and transforms it to a colorized format for on-screen output

.PARAMETER MESSAGE
    Object you want to output to the screen in a colorized fashion

.NOTES
  Version:        1.0
  Author:         Marc Dekeyser
  Creation Date:  November 2018
  Purpose/Change: Script dev
  
.EXAMPLE
  Invoke-ColorOutput -Message "WE DID IT!!" -Fore Green -Back Yellow
#>

function Invoke-ColorOutput{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$False,Position=1,ValueFromPipeline=$True,ValueFromPipelinebyPropertyName=$True)]
		[alias('message')]
		[alias('msg')]
		[Object]$Object,
        [Parameter(Mandatory=$False,Position=2,ValueFromPipeline=$True,ValueFromPipelinebyPropertyName=$True)]
		[ValidateSet('Black', 'DarkBlue', 'DarkGreen', 'DarkCyan', 'DarkRed', 'DarkMagenta', 'DarkYellow', 'Gray', 'DarkGray', 'Blue', 'Green', 'Cyan', 'Red', 'Magenta', 'Yellow', 'White')]
		[alias('fore')]
		[ConsoleColor] $ForegroundColor,
        [Parameter(Mandatory=$False,Position=3,ValueFromPipeline=$True,ValueFromPipelinebyPropertyName=$True)]
		[ValidateSet('Black', 'DarkBlue', 'DarkGreen', 'DarkCyan', 'DarkRed', 'DarkMagenta', 'DarkYellow', 'Gray', 'DarkGray', 'Blue', 'Green', 'Cyan', 'Red', 'Magenta', 'Yellow', 'White')]
		[alias('back')]
		[alias('BGR')]
		[ConsoleColor] $BackgroundColor,
        [Switch]$NoNewline
    )    

    # Save previous colors
    $previousForegroundColor = $host.UI.RawUI.ForegroundColor
    $previousBackgroundColor = $host.UI.RawUI.BackgroundColor

    # Set BackgroundColor if available
    if($BackgroundColor -ne $null)
    { 
       $host.UI.RawUI.BackgroundColor = $BackgroundColor
    }

    # Set $ForegroundColor if available
    if($ForegroundColor -ne $null)
    {
        $host.UI.RawUI.ForegroundColor = $ForegroundColor
    }

    # Always write (if we want just a NewLine)
    if($Object -eq $null)
    {
        $Object = ""
    }

    if($NoNewline)
    {
        [Console]::Write($Object)
    }
    else
    {
        Write-Output $Object
    }

    # Restore previous colors
    $host.UI.RawUI.ForegroundColor = $previousForegroundColor
    $host.UI.RawUI.BackgroundColor = $previousBackgroundColor
}