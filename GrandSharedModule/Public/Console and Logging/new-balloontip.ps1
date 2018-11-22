Function new-balloontip {
Param(
    [Parameter(Position=0,Mandatory=$True,HelpMessage="Enter a title for the balloon tip")]
    [ValidateNotNullorEmpty()]
    [string]$Title,
    [Parameter(Position=1,Mandatory=$True,HelpMessage="Enter a message for the balloon tip")]
    [ValidateNotNullorEmpty()]
    [string]$message,
    [Parameter(Position=2,HelpMessage="Enter a button group")]
    [ValidateSet("Error","Info","None","Warning")]
    [string]$ToolTipIcon="None",
    [Parameter(Position=3,HelpMessage="Duration for the balloon tip to be visable in milliseconds")]
    [int]$duration = 5000
)
Add-Type -AssemblyName System.Windows.Forms

# Have to check if it already exisist, otherwise we get errors!
if(!$global:balloon){
    $global:balloon = New-Object System.Windows.Forms.NotifyIcon

    # Double click to remove the balloontip
    [void](Register-ObjectEvent  -InputObject $balloon  -EventName MouseDoubleClick  -SourceIdentifier IconClicked  -Action {

  #Perform  cleanup actions on balloon tip
  $global:balloon.dispose()

  Unregister-Event  -SourceIdentifier IconClicked
  Remove-Job -Name IconClicked
  Remove-Variable  -Name balloon  -Scope Global

})
}

# This will take the icon of whatever process is launching it
$path = (Get-Process -id $pid).Path
$balloon.Icon  = [System.Drawing.Icon]::ExtractAssociatedIcon($path)

$balloon.BalloonTipText = $message
$balloon.BalloonTipTitle = $Title

switch($ToolTipIcon){
    "Error"   {$balloon.BalloonTipIcon  = [System.Windows.Forms.ToolTipIcon]::Error}
    "Info"    {$balloon.BalloonTipIcon  = [System.Windows.Forms.ToolTipIcon]::Info}
    "None"    {$balloon.BalloonTipIcon  = [System.Windows.Forms.ToolTipIcon]::None}
    "Warning" {$balloon.BalloonTipIcon  = [System.Windows.Forms.ToolTipIcon]::Warning}
}

$balloon.Visible  = $true

# Time in ms to display the tooltip
$balloon.ShowBalloonTip($Duration)
}