Function Invoke-NewPopup {
<#
        .SYNOPSIS
          Creates a pop up message
        .DESCRIPTION
          Allows you to create a fully customized pop up message.
        .PARAMETER MESSAGE
            Main text of the popup
        .PARAMETER TITLE
            Pop up title
        .PARAMETER TIME
            How long should the pop up be displayed? By default a button click is required to close it
        .PARAMETER BUTTONS
            What button group to display. Default is a single OK button
        .PARAMETER ICON
            What icon to display. Default is INFORMATION
        .PARAMETER DefaultButton
            Which button should be selected as the default
        .INPUTS
          none
        .OUTPUTS
          none
        .NOTES
          Version:        1.0
          Author:         Marc Dekeyser
          Creation Date:  Juli 7th, 2018
          Purpose/Change: Just having some fun
    #>

Param (
[Parameter(Position=0,Mandatory=$True,HelpMessage="Enter a message for the popup")]
[ValidateNotNullorEmpty()]
[string]$Message,
[Parameter(Position=1,Mandatory=$True,HelpMessage="Enter a title for the popup")]
[ValidateNotNullorEmpty()]
[string]$Title,
[Parameter(Position=2,HelpMessage="How many seconds to display? Use 0 require a button click.")]
[ValidateScript({$_ -ge 0})]
[int]$Time=0,
[Parameter(Position=3,HelpMessage="Enter a button group")]
[ValidateNotNullorEmpty()]
[ValidateSet("OK","OKCancel","AbortRetryIgnore","YesNo","YesNoCancel","RetryCancel","CancelTryAgainCOntinue")]
[string]$Buttons="OK",
[Parameter(Position=4,HelpMessage="Enter an icon set")]
[ValidateNotNullorEmpty()]
[ValidateSet("Stop","Question","Exclamation","Information" )]
[string]$Icon="Information",
[Parameter(HelpMessage="Select which button will be selected as the default")]
[ValidateSet("Second","Third")]
[string]$DefaultButton

)

#convert buttons to their integer equivalents
Switch ($Buttons) {
    "OK"               {$ButtonValue = 0}
    "OKCancel"         {$ButtonValue = 1}
    "AbortRetryIgnore" {$ButtonValue = 2}
    "YesNo"            {$ButtonValue = 4}
    "YesNoCancel"      {$ButtonValue = 3}
    "RetryCancel"      {$ButtonValue = 5}
    "CancelTryAgainCOntinue" {$ButtonValue = 6}
}

#set an integer value for Icon type
Switch ($Icon) {
    "Stop"        {$iconValue = 16}
    "Question"    {$iconValue = 32}
    "Exclamation" {$iconValue = 48}
    "Information" {$iconValue = 64}
}

Switch($DefaultButton){
    "Second" {$Selectedbtn = 256}
    "Third"  {$Selectedbtn = 512}

}

#create the COM Object
Try {
    $wshell = New-Object -ComObject Wscript.Shell -ErrorAction Stop
    #Button and icon type values are added together to create an integer value
    $wshell.Popup($Message,$Time,$Title,$ButtonValue + $iconValue + $Selectedbtn)
}
Catch {
    #You should never really run into an exception in normal usage
    Write-Warning "Failed to create Wscript.Shell COM object"
    Write-Warning $_.exception.message
}

}
