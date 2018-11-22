<#
.SYNOPSIS
  Custom logging script

.DESCRIPTION
  Takes input in the 'DESCRIPTION: Message' format and transforms it for on-screen output, whilst also (optionally) writing it to a log file

.PARAMETER MESSAGE
    Text to be handled

.PARAMETER Runlog
    File to write output to

.NOTES
  Version:        1.0
  Author:         Marc Dekeyser
  Creation Date:  November 2018
  Purpose/Change: Script dev
  
.EXAMPLE
  Invoke-WriteLog -Message "SUCCESS: WE DID IT!" -Runlog "C:\WE\DID\IT.log"
#>

Function invoke-WriteLog{
    Param(
        [Parameter(Mandatory=$True)]
        [STRING]$Message,
        [Parameter(Mandatory=$False)]
        [System.IO.FileInfo]$RunLog
	)

		$LogStatus = $Message.Split(":")[0]

		Switch($LogStatus){
			"SUCCESS"   {$Message = (get-date -Format HH:mm:ss) + " - " + $Message ; Write-Host $Message -ForegroundColor Green;Try{$Message | Out-File $RunLog -Append -ErrorAction SilentlyContinue}Catch{<#No Error Handling... This is horrible!#>}}
			"STATUS"    {$Message = (get-date -Format HH:mm:ss) + " - " + $Message ; Write-Host $Message -ForegroundColor DarkGray;Try{$Message | Out-File $RunLog -Append -ErrorAction SilentlyContinue}Catch{<#No Error Handling... This is horrible!#>}}
			"INFO"      {$Message = (get-date -Format HH:mm:ss) + " - " + $Message ; Write-Host $Message -ForegroundColor Cyan;Try{$Message | Out-File $RunLog -Append -ErrorAction SilentlyContinue}Catch{<#No Error Handling... This is horrible!#>}}
			"ALERT"     {$Message = (get-date -Format HH:mm:ss) + " - " + $Message ; Write-Host $Message -ForegroundColor Yellow;Try{$Message | Out-File $RunLog -Append -ErrorAction SilentlyContinue}Catch{<#No Error Handling... This is horrible!#>}}
			"ERROR"     {$Message = (get-date -Format HH:mm:ss) + " - " + $Message ; Write-Host $Message -ForegroundColor Red;Try{$Message | Out-File $RunLog -Append -ErrorAction SilentlyContinue}Catch{<#No Error Handling... This is horrible!#>}}
			"AUDIT"     {$Message = (get-date -Format HH:mm:ss) + " - " + $Message ; Write-Host $Message -ForegroundColor DarkGray;Try{$Message | Out-File $RunLog -Append -ErrorAction SilentlyContinue}Catch{<#No Error Handling... This is horrible!#>}}
			"NEWLINE"   {Write-Host "";try{Write-Output "`n" | Out-File $RunLog -Append -ErrorAction SilentlyContinue}Catch{}}
			"TITLE"     {$Message = $Message.split(":")[1] ; Write-Host "" ; Write-Host $Message -ForegroundColor Green ; Try{$Message | Out-File $RunLog -Append -ErrorAction SilentlyContinue}Catch{<#No Error Handling... This is horrible!#>}}
			"TEXT"      {$Message = $Message.split(":")[1] ; Write-Host $Message -ForegroundColor Green ; Try{$Message | Out-File $RunLog -Append -ErrorAction SilentlyContinue}Catch{<#No Error Handling... This is horrible!#>}}
			"PROMPT"    {$Message = (get-date -Format HH:mm:ss) + " - " + $Message.split(":")[1] ; Write-Host $Message -ForegroundColor Yellow -NoNewline}
		}
}