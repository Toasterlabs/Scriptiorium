[CmdletBinding()]
Param(
	[Parameter(Mandatory = $true, HelpMessage = "Name of the module to create")]
	[STRING]$ModuleName,
	[Parameter(Mandatory = $true, HelpMessage = "Path where the module will be created")]
	[System.IO.FileInfo]$ModulePath,
	[Parameter(Mandatory = $true, HelpMessage = "Path to where the Public & Private functions are stored (1 function per PS1 file!)")]
	[System.IO.FileInfo]$FunctionsPath
)

Write-Verbose "Starting module building"

# Test Path
If(!(Test-Path -Path $ModulePath)){
	Write-Warning "Path does not exist! Creating..."
	New-Item -ItemType Directory -Path $ModulePath | Out-Null
}

# Checking extension
Write-Verbose "Checking extension"
$Extension = [IO.Path]::GetExtension($ModuleName)
## No extension found
If([STRING]::IsNullOrEmpty($Extension)){
	Write-Verbose "Adding .psm1 to module name"
	$ModuleName = $ModuleName + ".psm1"
}

# Found extension but it is not our expected extension of psm1
If($Extension -ne ".psm1"){
	Write-Verbose "Found extension but it is not our expected extension of .psm1"
	$ModuleName = $ModuleName.Split(".")[0] + ".psm1"
}

# Variables
$PSM1Path = -join($ModulePath,"$ModuleName")

# Gathering list of PS1 files
Write-Verbose "Gathering functions"
$Source = Get-ChildItem -Path $FunctionsPath -File -Recurse -Filter "*.ps1"

# Getting content of each file and adding it to the module
Foreach($file in $source){
	Write-Verbose "processing $($File.Versioninfo.FileName)"

	# Grabbing content
	Write-Verbose "Grabbing content"
	$Content = Get-Content $file.VersionInfo.Filename

	# Adding content to PSM1
	Write-Verbose "Adding content to PSM1"
	$Content | Add-Content -Path $PSM1Path

	# Sleeping for 300 milliseconds to let the stream settle
	Start-Sleep -Milliseconds 300
	
}

Write-Output "Building of $ModuleName has been completed!"
Write-Output "Module is located in $ModulePath..."

