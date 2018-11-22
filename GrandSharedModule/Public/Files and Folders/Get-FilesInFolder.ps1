function Get-FilesInFolder {
    [CmdletBinding()]
    param(
        [System.IO.FileInfo]$Folderpath,
        [string]$Extension,
		[Switch]$Recurse
    )
    
	# Objects for storing information
	$ReturnFiles = @()

	# Getting files
	If($Extension -and $Recurse){
		$Files = Get-ChildItem -Path $Folderpath -Filter $Extension -Recurse
	}Elseif($Extension){
		$Files = Get-ChildItem -Path $Folderpath -Filter $Extension
	}ElseIf($Recurse){
		$Files = Get-ChildItem -Path $Folderpath -Recurse
	}Else{
		$Files = Get-ChildItem -Path $Folderpath
	}

	# Processing each
    foreach ($File in $Files) {
		# Obj for processing files
		$objTemp = [PSCustomObject]
		$objtemp = "" | Select Name, Path

		# Populating
		$objTemp.Name = $File.name
        $objtemp.Path = $File.VersionInfo.FileName

		# Adding to hierarchy
		$ReturnFiles += $objTemp
    }
    return $ReturnFiles
}