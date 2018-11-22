Function Invoke-FolderTest{
	Param(
		[Parameter(Mandatory = $True, HelpMessage = "Path to Folder")]
		[System.IO.FileInfo]$FolderPath
	)

	if(!(Test-Path $FolderPath)){
		Write-Verbose -Message "$Folderpath does not exist. Creating..."
		Try{
			New-Item -ItemType Directory -Path $FolderPath | Out-Null
			Write-Verbose -Message "$FolderPath Created"
			$result = $true
		}Catch{
			Write-Warning "Unable to create $FolderPath"
			Write-Output " "
			throw $_
			Write-Output " "
			$Result = $false
		}
	}Else{
		Write-Verbose "$FolderPath already exists..."
		$result = $false
	}

	Return $result
}