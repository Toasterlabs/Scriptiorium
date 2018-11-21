#region Check if Variable is a path or a file

	# if only a path was specified (i.e. file name not included at the end of the
	# directory path), then auto generate a file name in the format of YYYYMMDDhhmmss.csv
	# and append to the directory path
	If (!([System.IO.Path]::HasExtension($OutputFile)))
	{
		If ($OutputFile.substring($OutputFile.length - 1) -eq "\")
		{
			$OutputFile += "{0}.csv" -f (Get-Date -uformat %Y%m%d%H%M%S).ToString()
		}
		Else
		{
			$OutputFile += "\{0}.csv" -f (Get-Date -uformat %Y%m%d%H%M%S).ToString()
		}
	}

#endregion

#region Beep
[console]::Beep(500,300)
#endregion