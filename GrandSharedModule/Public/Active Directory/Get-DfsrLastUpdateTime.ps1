Function Get-DfsrLastUpdateTime {
        param ([string]$ComputerName)
        $ErrorActionPreference = "Stop"

        If (!$ComputerName){Throw "You must supply a value for ComputerName."}

        $DfsrWmiObj = Get-WmiObject -Namespace "root\microsoftdfs" -Class dfsrVolumeConfig -ComputerName $ComputerName
        If ($DfsrWmiObj.LastChangeTime.Count -le 1){
            [datetime]$LastChangeTime = [System.Management.ManagementDateTimeconverter]::ToDateTime($DfsrWmiObj.LastChangeTime)
        }
        Else {
            $OldestChangeTime = ($DfsrWmiObj.LastChangeTime | Measure-Object -Minimum).Minimum
            [datetime]$LastChangeTime = [System.Management.ManagementDateTimeconverter]::ToDateTime($OldestChangeTime)
        }

        Return $LastChangeTime
    }