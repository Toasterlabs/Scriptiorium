Function Get-DfsrGuid {
        param ([string]$ComputerName)
        $ErrorActionPreference = "Stop"

        If (!$ComputerName){Throw "You must supply a value for ComputerName."}

        $DfsrWmiObj = Get-WmiObject -Namespace "root\microsoftdfs" -Class dfsrVolumeConfig -ComputerName $ComputerName

        Return $DfsrWmiObj.VolumeGUID
    }