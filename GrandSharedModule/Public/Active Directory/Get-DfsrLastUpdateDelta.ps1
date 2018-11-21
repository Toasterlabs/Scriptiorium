Function Get-DfsrLastUpdateDelta {
        param ([string]$ComputerName)
        $ErrorActionPreference = "Stop"

        If (!$ComputerName){Throw "You must supply a value for ComputerName."}

        $LastUpdateTime = Get-DfsrLastUpdateTime -ComputerName $ComputerName
        $TimeDelta = (Get-Date) - $LastUpdateTime
    
        Return $TimeDelta 
    }