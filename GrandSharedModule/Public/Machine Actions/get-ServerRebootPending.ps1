Function Get-ServerRebootPending {
param(
[Parameter(Mandatory=$true)][string]$Machine_Name
)

    $PendingFileReboot = $false #Pending File Rename operations 
    $PendingAutoUpdateReboot = $false
    $PendingCBSReboot = $false #Component-Based Servicing Reboot 
    $PendingSCCMReboot = $false
    $ServerPendingReboot = $false

    #Pending File Rename operations 
    Function Get-PendingFileReboot {

        $PendingFileKeyPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\"
        $file = Get-ItemProperty -Path $PendingFileKeyPath -Name PendingFileRenameOperations
        if($file)
        {
            return $true
        }
        return $false
    }

    Function Get-PendingAutoUpdateReboot {

        if(Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired")
        {
            return $true
        }
        return $false
    }

    Function Get-PendingCBSReboot {

        if(Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending")
        {
            return $true
        }
        return $false
    }

    Function Get-PendingSCCMReboot {

        $SCCMReboot = Invoke-CimMethod -Namespace 'Root\ccm\clientSDK' -ClassName 'CCM_ClientUtilities' -Name 'DetermineIfRebootPending'

        if($SCCMReboot)
        {
            If($SCCMReboot.RebootPending -or $SCCMReboot.IsHardRebootPending)
            {
                return $true
            }
        }
        return $false
    }

    Function Execute-ScriptBlock{
    param(
    [Parameter(Mandatory=$true)][string]$Machine_Name,
    [Parameter(Mandatory=$true)][scriptblock]$Script_Block,
    [Parameter(Mandatory=$true)][string]$Script_Block_Name
    )

        $oldErrorAction = $ErrorActionPreference
        $ErrorActionPreference = "Stop"
        $returnValue = $false
        try 
        {
            $returnValue = Invoke-Command -ComputerName $Machine_Name -ScriptBlock $Script_Block
        }
        catch 
        {
            
            $Script:iErrorExcluded++
        }
        finally 
        {
            $ErrorActionPreference = $oldErrorAction
        }
        return $returnValue
    }

    Function Execute-LocalMethods {
    param(
    [Parameter(Mandatory=$true)][string]$Machine_Name,
    [Parameter(Mandatory=$true)][ScriptBlock]$Script_Block,
    [Parameter(Mandatory=$true)][string]$Script_Block_Name
    )
        $oldErrorAction = $ErrorActionPreference
        $ErrorActionPreference = "Stop"
        $returnValue = $false
        
        try 
        {
            $returnValue = & $Script_Block
        }
        catch 
        {
        
            $Script:iErrorExcluded++
        }
        finally 
        {
            $ErrorActionPreference = $oldErrorAction
        }
        return $returnValue
    }

    if($Machine_Name -eq $env:COMPUTERNAME)
    {
        
        $PendingFileReboot = Execute-LocalMethods -Machine_Name $Machine_Name -Script_Block ${Function:Get-PendingFileReboot} -Script_Block_Name "Get-PendingFileReboot"
        $PendingAutoUpdateReboot = Execute-LocalMethods -Machine_Name $Machine_Name -Script_Block ${Function:Get-PendingAutoUpdateReboot} -Script_Block_Name "Get-PendingAutoUpdateReboot"
        $PendingCBSReboot = Execute-LocalMethods -Machine_Name $Machine_Name -Script_Block ${Function:Get-PendingCBSReboot} -Script_Block_Name "Get-PendingCBSReboot"
        $PendingSCCMReboot = Execute-LocalMethods -Machine_Name $Machine_Name -Script_Block ${Function:Get-PendingSCCMReboot} -Script_Block_Name "Get-PendingSCCMReboot"
    }
    else 
    {
        
        $PendingFileReboot = Execute-ScriptBlock -Machine_Name $Machine_Name -Script_Block ${Function:Get-PendingFileReboot} -Script_Block_Name "Get-PendingFileReboot"
        $PendingAutoUpdateReboot = Execute-ScriptBlock -Machine_Name $Machine_Name -Script_Block ${Function:Get-PendingAutoUpdateReboot} -Script_Block_Name "Get-PendingAutoUpdateReboot"
        $PendingCBSReboot = Execute-ScriptBlock -Machine_Name $Machine_Name -Script_Block ${Function:Get-PendingCBSReboot} -Script_Block_Name "Get-PendingCBSReboot"
        $PendingSCCMReboot = Execute-ScriptBlock -Machine_Name $Machine_Name -Script_Block ${Function:Get-PendingSCCMReboot} -Script_Block_Name "Get-PendingSCCMReboot"
    }

    
    if($PendingFileReboot -or $PendingAutoUpdateReboot -or $PendingCBSReboot -or $PendingSCCMReboot)
    {
        $ServerPendingReboot = $true
    }

    
    return $ServerPendingReboot
}