<#
	Wrapper module
#>

# Get public and private function definition files.
$Public = @( Get-ChildItem -Path $PSScriptRoot\Public\*.ps1 -ErrorAction SilentlyContinue -Recurse )
$Private = @( Get-ChildItem -Path $PSScriptRoot\Private\*.ps1 -ErrorAction SilentlyContinue -Recurse )
$Assembly = @( Get-ChildItem -Path $PSScriptRoot\Lib\*.dll -ErrorAction SilentlyContinue -Recurse )

# Importing Public and Private scripts
Foreach ($import in @($Public + $Private)) {
    Try {
        . $import.fullname
    } Catch {
        Write-Error -Message "Failed to import function $($import.fullname): $_"
    }
}

# Importing Library assemblies
Foreach ($import in @($Assembly)) {
    Try {
        Add-Type -Path $import.fullname
    } Catch {
        Write-Error -Message "Failed to import DLL $($import.fullname): $_"
    }
}

# Define the cmds of all functions to export
$FunctionsToExport = 'Invoke-WriteLog', 'Invoke-ColorOutput'

# Export functions
Export-ModuleMember -Function $FunctionsToExport

# manifest updating
[string] $ManifestFile = '{0}.psd1' -f (Get-Item $PSCommandPath).BaseName;
$ManifestPathAndFile = Join-Path -Path $PSScriptRoot -ChildPath $ManifestFile;
if ( Test-Path -Path $ManifestPathAndFile) {
    $Manifest = (Get-Content -raw $ManifestPathAndFile) | iex;
    foreach ( $ScriptToProcess in $Manifest.ScriptsToProcess) {
        $ModuleToRemove = (Get-Item (Join-Path -Path $PSScriptRoot -ChildPath $ScriptToProcess)).BaseName;
        if (Get-Module $ModuleToRemove) {
            Remove-Module $ModuleToRemove;
        }
    }
}