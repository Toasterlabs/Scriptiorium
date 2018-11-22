Function New-RandomFile{
<#
.Synopsis
  Creates a new file with random content

.Description
  Used for uploads and download testing. the random content "should" fool WAN optimizers

.PARAMETER  filename
  Filename for file

.PARAMETER filesize
  Size the file should be

.OUTPUTS
  Text file

#>
    Param(
        [Parameter(Mandatory=$True)]
        $filename,
        [Parameter(Mandatory=$True)]
        $filesize
        )

    $path = "$env:temp\$filename"

    if(Test-Path $path){
        Write-Verbose "Test file $path already exists."
    } Else {

        # Randomizing data
        $Chunk = { [guid]::NewGuid().Guid + [guid]::NewGuid().Guid + [guid]::NewGuid().Guid + [guid]::NewGuid().Guid +
                   [guid]::NewGuid().Guid + [guid]::NewGuid().Guid + [guid]::NewGuid().Guid + [guid]::NewGuid().Guid +
                   [guid]::NewGuid().Guid + [guid]::NewGuid().Guid + [guid]::NewGuid().Guid + [guid]::NewGuid().Guid +
                   [guid]::NewGuid().Guid + [guid]::NewGuid().Guid + [guid]::NewGuid().Guid + [guid]::NewGuid().Guid +
                   [guid]::NewGuid().Guid + [guid]::NewGuid().Guid + [guid]::NewGuid().Guid + [guid]::NewGuid().Guid +
                   [guid]::NewGuid().Guid + [guid]::NewGuid().Guid + [guid]::NewGuid().Guid + [guid]::NewGuid().Guid +
                   [guid]::NewGuid().Guid + [guid]::NewGuid().Guid + [guid]::NewGuid().Guid + [guid]::NewGuid().Guid +
                   [guid]::NewGuid().Guid + [guid]::NewGuid().Guid + [guid]::NewGuid().Guid + [guid]::NewGuid().Guid -Replace "-" }

        # Going until we have the proper filesize
        $Chunks = [math]::Ceiling($FileSize/1kb)
        $ChunkString = $Chunk.Invoke()

        # Writing file
        [io.file]::WriteAllText("$Path","$(-Join (1..($Chunks)).foreach({ $ChunkString }))")

    }

    # returning file path
    return $path
}