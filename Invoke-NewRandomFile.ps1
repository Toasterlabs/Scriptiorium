Function Invoke-NewRandomFile{
<#
    <#
        .SYNOPSIS
          Creates a file with random content
        .DESCRIPTION
          Creates a file of a specifized size in the TEMP folder, and fills it with random jibberish. Once created, it returns the filepath
        .PARAMETER FILENAME
            name for the file
        .PARAMETER FILESIZE
            Size of the file to be created
        .INPUTS
          none
        .OUTPUTS
          none
        .NOTES
          Version:        1.0
          Author:         Marc Dekeyser
          Creation Date:  Juli 7th, 2018
          Purpose/Change: Just having some fun
  
        .EXAMPLE
          Invoke-NewRandomFile -filename randomfile.txt -filesize 45MB
    #>

#>
    Param(
        [Parameter(Mandatory=$True)]
        $filename,
        [Parameter(Mandatory=$True)]
        $filesize
        )

    $path = "$env:temp\$filename"

    if(Test-Path $path){
        Write-Verbose "Random file $path already exists."
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
