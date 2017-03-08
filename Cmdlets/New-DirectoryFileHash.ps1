Function New-DirectoryFileHash {
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   Position=0)]
        [String] $OutputPath = $null,

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=1)]
        [String] $Path = "."
    )

    Begin
    {
    }
    Process
    {
        if(!(Test-path $path -PathType Container)) {
            Write-Error -ErrorAction Stop "Not a folder: $path"
        }

        $Path = (Get-Item $Path).FullName


        dir -Recurse $Path -File | Get-FileHash | Select Hash, @{Label="RelativePath"; Expression={$_.Path.Replace($path,"")}} | Export-Clixml -Path $OutputPath
        
    }
    End
    {
    }

}
