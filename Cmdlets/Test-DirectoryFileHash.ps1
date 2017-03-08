Function Test-DirectoryFileHash {
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   Position=0)]
        [String] $XMLPath = $null,

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

        $FilesFromXML = Import-Clixml -Path $XMLPath | group RelativePath -AsHashTable
        $FilesInFolder = @{}
        dir -Recurse $Path -File | Get-FileHash | Select Hash, @{Label="RelativePath"; Expression={$_.Path.Replace($path,"")}} | foreach{$FilesInFolder[$_.RelativePath] = $_} 
        
        $Errors = @()
        $FilesFromXML.Values | where{!$FilesInFolder.ContainsKey($_.RelativePath)} | foreach{
            $Errors += [PSCustomObject]@{File = $_.RelativePath; Error = "Missing"}            
        }
        
        $FilesInFolder.Values | Where{$FilesFromXML.ContainsKey($_.RelativePath)} | where{$FilesFromXML[$_.RelativePath].Hash -ne $_.Hash} | foreach{
            $Errors += [PSCustomObject]@{File = $_.RelativePath; Error = "File corrupt"}   
        }

        $Errors
    }
    End
    {
    }

}
