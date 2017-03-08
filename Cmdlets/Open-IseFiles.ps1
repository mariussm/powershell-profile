Function Open-IseFiles {
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [System.IO.FileInfo] $FullName
    )

    Begin
    {
    }
    Process
    {
        if(!(Test-path $FullName)) {
            Write-Error "No such file: $FullName"
            return;
        }

        $psise.CurrentPowerShellTab.Files.Add($FullName)
    }
    End
    {
    }

}
