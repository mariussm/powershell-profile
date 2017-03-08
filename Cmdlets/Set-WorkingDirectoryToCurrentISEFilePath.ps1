Function Set-WorkingDirectoryToCurrentISEFilePath {
    [CmdletBinding()]
    [Alias("cdise")]
    Param
    ()

    Process
    {
        if($psise.CurrentFile.FullPath) {
            cd (split-path -Parent -Path $psise.CurrentFile.FullPath)
        }
    }
    

}
