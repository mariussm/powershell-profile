Function Add-CodeSignatureToCurrentISEFile {
    [CmdletBinding()]
    [Alias()]
    Param
    (
    )

    Begin
    {
    }
    Process
    {
        if($psise.CurrentFile)
        {
            Add-CodeSignature -Files $psise.CurrentFile.FullPath
        }
    }
    End
    {
    }

}
