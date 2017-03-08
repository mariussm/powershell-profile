Function Get-DesktopPath {
    [CmdletBinding()]
    [OutputType([string])]
    Param
    (
    )

    Begin
    {
    }
    Process
    {
        return [Environment]::GetFolderPath("Desktop")
    }
    End
    {
    }

}
