function Start-RdpFiles
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        [Parameter(Mandatory=$false,Position=0)]
        [String] $Folder = (join-path ([System.Environment]::GetFolderPath("UserProfile")) "Downloads")
    )

    Begin
    {
    }
    Process
    {
        dir $Folder -Recurse | ? Extension -eq ".rdp" | Out-GridView -OutputMode Multiple | ii
    }
    End
    {
    }
}