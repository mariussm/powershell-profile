Function Get-ContainerNameFromDistinguishedName {
    [CmdletBinding()]
    [OutputType([string])]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $DN
    )

    Begin
    {
    }
    Process
    {
        $DN -split "[^\\],", 2 -split "=" | select -index 1
    }
    End
    {
    }

}
