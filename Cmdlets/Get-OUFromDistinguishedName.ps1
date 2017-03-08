Function Get-OUFromDistinguishedName {
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
        $DN -split "[^\\],", 2 | select -last 1
    }
    End
    {
    }

}
