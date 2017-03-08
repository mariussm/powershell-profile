Function Load-Assembly {
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Assembly
    )

    Begin
    {
    }
    Process
    {
        return [System.Reflection.Assembly]::LoadWithPartialName($Assembly)
    }
    End
    {
    }

}
