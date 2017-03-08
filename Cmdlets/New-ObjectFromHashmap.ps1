Function New-ObjectFromHashmap {
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $Hashmap
    )

    Begin
    {
    }
    Process
    {
        New-Object -TypeName PSCustomObject -Property $Hashmap
    }
    End
    {
    }

}
