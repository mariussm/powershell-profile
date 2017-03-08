Function Get-ContentAsString {
    [CmdletBinding()]
    [Alias()]
    [OutputType([string])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Path
    )

    Begin
    {
    }
    Process
    {
        return [IO.File]::ReadAllText((dir ($Path)).Fullname)
    }
    End
    {
    }

}
