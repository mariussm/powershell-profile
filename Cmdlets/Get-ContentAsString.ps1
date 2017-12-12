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
        $Path,

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [System.Text.Encoding] $Encoding = [System.Text.Encoding]::Default
    )

    Begin
    {
    }
    Process
    {
        return [IO.File]::ReadAllText((dir ($Path)).FullName, $Encoding)
    }
    End
    {
    }

}
