function Join-String
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([String])]
    Param
    (
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=0)]
        [String] $Delimiter = ", ",

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=1)]
        [String] $Qualifier,

        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=2)]
        [String] $String
    )

    Begin
    {
        $list = New-Object System.Collections.ArrayList
    }
    Process
    {
        $list.Add($Qualifier + $String + $Qualifier) | Out-Null
    }
    End
    {
        return ($list -join $Delimiter)
    }
}