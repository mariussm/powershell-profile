Function Copy-AttributeToClipboard {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=1)]
        $Object,

        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   Position=0)]
        [String] $Attribute
    )

    Begin
    {
    }
    Process
    {
        $Object | select -ExpandProperty $Attribute | clip
    }
    End
    {
    }

}
