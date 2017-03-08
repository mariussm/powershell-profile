Function Trim-String {
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=1)]
        $InputObject,

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=0)]
        [String] $Property = $null
    )

    Begin
    {
    }
    Process
    {
        if($Property) {
            $OutputObject = $InputObject | select -Property * 
            $OutputObject.$Property = $OutputObject.$Property.Trim()
            $OutputObject
        } else {
            $InputObject.Trim()
        }
    }
    End
    {
    }

}
