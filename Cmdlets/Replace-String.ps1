Function Replace-String {
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=4)]
        $InputObject,

        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   Position=0)]
        [String] $Pattern = $null,

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=1)]
        [String] $Replacement = "",

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=2)]
        [String] $Property = $null,

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=3)]
        [Boolean] $CaseSensitive = $false
    )

    Begin
    {
    }
    Process
    {
        if($Property) {
            $OutputObject = $InputObject | select -Property * 
            if($CaseSensitive) {
                $OutputObject.$Property = $OutputObject.$Property -creplace $Pattern, $Replacement
            } else {
                $OutputObject.$Property = $OutputObject.$Property -replace $Pattern, $Replacement
            }
            $OutputObject
        } else {
            if($CaseSensitive) {
                $InputObject -creplace $Pattern, $Replacement
            } else {
                $InputObject -ireplace $Pattern, $Replacement
            }
        }
    }
    End
    {
    }

}
