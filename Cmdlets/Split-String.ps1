Function Split-String {
    [CmdletBinding()]
    [OutputType([string[]])]
    Param
    (
        # The input string object
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        [String] $InputObject,

        # Split delimiter
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=1)]
        [String] $Delimiter = "`n",

        # Do trimming or not
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=2)]
        [Boolean] $Trim = $true

    )

    Begin{}
    Process {
        if($Trim) {
            return $InputObject -split $Delimiter | foreach{$_.Trim()}
        } else {
            return $InputObject -split $Delimiter
        }
    } 
    End{}

}
