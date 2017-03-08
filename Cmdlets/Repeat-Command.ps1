Function Repeat-Command {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=2)]
        [scriptblock] $ScriptBlock,

        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   Position=1)]
        [Int] $Sleep,

        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   Position=0)]
        [Int] $Times
    )

    Begin
    {
    }
    Process
    {
        1..$times | foreach {
            $ScriptBlock.InvokeReturnAsIs()
            if($_ -ne $times) {
                 sleep -Milliseconds $Sleep
            }
        }
    }
    End
    {
    }

}
