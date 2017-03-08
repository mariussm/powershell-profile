Function Start-WaitUntil {
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=2)]
        $Object,

        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   Position=0)]
        [System.Management.Automation.ScriptBlock]
        $CheckScript,

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=1)]
        [System.Management.Automation.ScriptBlock]
        $DoneScript = {}
    )

    Begin
    {
        $BreakDone = $false
    }
    Process
    {
        if($BreakDone) {
            break
        } elseif ($Object | where -FilterScript $CheckScript) {
            $DoneScript.Invoke()
            $BreakDone = $true
            break
        }
    }
    End
    {
        
    }

}
