Function Start-WaitForCmdletOutputChange {
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   Position=0)]
        [System.Management.Automation.ScriptBlock] $Script,

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=1)]
        [int] $SleepSec = 1
    )

    Begin
    {
    }
    Process
    {
        $Orig = $Script.Invoke() | out-string
        Write-Verbose "Original output: $Orig"
        do {
            $Result = $Script.Invoke() | out-string
            Write-Verbose "Result: $Result"
            Sleep -Seconds $SleepSec
        } until ($Result -ne $Orig)
    }
    End
    {
    }

}
