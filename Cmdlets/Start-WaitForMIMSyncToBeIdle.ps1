<#
.Synopsis
   Waits until no MAs are active (or has been within the last 30 seconds)
.DESCRIPTION
   Waits until no MAs are active (or has been within the last 30 seconds)
.EXAMPLE
   Start-WaitForMIMSyncToBeIdle
#>
function Start-WaitForMIMSyncToBeIdle
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=0)]
        [int] $Wait = 30
    )

    Begin
    {
        $wmi = Get-WmiObject -Namespace root/MicrosoftIdentityIntegrationServer -Class MIIS_ManagementAgent
    }
    Process
    {
        if($wmi) {
            $sleepTime = 0
            do {
                $inProgress = $wmi | where {
                    $value = $_.RunEndTime().ReturnValue
                    if($value -eq "in-progress"){return $true}
                    if($value -ne "") {
                        (Get-Date ($value)) -gt (Get-Date).AddSeconds(0 - $wait)
                    }
                }

                sleep -Seconds $sleepTime
                $sleepTime = 10
            } while($inProgress)
        }
    }
    End
    {
    }
}