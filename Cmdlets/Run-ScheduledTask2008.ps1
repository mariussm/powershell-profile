Function Run-ScheduledTask2008 {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True,Position=0)]
        [string]$Name
    )

    $ret = schtasks.exe /Run /TN "$Name"

    return ($ret -like "SUCCESS:*") -eq $true

}
