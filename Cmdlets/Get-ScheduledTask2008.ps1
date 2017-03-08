Function Get-ScheduledTask2008 {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True,Position=0)]
        [string]$Name
    )

    return schtasks.exe /Query /V /FO CSV /TN "$Name" | ConvertFrom-Csv

}
