Function ConvertTo-Mbps {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True,Position=0,ValueFromPipeline=$true)]
        [string]$MegaBytesPerMinute
    )

    Begin{}
    Process{
        return ($MegaBytesPerMinute / 60 * 8)
    }
    End{}

}
