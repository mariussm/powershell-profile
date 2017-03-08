Function ConvertTo-MB {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True,Position=0,ValueFromPipeline=$true)]
        [string]$ByteQuantifiedSize
    )

    Begin {}
    Process {
        return (ConvertTo-Bytes $ByteQuantifiedSize) / 1024 / 1024
    }
    End{}

}
