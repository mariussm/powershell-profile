Function ConvertTo-Bytes {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True,Position=0,ValueFromPipeline=$true)]
        [string]$ByteQuantifiedSize
    )

    return [long] ([string] $ByteQuantifiedSize).Split("(")[1].Split(" ")[0].Replace(",","")

}
