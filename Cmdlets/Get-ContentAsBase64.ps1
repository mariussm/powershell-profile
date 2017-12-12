function Get-ContentAsBase64
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([string])]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Path
    )

    Process
    {
        ConvertTo-Base64 -ByteArray ([IO.File]::ReadAllBytes((Resolve-Path $Path).Path))
    }
}