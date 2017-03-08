<#
.Synopsis
   Converts GUID in AD to ImmutableID
.DESCRIPTION
   Converts GUID in AD to ImmutableID
.EXAMPLE
   GetADUser | Select -ExpandProperty ImmutableID | ConvertFrom-ImmutableID
#>
function ConvertTo-ImmutableID
{
    [CmdletBinding()]
    [OutputType([GUID])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [GUID] $ObjectGUID
    )

    Process 
    {
        return [system.convert]::ToBase64String($ObjectGUID.ToByteArray())
    }
}
