<#
.Synopsis
   Converts immutableID in Office 365 to GUID
.DESCRIPTION
   Converts immutableID in Office 365 to GUID
.EXAMPLE
   Get-MsolUser -UserPrincipalName marius@goodworkaround.com | Select -ExpandProperty ImmutableID | ConvertFrom-ImmutableID
#>
function ConvertFrom-ImmutableID
{
    [CmdletBinding()]
    [OutputType([GUID])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $ImmutableID
    )

    Process 
    {
        return [guid]([system.convert]::frombase64string($ImmutableID) )
    }
}