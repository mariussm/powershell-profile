<#
.Synopsis
   Converts a filetime to datetime. Can be used on lastLogonTimestamp in AD.
.DESCRIPTION
   Converts a filetime to datetime. Can be used on lastLogonTimestamp in AD.
.EXAMPLE
   Get-ADUser masol -property lastLogonTimestamp | Select-Object -ExpandProperty lastLogonTimestamp | ConvertFrom-LastLogonTimestamp
.EXAMPLE
   ConvertFrom-LastLogonTimestamp 129948127853609000
#>
function ConvertFrom-LastLogonTimestamp
{
    [CmdletBinding()]
    [OutputType([datetime])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $LastLogonTimestamp
    )

    return [datetime]::FromFileTime($LastLogonTimestamp)
}