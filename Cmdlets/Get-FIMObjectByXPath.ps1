<#
.Synopsis
   Returns objects matching xpath
.DESCRIPTION
   Returns objects matching xpath
.EXAMPLE
   Get-FIMObjectByXPath "/testUser"
#>
function Get-FIMObjectByXPath
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $XPath
    )

    Begin
    {
    }
    Process
    {
        return (Export-FimConfig -CustomConfig $XPath -OnlyBaseResources -Uri "http://localhost:5725/" | Convert-FimExportToPSObject)
    }
    End
    {
    }
}