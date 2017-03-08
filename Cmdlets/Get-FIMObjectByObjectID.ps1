<#
.Synopsis
   Returns object with object id
.DESCRIPTION
   Returns object with object id
.EXAMPLE
   Get-FIMObjectByObjectID "0a0b2dsa-ccccc-cccc-cccccccccccc"
#>
function Get-FIMObjectByObjectID
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        [String] $ObjectID
    )

    Begin
    {
    }
    Process
    {
        $ObjectID = $ObjectID.Replace("urn:uuid:","")
        return (Export-FimConfig -CustomConfig ("/*[ObjectID='$($ObjectID)']") -OnlyBaseResources -Uri "http://localhost:5725/" | Convert-FimExportToPSObject)
    }
    End
    {
    }
}