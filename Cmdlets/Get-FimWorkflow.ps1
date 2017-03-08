<#
.Synopsis
   Returns all Fim workflows matching pattern
.DESCRIPTION
   Returns all Fim workflows matching pattern
.EXAMPLE
   Get-FimWorkflow *accountname*
#>
function Get-FimWorkflow
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   Position=0)]
        $Name
    )

    Begin
    {
    }
    Process
    {
        return (Export-FimConfig -CustomConfig ("/WorkflowDefinition[DisplayName='{0}']" -f $Name) -OnlyBaseResources -Uri "http://localhost:5725/" | Convert-FimExportToPSObject)
    }
    End
    {
    }
}