<#
.Synopsis
   Returns all MPRs that triggers an action workflow
.DESCRIPTION
   Returns all MPRs that triggers an action workflow
.EXAMPLE
   Get-FimWorkflow *accountname* | Get-FIMManagementPolicyRuleByActionWorkflowDefinition
#>
function Get-FIMManagementPolicyRuleByActionWorkflowDefinition
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $FIMWorkflow
    )

    Begin
    {
    }
    Process
    {
        return (Export-FimConfig -CustomConfig ("/ManagementPolicyRule[ActionWorkflowDefinition='$($FIMWorkflow.ObjectID.Replace('urn:uuid:',''))']") -OnlyBaseResources -Uri "http://localhost:5725/" | Convert-FimExportToPSObject)
    }
    End
    {
    }
}