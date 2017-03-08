<#
.Synopsis
   Copies the input workflow definition to new workflow object
.DESCRIPTION
   Copies the input workflow definition to new workflow object
.EXAMPLE
   Get-FIMWorkflow *accountname* | New-FIMWorkflowCopy
#>
function New-FIMWorkflowCopy
{
    [CmdletBinding()]
    [OutputType([int])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $Source
    )

    Begin
    {
    }
    Process
    {
        $attributes = @(
            New-FimImportChange -Operation None -AttributeName 'DisplayName' -AttributeValue "___COPY - $($Source.DisplayName)"
            New-FimImportChange -Operation None -AttributeName 'RunOnPolicyUpdate' -AttributeValue $Source.RunOnPolicyUpdate
            New-FimImportChange -Operation None -AttributeName 'RequestPhase' -AttributeValue $Source.RequestPhase
            New-FimImportChange -Operation None -AttributeName 'XOML' -AttributeValue $Source.XOML
        )

        New-FimImportObject -ObjectType "WorkflowDefinition" -State Create -Changes $attributes -ApplyNow:$true -PassThru -SkipDuplicateCheck:$true

    }
    End
    {
    }
}