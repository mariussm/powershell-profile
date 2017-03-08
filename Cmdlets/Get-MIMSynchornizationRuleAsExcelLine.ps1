<#
.Synopsis
   Returns excel line for deployment excel file
.DESCRIPTION
   Returns excel line for deployment excel file
.EXAMPLE
   Get-FIMObjectByXPath /SynchronizationRule | Get-MIMSynchornizationRuleAsExcelLine
#>
function Get-MIMSynchornizationRuleAsExcelLine
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $SynchronizationRule
    )

    Begin
    {
    }
    Process
    {
        "{0}`t{1}`t{2}`t{3}`t{4}`t{5}`t{6}`t{7}`t{8}`t{9}`t{10}`t{11}" -f 
            #(Get-FIMObjectByXPath ("/ma-data[ObjectID=""{0}""]" -f $SynchronizationRule.ManagementAgentID -replace "urn:uuid:","")).DisplayName,
            $SynchronizationRule.DisplayName,
            $SynchronizationRule.FlowType,
            $SynchronizationRule.ConnectedObjectType,
            $SynchronizationRule.ILMObjectType,
            ($SynchronizationRule.ConnectedSystemScope -join ";;;"),
            $SynchronizationRule.CreateConnectedSystemObject,
            $SynchronizationRule.CreateILMObject,
            $SynchronizationRule.DisconnectConnectedSystemObject,
            ($SynchronizationRule.RelationshipCriteria -join ";;;"),
            ($SynchronizationRule.PersistentFlow -join ";;;"),
            ($SynchronizationRule.InitialFlow -join ";;;"),
            ($SynchronizationRule.ExistenceTest -join ";;;")

    }
    End
    {
    }
}
