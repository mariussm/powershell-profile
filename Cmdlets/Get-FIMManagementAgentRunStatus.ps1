<#
.Synopsis
   Returns fim management agent run status for all, one or some MAs
.DESCRIPTION
   Returns fim management agent run status for all, one or some MAs
.EXAMPLE
   Get-FIMManagementAgentRunStatus "SP - *"
#>
function Get-FIMManagementAgentRunStatus
{
    [CmdletBinding()]
    [OutputType([object])]
    Param
    (
        # The Management Agent name
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=0)]
        [string] $MA
    )

    Begin
    {
    }
    Process
    {
        if($MA) 
        {
            $MAs = Get-FIMManagementAgent -MA $MA
        }   
        else 
        {
            $MAs = Get-FIMManagementAgent -MA *
        }
        
        return ($MAs | foreach{New-Object -TypeName PSObject -Property @{ManagementAgent=$_.Name;RunStatus=$_.RunStatus().ReturnValue}})
    }
    End
    {   
    }
}