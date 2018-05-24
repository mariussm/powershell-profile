<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Get-AsanaProjectWithTasksAndPhase
{
    [CmdletBinding()]
    [Alias()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $ProjectNumber,

        [Parameter(Mandatory=$false,
                   Position=1)]
        [String] $Phase1Name = "Phase 1:",

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=2)]
        [Boolean] $IncludeSubTasks = $false
    )

    Begin
    {
    }
    Process
    {
        $project = Get-AsanaProject -Id $ProjectNumber -Verbose -IncludeSubTasks:$IncludeSubTasks
        
        $currentPhase = $Phase1Name
        $project.Tasks | foreach {
            if($_.Task.Name -like "*:") {
                $currentPhase = $_.Task.Name -replace ":$",""
            } 

            $_ | Add-Member -Force NoteProperty -Name Phase -Value $currentPhase
        }

        return $project
    }
    End
    {
    }
}
