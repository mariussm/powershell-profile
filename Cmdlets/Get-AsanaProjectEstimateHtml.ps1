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
function Get-AsanaProjectEstimateHtml
{
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        [String] $ProjectNumber,

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=1)]
        [Boolean] $ShowEstimatedHours = $true,

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=2)]
        [Boolean] $ShowEstimatedPercentOfTotalProject = $true,

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=3)]
        [String] $EstimatedHoursText = "Estimated hours for this phase",

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=4)]
        [String] $EstimatedPercentOfTotalProjectText = "Estimated percent of total project"
    )

    Begin
    {
    }
    Process
    {
        $AsanaToken = ([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR((ConvertTo-SecureString ([IO.File]::ReadAllText((Resolve-Path "$($env:APPDATA)\asana.txt"))).Trim()))))
        Set-AsanaPersonalToken -Token $AsanaToken

        $Project = Get-AsanaProjectWithTasksAndPhase -ProjectNumber $ProjectNumber

        $ShowEstimate = $true
        $EstimateTotal = $Project.Tasks | foreach{$_.CustomFieldsByName["Estimated Hours"]}  | measure -Sum -Property number_value
        $Project.Tasks | Foreach {
            if($_.Task.Name -like "*:") {
                $Name = $_.Task.name -replace ":$",""
                "<h1>$Name</h1>"
                if($_.Task.notes) {
                    $notes = $_.Task.notes -replace "`n","</p><p class='notes'>"
                    "<p class='notes'>$notes</p>"
                }
        
                $Estimate = $Project.Tasks | ? Phase -eq $Name | foreach{$_.CustomFieldsByName["Estimated Hours"]}  | measure -Sum -Property number_value
                $Percent = [Math]::Ceiling(100 * $Estimate.Sum / $EstimateTotal.Sum)
                if($ShowEstimatedHours) {
                    "<p class='estimate'>$($EstimatedHoursText): $($Estimate.Sum)</p>"
                }

                if($ShowEstimatedPercentOfTotalProject) {
                    "<p class='estimate'>$($EstimatedPercentOfTotalProjectText): $($Percent)</p>"
                }
                # "<p class='tasklist'>Tasks in this phase:</p>"

                $Project.Tasks | ? Phase -eq $Name | ?{$_.task.name -notlike "*:"} | foreach -Begin {"<ul>"} -End {"</ul>"} -Process {"<li>$($_.task.name)</li>"}
            } else {
                # "<p class='task'>$($_.Task.Name)</p>"
            }
        } | Get-StringsAsHtml -Style "* {font-family: arial} h1 {font-size: 20px; margin: 10px 0 2px 0} p.estimate {margin: 2px 0} p.tasklist {margin: 15px 0 2px 0; font-weight: bold;} p.task {margin: 2px 0 2px 10px}" | set-content "$($env:temp)\1.html"
        ii "$($env:temp)\1.html"

    }
    End
    {
    }
}