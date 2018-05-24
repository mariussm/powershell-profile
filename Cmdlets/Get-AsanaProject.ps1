function Get-AsanaProject
{
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [String] $Id,

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [Boolean] $IncludeTasks = $true,

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
        Write-Verbose "Getting asana project $id"

        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12, [Net.SecurityProtocolType]::Tls, [Net.SecurityProtocolType]::Ssl3

        $headers = Get-AsanaHttpHeaders
        
        $uri = Get-AsanaUri "/projects/$Id"
        $project = Invoke-RestMethod -Uri $uri -Method GET -Headers $headers -ErrorAction SilentlyContinue 
        
        
        
        $returnObject = @{
            Project = $project.data
        }

        if($IncludeTasks) {
            $returnObject["Tasks"] = Get-AsanaTasks -ProjectId $Id | Get-AsanaTask -IncludeSubTasks:$IncludeSubTasks
        }

        $returnObject["Users"] = @($returnObject["Tasks"] | foreach{$_.task ; $_.subtasks | foreach{$_.task}} | foreach{$_.assignee.id}; $project.data.members.id) | sort -Unique | foreach{Get-AsanaUser -Id $_} | Group -AsHashTable -Property id


        [PSCustomObject] $returnObject
    }
    End
    {
    }
}