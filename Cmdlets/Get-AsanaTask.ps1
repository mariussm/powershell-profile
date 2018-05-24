function Get-AsanaTask
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
        [Boolean] $IncludeSubTasks = $false
    )

    Begin
    {
    }
    Process
    {
        Write-Verbose "Getting asana task $Id"
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12, [Net.SecurityProtocolType]::Tls, [Net.SecurityProtocolType]::Ssl3

        $headers = Get-AsanaHttpHeaders
        
        $uri = Get-AsanaUri "/tasks/$Id"
        $task = Invoke-RestMethod -Uri $uri -Method GET -Headers $headers -ErrorAction SilentlyContinue 
        $subtasks = @()
        
        if($IncludeSubTasks) {
            $subtaskuri = Get-AsanaUri "/tasks/$Id/subtasks"
            $subtasksresponse = Invoke-RestMethod -Uri $subtaskuri -Method GET -Headers $headers -ErrorAction SilentlyContinue 
            if($subtasksresponse.data) {
                $subtasks = $subtasksresponse.data | Get-AsanaTask -IncludeSubTasks:$IncludeSubTasks
            }
        }
        
        $ReturnObject = @{
            Task = $task.data
            CustomFieldsById = @{}
            CustomFieldsByName = @{}
            Subtasks = $subtasks
        }

        $task.data.custom_fields | foreach {
            if($_.id) {
                $ReturnObject.CustomFieldsById[$_.id] = $_
            }

            if($_.name) {
                $ReturnObject.CustomFieldsByName[$_.name] = $_
            }
        }

        [PSCustomObject] $ReturnObject
    }
    End
    {
    }
}