function Get-AsanaTasks
{
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [String] $ProjectId
    )

    Begin
    {
    }
    Process
    {
        Write-Verbose "Getting asana tasks for project $ProjectId"

        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12, [Net.SecurityProtocolType]::Tls, [Net.SecurityProtocolType]::Ssl3

        $headers = Get-AsanaHttpHeaders
        
        $uri = Get-AsanaUri "/projects/$Id/tasks?limit=100"
        $tasks = Invoke-RestMethod -Uri $uri -Method GET -Headers $headers -ErrorAction SilentlyContinue 
        $data = $tasks.data
        while($tasks.next_page) {
            $tasks = Invoke-RestMethod -Uri $tasks.next_page.uri -Method GET -Headers $headers -ErrorAction SilentlyContinue 
            $data += $tasks.data
        }

        return $data 
    }
    End
    {
    }
}