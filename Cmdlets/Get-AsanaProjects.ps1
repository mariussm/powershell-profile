function Get-AsanaProjects
{
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [ScriptBlock] $Filter = {$_.name -like "#* - *"}
    )

    Begin
    {
    }
    Process
    {
        Write-Verbose "Getting all asana projects"

        $headers = Get-AsanaHttpHeaders
        $uri = Get-AsanaUri "/projects?archived=false" # ?limit=10&workspace=70246419796023" #/114447556184299?opt_fields=name"

        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12, [Net.SecurityProtocolType]::Tls, [Net.SecurityProtocolType]::Ssl3
        $result = Invoke-RestMethod -Uri $uri -Method GET -Headers $headers -ErrorAction SilentlyContinue 
        $result.data | Where $Filter 
    }
    End
    {
    }
}