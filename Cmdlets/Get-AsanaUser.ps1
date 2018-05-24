function Get-AsanaUser
{
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [String] $Id
    )

    Begin
    {
    }
    Process
    {
        Write-Verbose "Getting asana user $Id"

        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12, [Net.SecurityProtocolType]::Tls, [Net.SecurityProtocolType]::Ssl3

        $headers = Get-AsanaHttpHeaders
        
        $uri = Get-AsanaUri "/users/$Id"
        $user = Invoke-RestMethod -Uri $uri -Method GET -Headers $headers -ErrorAction SilentlyContinue 
        $data = $user.data

        return $data 
    }
    End
    {
    }
}