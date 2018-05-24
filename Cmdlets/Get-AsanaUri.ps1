
function Get-AsanaUri
{
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [String] $Endpoint
    )

    Begin
    {
    }
    Process
    {
        if(!$Endpoint.StartsWith("/")) {
            $Endpoint = "/$Endpoint"
        }
        return "https://app.asana.com/api/1.0" + $Endpoint
    }
    End
    {
    }
}