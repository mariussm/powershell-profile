function Get-AsanaPersonalToken
{
    [CmdletBinding()]
    [Alias()]
    Param
    ()

    Begin
    {
    }
    Process
    {
        if($GLOBAL:AsanaPersonalAccessToken) {
            return $GLOBAL:AsanaPersonalAccessToken
        } else {
            throw "No personal access token found"
        }
    }
    End
    {
    }
}