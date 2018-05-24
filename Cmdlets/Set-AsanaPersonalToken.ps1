function Set-AsanaPersonalToken
{
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [String] $Token
    )

    Begin
    {
    }
    Process
    {
        $GLOBAL:AsanaPersonalAccessToken = $Token
    }
    End
    {
    }
}