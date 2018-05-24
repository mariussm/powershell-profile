function Get-AsanaHttpHeaders
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
        return @{
            "Authorization" = "Bearer {0}" -f (Get-AsanaPersonalToken)
            #"Authorization" = "Bearer {0}" -f [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes((Get-AsanaPersonalToken)))
        }

        
    }
    End
    {
    }
}