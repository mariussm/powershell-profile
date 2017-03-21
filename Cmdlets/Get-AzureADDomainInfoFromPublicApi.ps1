function Get-AzureADDomainInfoFromPublicApi
{
    [CmdletBinding()]
    [Alias()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        [String] $Domain
    )

    Begin
    {
    }
    Process
    {
        $Url = "https://login.microsoftonline.com/common/userrealm/?user=someone.random@" + $Domain + "&checkForMicrosoftAccount=true&api-version=2.1"
        Invoke-RestMethod $Url
    }
    End
    {
    }
}