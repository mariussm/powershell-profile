<#
.Synopsis
   Returns fim run history for all or one MA
.DESCRIPTION
   Returns fim run history for all or one MA
.EXAMPLE
   Get-FIMRunHistory "SharePoint Internal"
#>
function Get-FIMRunHistory
{
    [CmdletBinding()]
    [OutputType([object])]
    Param
    (
        # The Management Agent name
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=0)]
        [string] $MA,

        # Return only first match
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=1)]
        [bool] $FirstOnly = $true
    )

    Begin
    {
    }
    Process
    {
        if($MA) 
        {
            if($FirstOnly) 
            {
                return Get-WmiObject -Namespace root/MicrosoftIdentityIntegrationServer -Class MIIS_RunHistory -Filter ("MaName='{0}'" -f $MA) | select -First 1
            } else 
            {
                return Get-WmiObject -Namespace root/MicrosoftIdentityIntegrationServer -Class MIIS_RunHistory -Filter ("MaName='{0}'" -f $MA)
            }
        }
        else 
        {
            if($FirstOnly)
            {
                return Get-WmiObject -Namespace root/MicrosoftIdentityIntegrationServer -Class MIIS_RunHistory | select -First 1
            } 
            else 
            {
                return Get-WmiObject -Namespace root/MicrosoftIdentityIntegrationServer -Class MIIS_RunHistory
            }
        }
        return ($wmi | where{$_.Name -like $MA})
    }
    End
    {   
    }
}