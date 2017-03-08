<#
.Synopsis
   Returns fim management agent matching pattern
.DESCRIPTION
   This method uses WMI to get and return FIM Management Agents
.EXAMPLE
   Get-FIMManagementAgent "SP - *"
#>
function Get-FIMManagementAgent
{
    [CmdletBinding()]
    [OutputType([object])]
    Param
    (
        # The Management Agent name
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$false,
                   Position=0)]
        [string] $MA
    )

    Begin
    {
        # Connect to database
        Write-Verbose ("Connecting to WMI root/MicrosoftIdentityIntegrationServer class MIIS_ManagementAgent")
        $wmi = Get-WmiObject -Namespace root/MicrosoftIdentityIntegrationServer -Class MIIS_ManagementAgent
    }
    Process
    {
        return ($wmi | where{$_.Name -like $MA})
    }
    End
    {
        
    }
}