<#
.Synopsis
   Function to invoke FIM run profiles
.DESCRIPTION
   This method uses WMI to trigger FIM run profiles.
.EXAMPLE
   Invoke-FIMRunProfile "AD" "Full import"
   
   This example trigger the "Full import" run profile on the "AD" management agent
.EXAMPLE
   The following example trigger the "Full import" run profile on the "AD" management agent

   Invoke-FIMRunProfile -MA "AD" -RunProfile "Full import"   
.EXAMPLE
   Invoke-FIMRunProfile "AD"
   
   This will trigger the "Delta import Delta sync" run profile on the "AD" management agent
#>
function Invoke-FIMRunProfile
{
    [CmdletBinding()]
    [OutputType([object])]
    Param
    (
        # The Management Agent name
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$false,
                   Position=0)]
        [string] $MA,

        # The run profile to trigger
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=1)]
        [string[]] $RunProfile = @("Delta import Delta sync"),
        
        # Only trigger RunProfile if there are something to export
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=2)]
        [boolean] $DoNotRunWhenNoExports = $false
            
        
    )

    Begin
    {
        # Connect to database
        Write-Verbose ("Connecting to WMI root/MicrosoftIdentityIntegrationServer class MIIS_ManagementAgent")
        $wmi = Get-WmiObject -Namespace root/MicrosoftIdentityIntegrationServer -Class MIIS_ManagementAgent
    }
    Process
    {
        $WMIMAs = ($wmi | where{$_.Name -like $MA})
        
        foreach($WMIMA in $WMIMAs) {
            if($DoNotRunWhenNoExports -and (([int]$WMIMA.NumExportAdd().ReturnValue + [int]$WMIMA.NumExportDelete().ReturnValue + [int]$WMIMA.NumExportUpdate().ReturnValue) -eq 0)) {
                Write-Verbose "Found nothing to export"
                $result = @{ReturnValue="Nothing to export"}
                New-Object -TypeName PSObject -Property @{"Management Agent"=$WMIMA.Name;"Run Profile"=$RunProfile;Result=$result.ReturnValue}
            } else {
                # Execute WMI query to run the run profile and store the result in $result
                Write-Verbose ("Executing run profile ""{0}""" -f $RunProfile)
                $RunProfile | Foreach {
                    $result = $WMIMA.Execute($_)
                    New-Object -TypeName PSObject -Property @{"Management Agent"=$WMIMA.Name;"Run Profile"=$_;Result=$result.ReturnValue}
                }
            }    
        }
        
        
    }
    End
    {
        
    }
}