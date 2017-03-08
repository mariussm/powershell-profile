<#
.Synopsis
    Returns the federation metadata URL
.DESCRIPTION
    Returns the federation metadata URL
.EXAMPLE
    Get-FederationMetadataURL "adfs.goodworkaround.com"
#>
function Get-FederationMetadataURL
{
    [CmdletBinding()]
    [OutputType([xml])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                ValueFromPipeline=$true,
                Position=0)]
        $FQDN
    )
 
    Begin
    {
    }
    Process
    {
    return ("https://{0}/FederationMetadata/2007-06/FederationMetadata.xml" -f $FQDN)
    }
    End
    {
    }
}