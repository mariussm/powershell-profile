<#
.Synopsis
    Returns the federation metadata as XML
.DESCRIPTION
    Returns the federation metadata as XML
.EXAMPLE
    Get-FederationMetadata "adfs.goodworkaround.com"
#>
function Get-FederationMetadata
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
    return Invoke-RestMethod -Uri ("https://{0}/FederationMetadata/2007-06/FederationMetadata.xml" -f $FQDN)
    }
    End
    {
    }
}