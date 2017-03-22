<#
.Synopsis
    Returns the federation metadata as XML
.DESCRIPTION
    Returns the federation metadata as XML
.EXAMPLE
    Get-AzureADFederationMetadata "microsoft.com"
#>
function Get-AzureADFederationMetadata
{
    [CmdletBinding()]
    [OutputType([xml])]
    Param
    (
        [Parameter(Mandatory=$true,
                ValueFromPipeline=$true,
                Position=0)]
        [String] $Domain,

        [Parameter(Mandatory=$false,
                ValueFromPipeline=$false,
                Position=1)]
        [String] $STS = "sts.windows.net"
    )
 
    Begin
    {
    }
    Process
    {
        $XDocument = [System.Xml.Linq.XDocument]::Load( ("https://$STS/{0}/FederationMetadata/2007-06/FederationMetadata.xml" -f $Domain))
        [xml] $XDocument
    }
    End
    {
    }
}