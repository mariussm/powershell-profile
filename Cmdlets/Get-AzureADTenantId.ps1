<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Get-AzureADTenantId
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([string])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $DomainName
    )

    Begin
    {
        Load-Assembly System.Xml.Linq | Out-Null
    }
    Process
    {
        $FederationMetadata = Get-AzureADFederationMetadata -Domain $DomainName
        $FederationMetadata.EntityDescriptor.entityID -split "/" | where{$_ -match "^[a-zA-Z0-9]{8}-[a-zA-Z0-9]{4}-[a-zA-Z0-9]{4}-[a-zA-Z0-9]{4}-[a-zA-Z0-9]{12}$"}
    }
    End
    {
    }
}