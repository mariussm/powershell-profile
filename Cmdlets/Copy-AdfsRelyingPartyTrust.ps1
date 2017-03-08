<#
.Synopsis
   Copies relying party trust
.DESCRIPTION
   Copies relying party trust
.EXAMPLE
   Copy-AdfsRelyingPartyTrust "Intranett Test" "Intranett Prod" "urn:sharepoint:prod"
#>
function Copy-AdfsRelyingPartyTrust
{
    [CmdletBinding()]
    [OutputType([int])]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   Position=0)]
        $SourceRelyingPartyTrustName,

        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   Position=1)]
        $NewRelyingPartyTrustName,

        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   Position=2)]
        $NewRelyingPartyTrustIdentifier
    )

    Begin
    {
    }
    Process
    {
        $SourceRelyingPartyTrust  = Get-AdfsRelyingPartyTrust -Name $SourceRelyingPartyTrustName

        $exceptedAttributes = @("ConflictWithPublishedPolicy","OrganizationInfo","ProxyEndpointMappings","LastUpdateTime","PublishedThroughProxy","LastMonitoredTime")
        $parameters = @{}
        $SourceRelyingPartyTrust | Get-Member -MemberType Property | where{$_.name -notin $exceptedAttributes} | foreach {
            if($SourceRelyingPartyTrust.($_.Name) -ne $null) {
                $parameters[$_.Name] = $SourceRelyingPartyTrust.($_.Name)
            }
        }
        $parameters.Name = $NewRelyingPartyTrustName
        $parameters.Identifier = $NewRelyingPartyTrustIdentifier
        
        Add-AdfsRelyingPartyTrust @parameters
    }
    End
    {
    }
}