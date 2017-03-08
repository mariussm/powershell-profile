<#
.Synopsis
   Returns the thumbprint of the ADFS token signing certificate
.DESCRIPTION
   Returns the thumbprint of the ADFS token signing certificate
.EXAMPLE
   Get-AdfsTokenSigningThumbprint adfs.goodworkaround.com
#>
function Get-AdfsTokenSigningThumbprint
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$false,
                   Position=0)]
        $ADFS
    )

    Begin
    {
    }
    Process
    {
        $metadata = Invoke-RestMethod -Uri ("https://{0}/FederationMetadata/2007-06/FederationMetadata.xml" -f $ADFS)
        $tempfile = "{0}\adfsTempCert.cer" -f $env:temp
        $metadata.EntityDescriptor.Signature.KeyInfo.X509Data.X509Certificate | Set-Content -Path $tempfile
        $cert = (New-Object System.Security.Cryptography.X509Certificates.X509Certificate2)
        $cert.Import($tempfile)

        return $cert.Thumbprint
    }
    End
    {
    }
}