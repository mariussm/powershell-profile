<#
.Synopsis
   Returns the ADFS token signing and encryption certificates
.DESCRIPTION
   Returns the ADFS token signing and encryption certificates
.EXAMPLE
   Get-AdfsCertificates adfs.goodworkaround.com
#>
function Get-AdfsCertificates
{
    [CmdletBinding()]
    Param
    (
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

        $metadata.EntityDescriptor.RoleDescriptor.KeyDescriptor | foreach {
            $tempfile = "{0}\adfsTempCert.cer" -f $env:temp
            $_.KeyInfo.X509Data.X509Certificate | Set-Content -Path $tempfile

            $cert = (New-Object System.Security.Cryptography.X509Certificates.X509Certificate2)
            $cert.Import($tempfile)

            New-Object -TypeName PSCustomObject -Property @{
                FoundIn = "KeyDescriptor"
                Use = $_.Use
                Subject = $cert.Subject
                Issuer = $cert.Issuer
                ThumbPrint = $cert.Thumbprint
                NotBefore = $cert.NotBefore
                NotAfter = $cert.NotAfter
                Data = $_.KeyInfo.X509Data.X509Certificate
            }
        }

        $tempfile = "{0}\adfsTempCert.cer" -f $env:temp
        $metadata.EntityDescriptor.Signature.KeyInfo.X509Data.X509Certificate | Set-Content -Path $tempfile
        $cert = (New-Object System.Security.Cryptography.X509Certificates.X509Certificate2)
        $cert.Import($tempfile)

        New-Object -TypeName PSCustomObject -Property @{
            FoundIn = "Active Signature"
            Use = "signing"
            Subject = $cert.Subject
            Issuer = $cert.Issuer
            ThumbPrint = $cert.Thumbprint
            NotBefore = $cert.NotBefore
            NotAfter = $cert.NotAfter
            Data = $metadata.EntityDescriptor.Signature.KeyInfo.X509Data.X509Certificate
        }
    }
    End
    {
    }
}