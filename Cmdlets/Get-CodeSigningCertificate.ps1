Function Get-CodeSigningCertificate {
    [CmdletBinding()]
    [Alias()]
    [OutputType([System.Security.Cryptography.X509Certificates.X509Certificate])]
    Param
    (
    )

    Begin
    {
    }
    Process
    {
        get-childitem Cert:\CurrentUser\my -CodeSigningCert
    }
    End
    {
    }

}
