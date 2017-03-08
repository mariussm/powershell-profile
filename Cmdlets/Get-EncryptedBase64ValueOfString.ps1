Function Get-EncryptedBase64ValueOfString {
    [CmdletBinding()]
    [Alias()]
    [OutputType([string])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [string] $InputString,

        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$false,
                   Position=0)]
        [String] $Thumbprint
    )

    Begin
    {
        $Cert = ((dir Cert:\LocalMachine\my) | ?{$_.PrivateKey.KeyExchangeAlgorithm -and $_.Verify()}) , ((dir Cert:\CurrentUser\my) | ?{$_.PrivateKey.KeyExchangeAlgorithm -and $_.Verify()}) | Where{$_.Thumbprint -eq $Thumbprint}
        if(!$Cert) {
            throw "No certificate with thumbprint $Thumbprint found"
        }
    }
    Process
    {
        $EncodedInputString = [system.text.encoding]::UTF8.GetBytes($InputString)
        $EncryptedBytes = $Cert.PublicKey.Key.Encrypt($EncodedInputString, $true)
        return [System.Convert]::ToBase64String($EncryptedBytes)

    }
    End
    {
    }

}
