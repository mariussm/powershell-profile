Function Get-HashValue {
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=1)]
        [String] $String,

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   Position=0)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("SHA1", "SHA256", "MD5")]
        [String] $Algorithm = "SHA1"
    )

    Process
    {
        if($Algorithm -eq "SHA1") {
            $hasher = new-object System.Security.Cryptography.SHA1Managed
        } elseif($Algorithm -eq "SHA256") {
            $hasher = new-object System.Security.Cryptography.SHA256Managed
        } elseif($Algorithm -eq "MD5") {
            $hasher = new-object System.Security.Cryptography.MD5CryptoServiceProvider
        }
        $toHash = [System.Text.Encoding]::UTF8.GetBytes($String)
        $hashByteArray = $hasher.ComputeHash($toHash)
        $res = ""
        foreach($byte in $hashByteArray)
        {
             $res += [System.String]::Format("{0:X2}", $byte)
        }
        return $res;
    }

}
