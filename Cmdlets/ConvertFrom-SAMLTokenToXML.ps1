function ConvertFrom-SAMLTokenToXML
{
    [CmdletBinding()]
    [Alias()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        [String] $String
    )

    Begin
    {
    }
    Process
    {
        if($String.Substring(0,1) -ne "<") {
            Write-Verbose "Detected token as not saml, trying to convert from base64 first"
            $String = ConvertFrom-Base64 $String
        }

        return ([xml] $String)
    }
    End
    {
    }
}