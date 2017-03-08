Function Get-FileFromURI {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        [String] $URI,

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=1)]
        [String] $DestinationFileName
    )

    Begin
    {
    }
    Process
    {
        $_DestinationFileName = $DestinationFileName

        $wc = New-Object System.Net.WebClient
        if(!$_DestinationFileName) {
            $tempURI = $URI -replace "http://",""
            $_DestinationFileName = (Split-Path -Leaf $tempURI) -replace "%20"," "
            Write-Verbose "Setting destination file name to: $_DestinationFileName"
        }

        if($_DestinationFileName.Substring(1,1) -ne ":") {
            $_DestinationFileName = (pwd).Path + "\" + $_DestinationFileName
            Write-Verbose "Full path: $_DestinationFileName"
        }

        Write-Verbose "Downloading $uri -> $_DestinationFileName"
        $wc.DownloadFile($uri, $_DestinationFileName)
    }
    End
    {
    }

}
