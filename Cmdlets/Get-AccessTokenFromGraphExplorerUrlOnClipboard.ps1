function Get-AccessTokenFromGraphExplorerUrlOnClipboard
{
    [CmdletBinding()]
    [Alias()]
    Param
    ()

    Process
    {
        $first = $true
        do {
            if(!$first) {
                Sleep -Seconds 1   
            }
            $first = $false 

            Write-Verbose "Trying to get Graph Explorer URL from clipboard"
            $url = Get-Clipboard
            if($url -ne $null -and $url.StartsWith("https://developer.microsoft.com/en-us/graph/graph-explorer#access_token=")) {
                $token = $url -split "[=&]" | Select -Index 1
            }
        } while($token -eq $null -or !$token.StartsWith("ey"))
        $token
    }
}