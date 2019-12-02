function Get-AccessTokenFromGraphExplorerUrlOnClipboard
{
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [string[]] $RequiredScopes = @()
    )

    Process
    {
        $first = $true
        do {
            if(!$first) {
                Start-Sleep -Seconds 1   
            }
            $first = $false 

            Write-Verbose "Trying to get Graph Explorer URL from clipboard, with requires scopes: $RequiredScopes" 
            $url = Get-Clipboard
            if(![string]::IsNullOrEmpty($url) -and $url.StartsWith("https://developer.microsoft.com/en-us/graph/graph-explorer#access_token=")) {
                Write-Verbose "Found relevant url on the clipboard, starting verification"
                $token = $url -split "[=&]" | Select-Object -Index 1

                # Fix padding length for base 64
                $token2 = $token.Split(".")[1] + [String]::new("=", 4 - ($token.Split(".")[1].Length % 4))
                
                # Converting from json
                $tokenObject = [System.Text.Encoding]::UTF8.GetString(([System.Convert]::FromBase64String($token2))) | ConvertFrom-Json

                # Checking for expiry (with 5 minutes minimum required)
                $date = get-date "1/1/1970"
                if((Get-Date).AddMinutes(5) -gt $date.AddSeconds($tokenObject.exp ).ToLocalTime()) {
                    Write-Verbose "Token on clipboard is expired, not using it"
                    $token = ""
                }
                
                # Check scopes
                if($RequiredScopes.Count -gt 0) {
                    $tokenScopes = $tokenObject.scp.Split(" ")
                    $RequiredScopes | Where-Object {$_ -notin $tokenScopes} | ForEach-Object {
                        Write-Verbose "Token did not contain scope $($_), not using it"
                        $token = ""
                    }
                }
                
            }
        } while([string]::IsNullOrEmpty($token) -or !$token.StartsWith("ey"))
        
        Write-Verbose "Token found"
        $token
    }
}