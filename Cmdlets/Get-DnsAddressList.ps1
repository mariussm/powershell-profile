Function Get-DnsAddressList {
    param(
        [parameter(Mandatory=$true)][Alias("Host")]
          [string]$HostName)

    try {
        return [System.Net.Dns]::GetHostEntry($HostName).AddressList
    }
    catch [System.Net.Sockets.SocketException] {
        if ($_.Exception.ErrorCode -ne 11001) {
            throw $_
        }
        return = @()
    }

}
