function Disconnect-ExchangeOnline {
    [CmdletBinding()]
    Param()
    Get-PSSession | ?{$_.ComputerName -like "*outlook.com"} | Remove-PSSession
}