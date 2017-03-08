function Connect-ExchangeOnline{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True,Position=0)]
        [System.Management.Automation.PSCredential]$Credentials
    )
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Authentication Basic -AllowRedirection -Credential $Credentials
    Import-PSSession $session -DisableNameChecking
}