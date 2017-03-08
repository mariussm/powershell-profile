Function Store-Credentials {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [string]$File,
	
        [Parameter(Mandatory=$True)]
        [System.Management.Automation.PSCredential]$Credentials
    )

    $info = @{username=$Credentials.UserName;password=($Credentials.Password | ConvertFrom-SecureString)}
    $obj = New-Object -TypeName PSObject -Property $info
    $obj | Export-Clixml -Path $File

}
