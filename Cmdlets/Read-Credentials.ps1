Function Read-Credentials {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True,Position=0)]
        [string]$File
    )

    $obj = Import-Clixml -Path $File
    return New-Object System.Management.Automation.PSCredential($obj.username, ($obj.password | ConvertTo-SecureString))

}
