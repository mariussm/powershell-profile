Function Load-Credential {
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [String] $Name = $null
    )

    Begin
    {
    }
    Process
    {
        $_file = "$($env:APPDATA)\credentials.csv"

        if(!(Test-path $_file)) {
            Write-Error "No such file: $_file" -ErrorAction Stop
        }
        
        if(!$Name -or $Name.Length -eq 0) {
            $_Credential = Import-Csv $_file | Out-Gridview -OutputMode Single -Title "Choose credential"
        } else {
            $_Credential = Import-Csv $_file | Where{$_.Name -like $Name}
            if(($_Credential| measure).Count -gt 1) {
                $_Credential = $_Credential | Out-Gridview -OutputMode Single -Title "Choose credential"
            }
        }

        if(!$_Credential) {
            Write-Error "No such credential: $Name"
        } else {
            return New-Object System.Management.Automation.PSCredential($_Credential.Username, ($_Credential.Password | ConvertTo-SecureString))
        }
    }
    End
    {
    }

}
