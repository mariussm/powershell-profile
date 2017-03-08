Function Save-Credential {
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [String] $Name,

        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [System.Management.Automation.PSCredential] $Credential,

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=2)]
        [String] $Description
    )

    Begin
    {
    }
    Process
    {
        $_file = "$($env:APPDATA)\credentials.csv"
        
        if(!(test-path $_file)) {
            Set-Content -Path $_file -Value "Name,Description,Username,Password"
        }

        $_Credentials = @(Import-Csv $_file | Where{$_.Name -ne $Name})

        $_Credentials += [PSCustomObject] @{
            Name = $Name 
            Description = $Description
            Username = $Credential.UserName
            Password = ($Credential.Password | ConvertFrom-SecureString)
        }
        
        $_Credentials | Export-Csv $_file -NoTypeInformation
    }
    End
    {
    }

}
