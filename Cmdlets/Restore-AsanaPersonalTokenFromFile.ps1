function Restore-AsanaPersonalTokenFromFile
{
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [String] $File
    )

    Begin
    {
    }
    Process
    {
        if(Test-path $file) {
            $AsanaToken = ([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR((ConvertTo-SecureString ([IO.File]::ReadAllText((Resolve-Path $File))).Trim()))))
            Set-AsanaPersonalToken -Token $AsanaToken
        } else {
            throw "Could not find file $file"
        }
    }
    End
    {
    }
}