Function Get-RandomPassword {
    [CmdletBinding()]
    [OutputType([string])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=0)]
        $Length = 32
    )

    Begin
    {
        $possibleCharacters = "a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z","A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","1","2","3","4","5","6","7","8","9"

    }
    Process
    {
        if($Length -lt 3) {
            Write-Error "Length too small"
        }
        do {
            $password = (1..$Length | foreach{$possibleCharacters | Get-Random -Count 1}) -join ""
        } while($password -cnotmatch "[a-z]" -or $password -cnotmatch "[A-Z]" -or $password -notmatch "[1-9]")
        return $password
    }
    End
    {
    }

}
