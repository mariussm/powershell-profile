Function Search-IseFiles {
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$false,
                   Position=0)]
        [ScriptBlock] $Where,

        # Param1 help description
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=1)]
        [string] $Path = "$((Split-Path -Parent $profile))\isefiles\"
    )

    Begin
    {
    }
    Process
    {
        if(!(Test-path $Path)) {
            Write-Error "No such path: $Path"
            return;
        }

        dir $Path | Where {
            Get-ContentAsString -Path $_.FullName | Where -FilterScript $Where
        }
    }
    End
    {
    }

}
