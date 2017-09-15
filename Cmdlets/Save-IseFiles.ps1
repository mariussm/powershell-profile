Function Save-IseFiles {
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=0)]
        [string] $Path = "$((Split-Path -Parent $profile))\isefiles\",
        

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=1)]
        [boolean] $Confirm = $true
    )

    Begin
    {
    }
    Process
    {
        if(!(Test-path $Path)) {
            mkdir $Path | Out-Null
        }

        $psISE.PowerShellTabs | Foreach {
            $_.Files | 
            where{!$_.IsSaved} |
                foreach {
                    if((Test-Path $_.FullPath)) {
                        Write-Verbose "File already exists, so this is an unsaved file: $($_.FullPath)"
                        $result = "y"
                        if($Confirm) {
                            $result = Read-Host "Save $($_.FullPath)? (y/N)"
                        }

                        if($result -eq "y") {
                            Write-Verbose "Saving: $($_.FullPath)"
                            $_.Save()
                        }
                    } else {
                        $ActualPath = (Join-Path $Path ([guid]::newguid()).ToString()) + ".ps1"
                        Write-Verbose "File does not exist, so this is a temp file - saving to $ActualPath"
                        $_.SaveAs($ActualPath)
                    }
                }
        }
    }
    End
    {
    }

}
