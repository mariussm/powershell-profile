Function Add-CodeSignature {
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=1)]
        $ThumbPrint,

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   Position=0)]
        [String[]] $Files
    )

    Begin
    {
    }
    Process
    {
        if($ThumbPrint)
        {
            $Certificate = Get-CodeSigningCertificate | Where{$_.Thumbprint -eq $Thumbprint}
        }
        else
        {
            $Certificates = Get-CodeSigningCertificate
            $Certificate = $Certificates | Select -First 1
            if(($Certificates | measure).Count -gt 1)
            {
                Write-Warning "Warning, there are multiple signing certificates."
            }
        }

        if(!$Certificate)
        {
            throw New-Object System.Exception("Cannot find code signing certificate with thumbprint $Thumbprint")
        }

        Set-AuthenticodeSignature -Certificate $Certificate -FilePath $Files

        if($psise) {
            $FullFileNames = $Files | dir | select -exp FullName
            $psise.CurrentPowerShellTab.Files | Where{$Fullpath = $_.FullPath; $FullFileNames | Where{$_ -eq $FullPath}} | Foreach {
                if(!$_.IsSaved){
                    $_.Save()
                }
                [int] $CaretLine = $_.Editor.CaretLine
                [int] $CaretColumn = $_.Editor.CaretColumn
                $_.Editor.Text = Get-Content -Raw $_.FullPath
                $_.Editor.SetCaretPosition($CaretLine, $CaretColumn)
                $_.Save()
                
            }
        }
    }
    End
    {
    }

}
