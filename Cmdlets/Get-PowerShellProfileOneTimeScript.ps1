Function Get-PowerShellProfileOneTimeScript {
    [CmdletBinding()]
    [Alias()]
    [OutputType([string])]
    Param()

    '"https://dl.dropboxusercontent.com/u/6872078/PS/365.psm1",
"https://dl.dropboxusercontent.com/u/6872078/PS/ad.psm1",
"https://dl.dropboxusercontent.com/u/6872078/PS/adfs.psm1",
"https://dl.dropboxusercontent.com/u/6872078/PS/dnvgl.psm1",
"https://dl.dropboxusercontent.com/u/6872078/PS/fim.psm1",
"https://dl.dropboxusercontent.com/u/6872078/PS/fimpsmodule.psm1",
"https://dl.dropboxusercontent.com/u/6872078/PS/gwrnd.psm1",
"https://dl.dropboxusercontent.com/u/6872078/PS/linqxml.psm1",
"https://dl.dropboxusercontent.com/u/6872078/PS/gwrnddsc.psm1",
"https://dl.dropboxusercontent.com/u/6872078/PS/tools.psm1" | foreach {
    Write-Verbose "Downloading file $($_)" -Verbose
    $wc = New-Object System.Net.WebClient
    $file = "{0}\{1}" -f $env:TEMP, ($_ -split "/" | select -last 1)
    $wc.DownloadFile($_, $file)

    Import-Module $file -DisableNameChecking
    Remove-Item $file -Force
}'

}
