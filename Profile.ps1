function Update-PowerShellProfile
{
    [CmdletBinding()]
    Param()

    Invoke-RestMethod -OutFile $profile -Uri "https://raw.githubusercontent.com/mariussm/powershell-profile/master/Profile.ps1"
    Invoke-RestMethod -OutFile (Join-Path (Split-Path -Parent -Path $profile) "Cmdlets.ps1") -Uri "https://raw.githubusercontent.com/mariussm/powershell-profile/master/Cmdlets.ps1"
    Invoke-RestMethod -OutFile (Join-Path (Split-Path -Parent -Path $profile) "FIMPowerShellModule.psm1") -Uri "https://raw.githubusercontent.com/mariussm/powershell-profile/master/Modules/FIMPowerShellModule.psm1"
    . $profile
}

. (Join-Path (Split-Path -Parent -Path $profile) "Cmdlets.ps1")
Import-Module (Join-Path (Split-Path -Parent -Path $profile) "FIMPowerShellModule.psm1")