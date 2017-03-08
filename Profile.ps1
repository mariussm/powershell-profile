function Update-PowerShellProfile
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=0)]
        $Branch = "master",

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=1)]
        $Repository = "mariussm/powershell-profile"
    )

    Invoke-RestMethod -OutFile $profile -Uri "https://raw.githubusercontent.com/$Repository/$Branch/Profile.ps1"
    Invoke-RestMethod -OutFile (Join-Path (Split-Path -Parent -Path $profile) "Cmdlets.ps1") -Uri "https://raw.githubusercontent.com/$Repository/$Branch/Cmdlets.ps1"
    Invoke-RestMethod -OutFile (Join-Path (Split-Path -Parent -Path $profile) "FIMPowerShellModule.psm1") -Uri "https://raw.githubusercontent.com/$Repository/$Branch/Modules/FIMPowerShellModule.psm1"
    . $profile
}

. (Join-Path (Split-Path -Parent -Path $profile) "Cmdlets.ps1")
Import-Module (Join-Path (Split-Path -Parent -Path $profile) "FIMPowerShellModule.psm1")

if($env:USERPROFILE -and (pwd).Path.StartsWith("C:\Windows")) {
    cd $env:USERPROFILE
}