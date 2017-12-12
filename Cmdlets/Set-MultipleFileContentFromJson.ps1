function Set-MultipleFileContentFromJson
{
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Json = (Read-Host),

        $Path = (Pwd).Path
    )

    Begin
    {
    }
    Process
    {
        $t = $Json | ConvertFrom-Json
        $t | Foreach {
            Write-Verbose "$((Join-Path $Path $_.Name))"
            Set-Content -Encoding Byte -Path (Join-Path $Path $_.Name) -Value ([System.Convert]::FromBase64String($_.Content))
        }
    }
    End
    {
    }
}