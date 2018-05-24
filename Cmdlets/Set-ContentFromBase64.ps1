<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Set-ContentFromBase64
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Base64Value = (Read-Host -Prompt "Base64 value"),

        [String] $File
    )

    Begin
    {
    }
    Process
    {
        $decoded = [System.Convert]::FromBase64String($Base64Value)
        set-content -Path $File -Value $decoded -Encoding Byte
    }
    End
    {
    }
}