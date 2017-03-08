Function ConvertFrom-Base64 {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True,
                   Position=0,
                   ValueFromPipeline=$true)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string]
        $Base64String
    )

    Begin{}
    Process{
        return [System.Text.Encoding]::UTF8.GetString(([System.Convert]::FromBase64String($Base64String)));
    }
    End{}

}
