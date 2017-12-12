function Get-MultipleFileContentAsBase64
{
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Path
    )

    Begin
    {
    }
    Process
    {
        ls $Path | Foreach {
            [PSCustomObject] @{
                Name = $_.Name
                Content = (Get-ContentAsBase64 -Path $_.FullName)
            }
        } | ConvertTo-Json 
    }
    End
    {
    }
}