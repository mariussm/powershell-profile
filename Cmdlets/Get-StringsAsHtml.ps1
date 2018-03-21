function Get-StringsAsHtml
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([string])]
    Param
    (
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=0)]
        [String] $Style,

        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=1)]
        [String] $StringObject
    )

    Begin
    {
        $Html = "<html><head><style type='text/css'>$Style</style></head><body>`n"
    }
    Process
    {
        $Html += $StringObject + "`n"
    }
    End
    {
        return $Html + "</body></html>"
    }
}