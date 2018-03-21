<#
.Synopsis
   Returns all email addresses from a string
.DESCRIPTION
   Returns all email addresses from a string
.EXAMPLE
   "randomstring" | Get-Matches
#>
function Get-Matches
{
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   Position=0)]
        $Pattern,
        
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=1)]
        $String
    )

    Begin
    {
    }
    Process
    {
        [System.Text.RegularExpressions.Regex]::Matches($String, $Pattern) | foreach{$_.Value}
    }
    End
    {
    }
}