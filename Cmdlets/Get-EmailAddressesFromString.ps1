<#
.Synopsis
   Returns all email addresses from a string
.DESCRIPTION
   Returns all email addresses from a string
.EXAMPLE
   "randomstring" | Get-EmailAddressesFromString
#>
function Get-EmailAddressesFromString
{
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $String,

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=1)]
        $Pattern = "[0-9a-zA-Z]+@[0-9a-zA-Z\.]+"
    )

    Begin
    {
    }
    Process
    {
        $Pattern = "[0-9a-zA-Z\._-]+@[0-9a-zA-Z][0-9a-zA-Z\._-]+\.[a-zA-Z0-9]{2,}"
        [System.Text.RegularExpressions.Regex]::Matches($String, $Pattern) | foreach{$_.Value}
    }
    End
    {
    }
}