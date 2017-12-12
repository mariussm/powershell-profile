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
function Invoke-CommandWithExceptionsAsErrors
{
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   Position=0)]
        [System.Management.Automation.ScriptBlock]
        $ScriptBlock
    )

    Begin
    {
    }
    Process
    {
        try {
            Invoke-Command -ScriptBlock $ScriptBlock
        } catch {
            Write-Error -Exception $_
        }   
    }
    End
    {
    }
}