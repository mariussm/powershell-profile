<#
.Synopsis
    Creates a copy of a set
.DESCRIPTION
    Creates a copy of a set
.EXAMPLE
    Copy-FIMSetByName "All People" "All People 2"
#>
function Copy-FIMSetByName
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
            ValueFromPipeline=$false,
            Position=0)]
        $Source,

        [Parameter(Mandatory=$false,
            ValueFromPipeline=$false,
            Position=1)]
        $Destination
    )
    Begin
    {
    }
    Process
    {
        $SourceSet = Get-FIMObjectByXPath "/Set[DisplayName=""$Source""]"
        if(!$SourceSet) {
            Write-Error "Set not found"
        } else {
            $changes = @{
                DisplayName = $Destination
                Filter = $SourceSet.Filter
            }
            New-FimImportObject -ObjectType Set -ApplyNow -PassThru -State Create -Changes $changes
        }
    }
    End
    {
    }
}