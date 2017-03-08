<#
.Synopsis
    Creates a copy of the input set(s)
.DESCRIPTION
    Creates a copy of the input set(s)
.EXAMPLE
    Get-FIMObjectByXPath '/Set[DisplayName="All People"]' | Copy-FIMSet
#>
function Copy-FIMSet
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
            ValueFromPipeline=$true,
            Position=0)]
        $Set,

        [Parameter(Mandatory=$false,
            ValueFromPipeline=$false,
            Position=1)]
        $Prefix = "- [COPY] "
    )
    Begin
    {
    }
    Process
    {
        if($Set.DisplayName -and $Set.Filter -and $Set.ObjectType -eq "Set") {
            $changes = @{
                DisplayName = ("{0}{1}" -f $Prefix, $Set.DisplayName)
                Filter = $Set.Filter
            }
            New-FimImportObject -ObjectType Set -ApplyNow -PassThru -State Create -Changes $changes
        } else 
        {
            Write-Error "Input object not valid"
        }
    }
    End
    {
    }
}