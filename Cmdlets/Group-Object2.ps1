Function Group-Object2 {
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=3)]
        $Object,

        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   Position=0)]
        [String] $Property,

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=1)]
        [String] $ExpandProperty = $null,

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=2)]
        [Boolean] $AsHashTable = $false
    )

    Begin
    {
        $_workingHashmap = @{}
    }
    Process
    {
        $groupValue = $Object.$Property
        if(!$groupValue) {
            Write-Verbose "Empty groupValue"
            $groupValue = ""
        }
        
        $groupObject = $Object
        if($ExpandProperty -ne $null -and $ExpandProperty -ne "") {
            Write-Verbose "Expanding property $ExpandProperty"
            $groupObject = $Object.$ExpandProperty
        }

        if(!$_workingHashmap[$groupValue]) {
            $_workingHashmap[$groupValue] = @()
        }
        $_workingHashmap[$groupValue] += $groupObject
    }
    End
    {
        if($AsHashTable) {
            return $_workingHashmap
        } else {
            $_workingHashmap.Keys | foreach {
                $_t = @{
                    Count = $_workingHashmap[$_].Count
                    Name = $_ 
                    Group = $_workingHashmap[$_]
                }
                New-Object -TypeName PSCustomObject -Property $_t 
            }
        }
    }

}
