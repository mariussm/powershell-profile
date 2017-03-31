<#
.Synopsis
   Creates a red-black tree optimal for searching
.EXAMPLE
   $Tree = dir c:\windows | select -exp name | New-BinarySearchTree
   $Tree.Contains("System32")
   $Tree.Contains("system32")
#>
function New-BinarySearchTree
{
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        [String] $Value
    )

    Begin
    {
        Add-Type -AssemblyName System.Core
        Add-Type -AssemblyName System.Collections
        $Tree = new-object System.Collections.Generic.SortedSet[String]
    }
    Process
    {
        $Tree.Add($Value) | Out-Null
    }
    End
    {
        return $Tree
    }
}