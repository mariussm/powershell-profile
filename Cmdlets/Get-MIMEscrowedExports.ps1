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
function Get-MIMEscrowedExports
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    [OutputType([int])]
    Param
    (
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=1)]
        [String] $MA = "Visma HRM Security",

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=1)]
        [String] $CSExportPath = "C:\Program Files\Microsoft Forefront Identity Manager\2010\Synchronization Service\Bin\csexport.exe"
    )

    Begin
    {
    }
    Process
    {
        $Tempfile = Join-Path $env:TEMP (([guid]::newguid()).ToString() + ".xml")
        . "C:\Program Files\Microsoft Forefront Identity Manager\2010\Synchronization Service\Bin\csexport.exe" $MA $Tempfile "/f:e" | Out-Null
        
        [xml] $xml = gc $Tempfile -Encoding UTF8

        $xml.'cs-objects'.'cs-object' | foreach {
            $csobject = $_ # $csobject = $xml.'cs-objects'.'cs-object' | select -index 2
            Write-Verbose "Working on object: $($csobject.'cs-dn')"

            $csdn = $csobject.'cs-dn'
            $objectoperation = $csobject."escrowed-export".delta.operation

            $csobject."escrowed-export".delta.attr | foreach {
                $attributename = $_.name
                $attributeoperation = $_.operation 
                $attributetype = $_.type

                $_.value | foreach {
                    if($_.operation) {
                        $value = $_."#text"
                        $valueoperation = $_.operation
                    } else {
                        $value = $_
                        $valueoperation = "none"
                    }

                    [PSCustomObject] @{
                        csdn = $csdn
                        objectoperation = $objectoperation   
                        attributename = $attributename
                        attributeoperation = $attributeoperation
                        attributetype = $attributetype
                        valueoperation = $valueoperation
                        value = $value
                    }
                }
            }
        }
    }
    End
    {
    }
}