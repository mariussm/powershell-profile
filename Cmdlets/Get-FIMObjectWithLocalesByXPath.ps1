<#
.Synopsis
   Returns objects matching xpath
.DESCRIPTION
   Returns objects matching xpath
.EXAMPLE
   Get-FIMObjectByXPath "/testUser"
#>
function Get-FIMObjectWithLocalesByXPath
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $XPath
    )

    Begin
    {
    }
    Process
    {
        $res = Export-FimConfig -CustomConfig $XPath -Uri "http://localhost:5725/" -AllLocales
        if($res) {
            $res | Foreach {
                $PSObject = $_ | Convert-FimExportToPSObject
                $PSObject | Add-Member -MemberType NoteProperty -Name Locale -Value @{}
                if($_.ResourceManagementObject.LocalizedResourceManagementAttributes) {
                    $_.ResourceManagementObject.LocalizedResourceManagementAttributes | foreach {
                        $culture = $_.Culture
                    
                        $attributes = @{}
                        $_.ResourceManagementAttributes | Foreach {
                            if($_.IsMultiValue) {
                                $attributes[$_.AttributeName] = $_.Values
                            } else {
                                $attributes[$_.AttributeName] = $_.Value
                            }
                        }
                        $PSObject.Locale[$culture] = New-Object -TypeName PSObject -Property $attributes
                    }
                }
                return $PSObject
            }
        }
    }
    End
    {
    }
}