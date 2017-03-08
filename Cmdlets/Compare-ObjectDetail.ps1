Function Compare-ObjectDetail {
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$false,
                   Position=0)]
        $ReferenceObject,

        # Param2 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$false,
                   Position=1)]
        $DifferenceObject
    )

    $ReferenceObject | Get-Member -MemberType Property -ErrorAction SilentlyContinue | foreach {
        $attribute = $_.Name
        $comp = Compare-Object -ReferenceObject $ReferenceObject -DifferenceObject $DifferenceObject -Property $attribute
        
        
        if($comp) {
            $comp | foreach {
                New-Object -TypeName PSObject -Property @{
                    Attribute = $attribute
                    Value = $_.$attribute
                    SideIndicator = $_.SideIndicator
                }
            }
        }
    }

}
