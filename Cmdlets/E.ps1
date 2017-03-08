Function E {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   Position=0)]
        [String] $Attribute,

        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=1)]
        $Object
    )

    Begin
    {
        $DetectedAttribute = $null
    }
    Process
    {
        if($DetectedAttribute) {
            $Object.$DetectedAttribute
        } elseif ($Object.$Attribute) {
            $Object.$Attribute
        } elseif (!$Attribute.Contains("*")) {
            
        } else {
            $AttributeObject = $Object | gm -MemberType CodeProperty, NoteProperty, ScriptProperty, Property | ? Name -like $Attribute | Select -First 1
            if($AttributeObject) {
                $DetectedAttribute = $AttributeObject.Name
                Write-Verbose "DetectedAttribute is $DetectedAttribute"
                $Object.$DetectedAttribute
            } else {
                Write-Error "No attribute matching '$Attribute'"
            }
        }
    }
    End
    {
    }

}
