<#
.Synopsis
   Deletes MIM Object
.DESCRIPTION
   Deletes MIM Object
.EXAMPLE
   Remove-MIMObject "65ca7eff-75ae-4d68-b026-df05f609133e"
#>
function Remove-MIMObject
{
    [cmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='High')]

    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        $ObjectID,

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=0)]
        [String] $ObjectType = "Person"
    )

    Process
    {
        $ObjectID = $ObjectID.Replace("urn:uuid:","")
        if($PSCmdlet.ShouldProcess($env:COMPUTERNAME,"Deleting object $ObjectID")) {
            New-FimImportObject -State Delete -ObjectType $ObjectType -AnchorPairs @{ObjectID = $objectID} -ApplyNow:$true
        }
    }
}