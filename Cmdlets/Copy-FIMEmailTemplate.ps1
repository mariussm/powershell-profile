<#
.Synopsis
   Creates a copy of a FIM email template
.DESCRIPTION
   Creates a copy of a FIM email template
.EXAMPLE
   Get-FIMObjectByXPath '/EmailTemplate' | where{$_.DisplayName -like "- Test user*"} | Copy-FIMEmailTemplate
#>
function Copy-FIMEmailTemplate
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $Template
    )

    Begin
    {
    }
    Process
    {
        if($Template.ObjectType -eq "EmailTemplate") {
            $changes = @{
                DisplayName = ("{0}{1}" -f "__COPY: ", $Template.DisplayName)
                EmailBody = $Template.EmailBody
                EmailSubject = $Template.EmailSubject
                EmailTemplateType = $Template.EmailTemplateType
            }

            New-FimImportObject -ObjectType EmailTemplate -State Create -ApplyNow -PassThru -Changes $changes
        } else {
            Write-Error "Invalid input object"
        }
    }
    End
    {
    }
}