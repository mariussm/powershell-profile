
function Get-PrettyPrintedXML
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([string])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $InputString,

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=1)]
        [ValidateSet("Base64","UrlDecodeBeforeBase64","Plain")]
        [string] $Type = "Plain"
    )

    Begin
    {
    }
    Process
    {
        if($Type -eq "UrlDecodeBeforeBase64") {
            $InputString = [System.Web.HttpUtility]::UrlDecode($InputString)
        }
        
        if($Type -in "UrlDecodeBeforeBase64","Base64") {
            $InputString = [System.Text.Encoding]::UTF8.GetString(([System.Convert]::FromBase64String($InputString)))
        }

        $doc = New-Object System.Xml.XmlDataDocument
        $doc.LoadXml($InputString)
        $sw=New-Object System.Io.Stringwriter
        $writer=New-Object System.Xml.XmlTextWriter($sw)
        $writer.Formatting = [System.Xml.Formatting]::Indented
        $doc.WriteContentTo($writer)
        $sw.ToString()
    }
    End
    {
    }
}