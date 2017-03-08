<#
.Synopsis
   Run report from Office 365 API
.DESCRIPTION
   Wrapper method for more easily running office 365 reports
.EXAMPLE
   Invoke-Office365Report $Credential MessageTrace -Filter ("RecipientAddress eq 'user@contoso.com' and StartDate eq datetime'{0}' and EndDate eq datetime'{1}'" -f (get-date (Get-Date).AddHours(-48) -Format "yyyy-MM-ddTHH:mm:ssZ"), (get-date -Format "yyyy-MM-ddTHH:mm:ssZ") ) -Verbose -OrderBy Received
.EXAMPLE
   Invoke-Office365Report $Credential MailTraffic -Filter ("AggregateBy eq 'Day' and StartDate eq datetime'{0}' and EndDate eq datetime'{1}'" -f (get-date (get-date).AddDays(-90) -Format "yyyy-MM-ddTHH:mm:ssZ"), (get-date -Format "yyyy-MM-ddTHH:mm:ssZ") ) 
#>
function Invoke-Office365Report
{
    [CmdletBinding()]
    [OutputType([int])]
    Param
    (
        # Credentials for running report
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$false,
                   Position=0)]
        [System.Management.Automation.PSCredential]$Credential,

        # Name of the report to run
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=1)]
        [string]$WebService = "MailTraffic",

        # CSV file to write to
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=2)]
        [bool]$OutputCSV,

        # What to select
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=3)]
        [String]$Select="*",

        # Filter
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=4)]
        [String]$Filter,

        # OrderBy
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=4)]
        [String]$OrderBy,

        # Top
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=5)]
        [String]$Top

    )

    Begin
    {
        $Root = "https://reports.office365.com/ecp/reportingwebservice/reporting.svc/"
        $Format = "`$format=JSON" 
        $Select = "`$select=$Select"
        
        $url = ($Root + $WebService + "/?" + $Select + "&" + $Format) 

        if($Filter) {
            $url += "&" + "`$filter=$Filter"
        }

        if($OrderBy) {
            $url += "&" + "`$orderby=$OrderBy"
        }

        if($Top) {
            $url += "&" + "`$top=$Top"
        }

        Write-Verbose "Built url: $url"
    }
    Process
    {
        Write-Verbose "Invoking rest method"
        $rawReportData = (Invoke-RestMethod -Credential $Credential -uri $url) 
    }
    End
    {
        if($OutputCSV) {
            Write-Verbose ("Outputing csv {0}.csv" -f $WebService)
            $rawReportData.d.results | Select-Object * -ExcludeProperty __metadata | Export-Csv -Path ("{0}.csv" -f $WebService) -NoTypeInformation
            return ("{0}.csv" -f $WebService)
        } else {
            Write-Verbose "Returning results"
            return ($rawReportData.d.results | Select-Object * -ExcludeProperty __metadata)
        }
    }
} 