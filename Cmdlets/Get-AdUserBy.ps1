<#
.Synopsis
   Runs Get-ADUser with filter based on the Attribute-parameter
.DESCRIPTION
   Runs Get-ADUser with filter based on the Attribute-parameter
.EXAMPLE
   Get-ADUserBy proxyAddresses marius.solbakken@goodworkaround.com
#>
function Get-ADUserBy
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$True,Position=0,ValueFromPipeline=$false)]
        [ValidateSet('proxyAddresses','mail','userPrincipalName','targetAddress','displayName','sAMAccountName','employeeid')]
        [string]$Attribute = "proxyAddresses",

        [Parameter(Mandatory=$True,Position=1,ValueFromPipeline=$true)]
        [string]$Value,

        [Parameter(Mandatory=$False,Position=2,ValueFromPipeline=$false)]
        [string[]]$Properties = @("targetaddress";"proxyaddresses";"mail";"lastLogonDate";"displayname"),

        [Parameter(Mandatory=$False,Position=3,ValueFromPipeline=$false)]
        [string]$Server,

        [Parameter(Mandatory=$False,Position=4,ValueFromPipeline=$false)]
        [alias("exp")]
        [string]$ExpandProperty
    )

    Begin
    {
        $baseparams = @{}
        if($Properties) 
        {
            $baseparams["Properties"] = $Properties
        }
        if($Server) 
        {
            $baseparams["Server"] = $Server
        }
    }
    Process
    {
        if($Value -notlike "*:*" -and $Attribute -eq "proxyAddresses") {
            $Value = "SMTP:$Value"
        }

        if($ExpandProperty) {
            $baseparams["Properties"] = $ExpandProperty
            Get-ADUser -Filter {$Attribute -like $Value} @baseparams | Select-Object -ExpandProperty $ExpandProperty
        } else {
            Get-ADUser -Filter {$Attribute -like $Value} @baseparams
        }
    }
    End
    {
    }
}