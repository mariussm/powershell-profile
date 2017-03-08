<#
.Synopsis
   Returns the groups that a user is member of, by default recusive
.DESCRIPTION
   Returns the groups that a user is member of, by default recusive
.EXAMPLE
   Get-ADUser masol | Get-ADUserGroups
.EXAMPLE
   "masol","admmasol" | Get-ADUser | Get-ADUserGroups -Properties mail -Recursive:$false
#>
function Get-ADUserGroups
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=2)]
        $ADUser,

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=0)]
        [boolean] $Recursive = $true,

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=1)]
        [string[]] $Properties = @("sAMAccountName")
    )

    Begin
    {
    }
    Process
    {
        if($Recursive) {
            Get-ADGroup -LDAPFilter ("(member:1.2.840.113556.1.4.1941:={0})" -f $ADUser.DistinguishedName) -Properties $Properties
        } else {
            Get-ADGroup -LDAPFilter ("(member={0})" -f $ADUser.DistinguishedName) -Properties $Properties
        }
    }
    End
    {
    }
}
