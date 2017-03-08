<#
.Synopsis
   Displays the AD thumbnailphoto
.DESCRIPTION
   Displays the AD thumbnailphoto
.EXAMPLE
   Show-ADThumbnailPhoto masol
#>
function Show-ADThumbnailPhoto
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$false,
                   Position=0)]
        $SamAccountName
    )

    Begin
    {
        Import-Module ActiveDirectory
    }
    Process
    {
        $aduser = Get-ADUser -Identity $SamAccountName -Properties thumbnailPhoto
        if($aduser.thumbnailPhoto) 
        {
            $aduser.thumbnailPhoto | Set-Content -Path "$($env:TEMP)\adphoto.png" -Encoding Byte
            ii "$($env:TEMP)\adphoto.png"
        } 
        else
        {
            Write-Error "User $SamAccountName has no photo"
        }
    }
    End
    {
    }
}