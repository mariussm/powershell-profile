<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Process-MIMExportErrors
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    [OutputType([int])]
    Param
    (
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=1)]
        [String] $MA = "isk.local",

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=1)]
        [String] $CSExportPath = "C:\Program Files\Microsoft Azure AD Sync\Bin\csexport.exe"
    )

    Begin
    {
    }
    Process
    {
        $Tempfile = Join-Path $env:TEMP (([guid]::newguid()).ToString() + ".xml")
        Start-Process -Wait -FilePath $CSExportPath -ArgumentList $MA, $Tempfile, "/f:e" #,"/o:e"
        
        [xml] $xml = gc $Tempfile -Encoding UTF8

        $xml.'cs-objects'.'cs-object' | foreach {
            $csobject = $_ # $csobject = $xml.'cs-objects'.'cs-object' | select -index 1
            Write-Verbose "Working on object: $($csobject.'cs-dn')"
            if($csobject.'export-errordetail'.'error-type' -eq "permission-issue") {
                if($csobject.'unapplied-export'.delta.operation -eq "update") {
                    if($csobject.'unapplied-export'.delta.attr) {
                        $csobject.'unapplied-export'.delta.attr | foreach {
                            $attribute = $_ # $attribute = $csobject.'unapplied-export'.delta.attr | select -first 1
                            Write-Verbose "Working on attribute: $($attribute.name)"
                        
                            if($attribute.multivalued -eq "true") {
                                if($attribute.operation -eq "delete") {
                                    Write-verbose "Emptying attribute $($attribute.name) of object: $($csobject.'cs-dn')"
                                    if($csobject.'object-type' -eq 'user') {
                                        Set-ADUser -Clear $attribute.name -Identity $csobject.'cs-dn' -WhatIf:([bool]$WhatIfPreference.IsPresent)
                                    } elseif($csobject.'object-type' -eq 'group') {
                                        Set-ADGroup -Clear $attribute.name -Identity $csobject.'cs-dn' -WhatIf:([bool]$WhatIfPreference.IsPresent)
                                    } else {
                                        Write-Warning "Don't know how to handle object type $($csobject.'object-type')"
                                    }
                                } else {
                                    $attribute.value | foreach {
                                        $value = $_ # $value = $attribute.value | select -first 1
                                        if($value.operation -eq "add") {
                                            Write-verbose "Adding value '$($value.'#text')' to $($attribute.name) of object: $($csobject.'cs-dn')"
                                            Set-ADObject -Identity $csobject.'cs-dn' -Add @{"$($attribute.name)" = $value."#text"} -WhatIf:([bool]$WhatIfPreference.IsPresent)
                                        } elseif($value.operation -eq $null -and $attribute.operation -eq "add") {
                                            Write-verbose "Adding value '$($value)' to $($attribute.name) of object: $($csobject.'cs-dn')"
                                            [string] $strvalue = $value
                                            Set-ADObject -Identity $csobject.'cs-dn' -Add @{"$($attribute.name)" = $strvalue} -WhatIf:([bool]$WhatIfPreference.IsPresent)
                                        } elseif($value.operation -eq "delete" -or ($value.operation -eq $null -and $attribute.operation -eq "delete")) {
                                            Write-verbose "Removing value '$($value.'#text')' from $($attribute.name) of object: $($csobject.'cs-dn')"
                                            Set-ADObject -Identity $csobject.'cs-dn' -Remove @{"$($attribute.name)" = $value."#text"} -WhatIf:([bool]$WhatIfPreference.IsPresent)
                                        } else {
                                            Write-Warning "Don't know how to handle attribute operation '$($_.operation)' for attribute $($attribute.name) for object: $($csobject.'cs-dn')"
                                        }
                                    }
                                }
                            } else {
                                Write-Warning "Have not implemented single valued attribute yet"
                            }
                        }
                    }

                    if($csobject.'unapplied-export'.delta.'dn-attr') {
                        $csobject.'unapplied-export'.delta.'dn-attr' | foreach {
                            $attribute = $_ # $attribute = $csobject.'unapplied-export'.delta.'dn-attr' | select -first 1
                            Write-Verbose "Working on attribute: $($attribute.name)"
                        
                            if($attribute.multivalued -eq "true") {
                                if($attribute.operation -eq "delete") {
                                    Write-verbose "Emptying attribute $($attribute.name) of object: $($csobject.'cs-dn')"
                                    if($csobject.'object-type' -eq 'user') {
                                        Set-ADUser -Clear $attribute.name -Identity $csobject.'cs-dn' -WhatIf:([bool]$WhatIfPreference.IsPresent)
                                    } elseif($csobject.'object-type' -eq 'group') {
                                        Set-ADGroup -Clear $attribute.name -Identity $csobject.'cs-dn' -WhatIf:([bool]$WhatIfPreference.IsPresent)
                                    } else {
                                        Write-Warning "Don't know how to handle object type $($csobject.'object-type')"
                                    }
                                } else {
                                    $attribute.'dn-value' | foreach {
                                        $value = $_ # $value = $attribute.'dn-value' | select -first 1
                                        if($value.operation -eq "add" ) {
                                            Write-verbose "Adding value '$($value.dn)' to $($attribute.name) of object: $($csobject.'cs-dn')"
                                            Set-ADObject -Identity $csobject.'cs-dn' -Add @{"$($attribute.name)" = $value.dn} -WhatIf:([bool]$WhatIfPreference.IsPresent)
                                        } elseif($value.operation -eq $null -and $attribute.operation -eq "add") {
                                            Write-verbose "Adding value '$($value)' to $($attribute.name) of object: $($csobject.'cs-dn')"
                                            [String] $strvalue = $value
                                            Set-ADObject -Identity $csobject.'cs-dn' -Add @{"$($attribute.name)" = $strvalue} -WhatIf:([bool]$WhatIfPreference.IsPresent)
                                        } elseif($value.operation -eq "delete" -or ($value.operation -eq $null -and $attribute.operation -eq "delete")) {
                                            Write-verbose "Removing value '$($value.dn)' from $($attribute.name) of object: $($csobject.'cs-dn')"
                                            Set-ADObject -Identity $csobject.'cs-dn' -Remove @{"$($attribute.name)" = $value.dn} -WhatIf:([bool]$WhatIfPreference.IsPresent)
                                        } else {
                                            Write-Warning "Don't know how to handle attribute operation '$($value.operation)' for attribute $($attribute.name) for object: $($csobject.'cs-dn')"
                                        }
                                    }
                                }
                            } else {
                                Write-Warning "Have not implemented single valued attribute yet"
                            }
                        }
                    }
                } else {
                    Write-Warning "Don't know how to handle operation '$($csobject.'unapplied-export'.delta.operation)' for object: $($csobject.'cs-dn')"
                }
            } else {
                Write-Warning "Don't know how to handle error type '$($csobject.'export-errordetail'.'error-type')' for object: $($csobject.'cs-dn')"
            } 
        }
    }
    End
    {
    }
}