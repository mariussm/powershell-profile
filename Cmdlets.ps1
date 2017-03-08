Function Add-CodeSignature {
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=1)]
        $ThumbPrint,

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   Position=0)]
        [String[]] $Files
    )

    Begin
    {
    }
    Process
    {
        if($ThumbPrint)
        {
            $Certificate = Get-CodeSigningCertificate | Where{$_.Thumbprint -eq $Thumbprint}
        }
        else
        {
            $Certificates = Get-CodeSigningCertificate
            $Certificate = $Certificates | Select -First 1
            if(($Certificates | measure).Count -gt 1)
            {
                Write-Warning "Warning, there are multiple signing certificates."
            }
        }

        if(!$Certificate)
        {
            throw New-Object System.Exception("Cannot find code signing certificate with thumbprint $Thumbprint")
        }

        Set-AuthenticodeSignature -Certificate $Certificate -FilePath $Files

        if($psise) {
            $FullFileNames = $Files | dir | select -exp FullName
            $psise.CurrentPowerShellTab.Files | Where{$Fullpath = $_.FullPath; $FullFileNames | Where{$_ -eq $FullPath}} | Foreach {
                if(!$_.IsSaved){
                    $_.Save()
                }
                [int] $CaretLine = $_.Editor.CaretLine
                [int] $CaretColumn = $_.Editor.CaretColumn
                $_.Editor.Text = Get-Content -Raw $_.FullPath
                $_.Editor.SetCaretPosition($CaretLine, $CaretColumn)
                $_.Save()
                
            }
        }
    }
    End
    {
    }

}

Function Add-CodeSignatureToCurrentISEFile {
    [CmdletBinding()]
    [Alias()]
    Param
    (
    )

    Begin
    {
    }
    Process
    {
        if($psise.CurrentFile)
        {
            Add-CodeSignature -Files $psise.CurrentFile.FullPath
        }
    }
    End
    {
    }

}

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

function Connect-ExchangeOnline{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True,Position=0)]
        [System.Management.Automation.PSCredential]$Credentials
    )
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Authentication Basic -AllowRedirection -Credential $Credentials
    Import-PSSession $session -DisableNameChecking
}
Function ConvertFrom-Base64 {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True,
                   Position=0,
                   ValueFromPipeline=$true)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string]
        $Base64String
    )

    Begin{}
    Process{
        return [System.Text.Encoding]::UTF8.GetString(([System.Convert]::FromBase64String($Base64String)));
    }
    End{}

}

<#
.Synopsis
   Converts immutableID in Office 365 to GUID
.DESCRIPTION
   Converts immutableID in Office 365 to GUID
.EXAMPLE
   Get-MsolUser -UserPrincipalName marius@goodworkaround.com | Select -ExpandProperty ImmutableID | ConvertFrom-ImmutableID
#>
function ConvertFrom-ImmutableID
{
    [CmdletBinding()]
    [OutputType([GUID])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $ImmutableID
    )

    Process 
    {
        return [guid]([system.convert]::frombase64string($ImmutableID) )
    }
}
<#
.Synopsis
   Converts a filetime to datetime. Can be used on lastLogonTimestamp in AD.
.DESCRIPTION
   Converts a filetime to datetime. Can be used on lastLogonTimestamp in AD.
.EXAMPLE
   Get-ADUser masol -property lastLogonTimestamp | Select-Object -ExpandProperty lastLogonTimestamp | ConvertFrom-LastLogonTimestamp
.EXAMPLE
   ConvertFrom-LastLogonTimestamp 129948127853609000
#>
function ConvertFrom-LastLogonTimestamp
{
    [CmdletBinding()]
    [OutputType([datetime])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $LastLogonTimestamp
    )

    return [datetime]::FromFileTime($LastLogonTimestamp)
}
Function ConvertFrom-SecureStringToString {
    [CmdletBinding()]
    [OutputType([System.String])]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        [System.Security.SecureString] $SecureString
    )

    Begin
    {
    }
    Process
    {
        [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
        )
    }
    End
    {
    }

}

Function ConvertTo-Base64 {
    [CmdletBinding(DefaultParameterSetName='String')]
    [OutputType([String])]
    Param
    (
        # String to convert to base64
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='String')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string]
        $String,

        # Param2 help description
        [Parameter(ParameterSetName='ByteArray')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [byte[]]
        $ByteArray
    )

    Begin{}
    Process{
        if($String) {
            return [System.Convert]::ToBase64String(([System.Text.Encoding]::UTF8.GetBytes($String)));
        } else {
            return [System.Convert]::ToBase64String($ByteArray);
        }
    }
    End{}

}

Function ConvertTo-Bytes {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True,Position=0,ValueFromPipeline=$true)]
        [string]$ByteQuantifiedSize
    )

    return [long] ([string] $ByteQuantifiedSize).Split("(")[1].Split(" ")[0].Replace(",","")

}

<#
.Synopsis
   Converts GUID in AD to ImmutableID
.DESCRIPTION
   Converts GUID in AD to ImmutableID
.EXAMPLE
   GetADUser | Select -ExpandProperty ImmutableID | ConvertFrom-ImmutableID
#>
function ConvertTo-ImmutableID
{
    [CmdletBinding()]
    [OutputType([GUID])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [GUID] $ObjectGUID
    )

    Process 
    {
        return [system.convert]::ToBase64String($ObjectGUID.ToByteArray())
    }
}

Function ConvertTo-MB {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True,Position=0,ValueFromPipeline=$true)]
        [string]$ByteQuantifiedSize
    )

    Begin {}
    Process {
        return (ConvertTo-Bytes $ByteQuantifiedSize) / 1024 / 1024
    }
    End{}

}

Function ConvertTo-Mbps {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True,Position=0,ValueFromPipeline=$true)]
        [string]$MegaBytesPerMinute
    )

    Begin{}
    Process{
        return ($MegaBytesPerMinute / 60 * 8)
    }
    End{}

}

<#
.Synopsis
   Copies all claim rules from one RPT to another
.DESCRIPTION
   Copies all claim rules from one RPT to another
.EXAMPLE
   Copy-ADFSClaimRules -SourceRelyingPartyTrustName "Office 365" -DestinationRelyingPartyTrustName "Token testing website - Marius"
#>
function Copy-ADFSClaimRules
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   Position=0)]
        [string] $SourceRelyingPartyTrustName,

        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   Position=1)]
        [string] $DestinationRelyingPartyTrustName
    )

    Begin
    {
    }
    Process
    {
        $SourceRPT = Get-AdfsRelyingPartyTrust -Name $SourceRelyingPartyTrustName
        $DestinationRPT = Get-AdfsRelyingPartyTrust -Name $DestinationRelyingPartyTrustName

        if(!$SourceRPT) {
            Write-Error "Could not find $SourceRelyingPartyTrustName"
        } elseif(!$DestinationRPT) {
            Write-Error "Could not find $DestinationRelyingPartyTrustName"
        }

        Set-AdfsRelyingPartyTrust -TargetRelyingParty $DestinationRPT -IssuanceTransformRules $SourceRPT.IssuanceTransformRules -IssuanceAuthorizationRules $SourceRPT.IssuanceAuthorizationRules -DelegationAuthorizationRules $SourceRpT.DelegationAuthorizationRules
    }
    End
    {
    }
}
<#
.Synopsis
   Copies relying party trust
.DESCRIPTION
   Copies relying party trust
.EXAMPLE
   Copy-AdfsRelyingPartyTrust "Intranett Test" "Intranett Prod" "urn:sharepoint:prod"
#>
function Copy-AdfsRelyingPartyTrust
{
    [CmdletBinding()]
    [OutputType([int])]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   Position=0)]
        $SourceRelyingPartyTrustName,

        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   Position=1)]
        $NewRelyingPartyTrustName,

        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   Position=2)]
        $NewRelyingPartyTrustIdentifier
    )

    Begin
    {
    }
    Process
    {
        $SourceRelyingPartyTrust  = Get-AdfsRelyingPartyTrust -Name $SourceRelyingPartyTrustName

        $exceptedAttributes = @("ConflictWithPublishedPolicy","OrganizationInfo","ProxyEndpointMappings","LastUpdateTime","PublishedThroughProxy","LastMonitoredTime")
        $parameters = @{}
        $SourceRelyingPartyTrust | Get-Member -MemberType Property | where{$_.name -notin $exceptedAttributes} | foreach {
            if($SourceRelyingPartyTrust.($_.Name) -ne $null) {
                $parameters[$_.Name] = $SourceRelyingPartyTrust.($_.Name)
            }
        }
        $parameters.Name = $NewRelyingPartyTrustName
        $parameters.Identifier = $NewRelyingPartyTrustIdentifier
        
        Add-AdfsRelyingPartyTrust @parameters
    }
    End
    {
    }
}
Function Copy-AttributeToClipboard {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=1)]
        $Object,

        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   Position=0)]
        [String] $Attribute
    )

    Begin
    {
    }
    Process
    {
        $Object | select -ExpandProperty $Attribute | clip
    }
    End
    {
    }

}

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
<#
.Synopsis
    Creates a copy of the input set(s)
.DESCRIPTION
    Creates a copy of the input set(s)
.EXAMPLE
    Get-FIMObjectByXPath '/Set[DisplayName="All People"]' | Copy-FIMSet
#>
function Copy-FIMSet
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
            ValueFromPipeline=$true,
            Position=0)]
        $Set,

        [Parameter(Mandatory=$false,
            ValueFromPipeline=$false,
            Position=1)]
        $Prefix = "- [COPY] "
    )
    Begin
    {
    }
    Process
    {
        if($Set.DisplayName -and $Set.Filter -and $Set.ObjectType -eq "Set") {
            $changes = @{
                DisplayName = ("{0}{1}" -f $Prefix, $Set.DisplayName)
                Filter = $Set.Filter
            }
            New-FimImportObject -ObjectType Set -ApplyNow -PassThru -State Create -Changes $changes
        } else 
        {
            Write-Error "Input object not valid"
        }
    }
    End
    {
    }
}
<#
.Synopsis
    Creates a copy of a set
.DESCRIPTION
    Creates a copy of a set
.EXAMPLE
    Copy-FIMSetByName "All People" "All People 2"
#>
function Copy-FIMSetByName
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
            ValueFromPipeline=$false,
            Position=0)]
        $Source,

        [Parameter(Mandatory=$false,
            ValueFromPipeline=$false,
            Position=1)]
        $Destination
    )
    Begin
    {
    }
    Process
    {
        $SourceSet = Get-FIMObjectByXPath "/Set[DisplayName=""$Source""]"
        if(!$SourceSet) {
            Write-Error "Set not found"
        } else {
            $changes = @{
                DisplayName = $Destination
                Filter = $SourceSet.Filter
            }
            New-FimImportObject -ObjectType Set -ApplyNow -PassThru -State Create -Changes $changes
        }
    }
    End
    {
    }
}
Function Disable-ADAL {
    [CmdletBinding()]
    Param
    ()

    Set-ItemProperty -Path HKCU:\Software\Microsoft\Office\15.0\Common\Identity\ -Name EnableADAL -Type DWord -Value 0

}

Function Disable-ScheduledTask2008 {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True,Position=0)]
        [string]$Name
    )

    $ret = schtasks.exe /Change /DISABLE /TN "$Name"

    return ($ret -like "SUCCESS:*") -eq $true

}

function Disconnect-ExchangeOnline {
    [CmdletBinding()]
    Param()
    Get-PSSession | ?{$_.ComputerName -like "*outlook.com"} | Remove-PSSession
}
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

Function Enable-ADAL {
    [CmdletBinding()]
    Param
    ()

    mkdir HKCU:\Software\Microsoft\Office\15.0\Common\Identity\ -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
    Set-ItemProperty -Path HKCU:\Software\Microsoft\Office\15.0\Common\Identity\ -Name EnableADAL -Type DWord -Value 1
    mkdir HKCU:\Software\Microsoft\Office\15.0\Common\Debug -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
    Set-ItemProperty -Path HKCU:\Software\Microsoft\Office\15.0\Common\Debug -Name TCOTrace -Type DWord -Value 3

}

Function Enable-ScheduledTask2008 {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True,Position=0)]
        [string]$Name
    )

    $ret = schtasks.exe /Change /ENABLE /TN "$Name"

    return ($ret -like "SUCCESS:*") -eq $true

}

Function End-ScheduledTask2008 {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True,Position=0)]
        [string]$Name
    )

    $ret = schtasks.exe /End /TN "$Name"

    return ($ret -like "SUCCESS:*") -eq $true

}

<#
.Synopsis
   Returns the ADFS token signing and encryption certificates
.DESCRIPTION
   Returns the ADFS token signing and encryption certificates
.EXAMPLE
   Get-AdfsCertificates adfs.goodworkaround.com
#>
function Get-AdfsCertificates
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$false,
                   Position=0)]
        $ADFS
    )

    Begin
    {
    }
    Process
    {
        $metadata = Invoke-RestMethod -Uri ("https://{0}/FederationMetadata/2007-06/FederationMetadata.xml" -f $ADFS)

        $metadata.EntityDescriptor.RoleDescriptor.KeyDescriptor | foreach {
            $tempfile = "{0}\adfsTempCert.cer" -f $env:temp
            $_.KeyInfo.X509Data.X509Certificate | Set-Content -Path $tempfile

            $cert = (New-Object System.Security.Cryptography.X509Certificates.X509Certificate2)
            $cert.Import($tempfile)

            New-Object -TypeName PSCustomObject -Property @{
                FoundIn = "KeyDescriptor"
                Use = $_.Use
                Subject = $cert.Subject
                Issuer = $cert.Issuer
                ThumbPrint = $cert.Thumbprint
                NotBefore = $cert.NotBefore
                NotAfter = $cert.NotAfter
                Data = $_.KeyInfo.X509Data.X509Certificate
            }
        }

        $tempfile = "{0}\adfsTempCert.cer" -f $env:temp
        $metadata.EntityDescriptor.Signature.KeyInfo.X509Data.X509Certificate | Set-Content -Path $tempfile
        $cert = (New-Object System.Security.Cryptography.X509Certificates.X509Certificate2)
        $cert.Import($tempfile)

        New-Object -TypeName PSCustomObject -Property @{
            FoundIn = "Active Signature"
            Use = "signing"
            Subject = $cert.Subject
            Issuer = $cert.Issuer
            ThumbPrint = $cert.Thumbprint
            NotBefore = $cert.NotBefore
            NotAfter = $cert.NotAfter
            Data = $metadata.EntityDescriptor.Signature.KeyInfo.X509Data.X509Certificate
        }
    }
    End
    {
    }
}
<#
.Synopsis
   Returns the thumbprint of the ADFS token signing certificate
.DESCRIPTION
   Returns the thumbprint of the ADFS token signing certificate
.EXAMPLE
   Get-AdfsTokenSigningThumbprint adfs.goodworkaround.com
#>
function Get-AdfsTokenSigningThumbprint
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$false,
                   Position=0)]
        $ADFS
    )

    Begin
    {
    }
    Process
    {
        $metadata = Invoke-RestMethod -Uri ("https://{0}/FederationMetadata/2007-06/FederationMetadata.xml" -f $ADFS)
        $tempfile = "{0}\adfsTempCert.cer" -f $env:temp
        $metadata.EntityDescriptor.Signature.KeyInfo.X509Data.X509Certificate | Set-Content -Path $tempfile
        $cert = (New-Object System.Security.Cryptography.X509Certificates.X509Certificate2)
        $cert.Import($tempfile)

        return $cert.Thumbprint
    }
    End
    {
    }
}
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

Function Get-CodeSigningCertificate {
    [CmdletBinding()]
    [Alias()]
    [OutputType([System.Security.Cryptography.X509Certificates.X509Certificate])]
    Param
    (
    )

    Begin
    {
    }
    Process
    {
        get-childitem Cert:\CurrentUser\my -CodeSigningCert
    }
    End
    {
    }

}

Function Get-CoffeeWaterAmount {
    [OutputType([int])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        [int] $GramCoffee
    )

    Begin
    {
    }
    Process
    {
        return [Math]::Floor(1000 / 65 * $GramCoffee)
    }
    End
    {
    }

}

Function Get-ContainerNameFromDistinguishedName {
    [CmdletBinding()]
    [OutputType([string])]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $DN
    )

    Begin
    {
    }
    Process
    {
        $DN -split "[^\\],", 2 -split "=" | select -index 1
    }
    End
    {
    }

}

Function Get-ContentAsString {
    [CmdletBinding()]
    [Alias()]
    [OutputType([string])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Path
    )

    Begin
    {
    }
    Process
    {
        return [IO.File]::ReadAllText((dir ($Path)).Fullname)
    }
    End
    {
    }

}

Function Get-DecryptedValueOfBase64String {
    [CmdletBinding()]
    [Alias()]
    [OutputType([string])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [string] $InputString,

        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$false,
                   Position=0)]
        [String] $Thumbprint
    )

    Begin
    {
        $Cert = ((dir Cert:\LocalMachine\my) | ?{$_.PrivateKey.KeyExchangeAlgorithm -and $_.Verify()}) , ((dir Cert:\CurrentUser\my) | ?{$_.PrivateKey.KeyExchangeAlgorithm -and $_.Verify()}) | Where{$_.Thumbprint -eq $Thumbprint}
        if(!$Cert) {
            throw "No certificate with thumbprint $Thumbprint found"
        }
    }
    Process
    {
        $EncryptedBytes = [System.Convert]::FromBase64String($InputString)
        $DecryptedBytes = $Cert.PrivateKey.Decrypt($EncryptedBytes, $true)
        return [system.text.encoding]::UTF8.GetString($DecryptedBytes)
    }
    End
    {
    }

}

Function Get-DesktopPath {
    [CmdletBinding()]
    [OutputType([string])]
    Param
    (
    )

    Begin
    {
    }
    Process
    {
        return [Environment]::GetFolderPath("Desktop")
    }
    End
    {
    }

}

Function Get-DnsAddressList {
    param(
        [parameter(Mandatory=$true)][Alias("Host")]
          [string]$HostName)

    try {
        return [System.Net.Dns]::GetHostEntry($HostName).AddressList
    }
    catch [System.Net.Sockets.SocketException] {
        if ($_.Exception.ErrorCode -ne 11001) {
            throw $_
        }
        return = @()
    }

}

Function Get-DnsMXQuery {
    param(
        [parameter(Mandatory=$true)]
          [string]$DomainName)

    if (-not $Script:global_dnsquery) {
        $Private:SourceCS = @'
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace PM.Dns {
  public class MXQuery {
    [DllImport("dnsapi", EntryPoint="DnsQuery_W", CharSet=CharSet.Unicode, SetLastError=true, ExactSpelling=true)]
    private static extern int DnsQuery(
        [MarshalAs(UnmanagedType.VBByRefStr)]
        ref string pszName, 
        ushort     wType, 
        uint       options, 
        IntPtr     aipServers, 
        ref IntPtr ppQueryResults, 
        IntPtr pReserved);

    [DllImport("dnsapi", CharSet=CharSet.Auto, SetLastError=true)]
    private static extern void DnsRecordListFree(IntPtr pRecordList, int FreeType);

    public static string[] Resolve(string domain)
    {
        if (Environment.OSVersion.Platform != PlatformID.Win32NT)
            throw new NotSupportedException();

        List<string> list = new List<string>();

        IntPtr ptr1 = IntPtr.Zero;
        IntPtr ptr2 = IntPtr.Zero;
        int num1 = DnsQuery(ref domain, 15, 0, IntPtr.Zero, ref ptr1, IntPtr.Zero);
        if (num1 != 0)
            throw new Win32Exception(num1);
        try {
            MXRecord recMx;
            for (ptr2 = ptr1; !ptr2.Equals(IntPtr.Zero); ptr2 = recMx.pNext) {
                recMx = (MXRecord)Marshal.PtrToStructure(ptr2, typeof(MXRecord));
                if (recMx.wType == 15)
                    list.Add(Marshal.PtrToStringAuto(recMx.pNameExchange));
            }
        }
        finally {
            DnsRecordListFree(ptr1, 0);
        }

        return list.ToArray();
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MXRecord
    {
        public IntPtr pNext;
        public string pName;
        public short  wType;
        public short  wDataLength;
        public int    flags;
        public int    dwTtl;
        public int    dwReserved;
        public IntPtr pNameExchange;
        public short  wPreference;
        public short  Pad;
    }
  }
}
'@

        Add-Type -TypeDefinition $Private:SourceCS -ErrorAction Stop
        $Script:global_dnsquery = $true
    }

    [PM.Dns.MXQuery]::Resolve($DomainName) | % {
        $rec = New-Object PSObject
        Add-Member -InputObject $rec -MemberType NoteProperty -Name "Host"        -Value $_
        Add-Member -InputObject $rec -MemberType NoteProperty -Name "AddressList" -Value $(Get-DnsAddressList $_)
        $rec
    }

}

Function Get-EncryptedBase64ValueOfString {
    [CmdletBinding()]
    [Alias()]
    [OutputType([string])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [string] $InputString,

        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$false,
                   Position=0)]
        [String] $Thumbprint
    )

    Begin
    {
        $Cert = ((dir Cert:\LocalMachine\my) | ?{$_.PrivateKey.KeyExchangeAlgorithm -and $_.Verify()}) , ((dir Cert:\CurrentUser\my) | ?{$_.PrivateKey.KeyExchangeAlgorithm -and $_.Verify()}) | Where{$_.Thumbprint -eq $Thumbprint}
        if(!$Cert) {
            throw "No certificate with thumbprint $Thumbprint found"
        }
    }
    Process
    {
        $EncodedInputString = [system.text.encoding]::UTF8.GetBytes($InputString)
        $EncryptedBytes = $Cert.PublicKey.Key.Encrypt($EncodedInputString, $true)
        return [System.Convert]::ToBase64String($EncryptedBytes)

    }
    End
    {
    }

}

<#
.Synopsis
    Returns the federation metadata as XML
.DESCRIPTION
    Returns the federation metadata as XML
.EXAMPLE
    Get-FederationMetadata "adfs.goodworkaround.com"
#>
function Get-FederationMetadata
{
    [CmdletBinding()]
    [OutputType([xml])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                ValueFromPipeline=$true,
                Position=0)]
        $FQDN
    )
 
    Begin
    {
    }
    Process
    {
    return Invoke-RestMethod -Uri ("https://{0}/FederationMetadata/2007-06/FederationMetadata.xml" -f $FQDN)
    }
    End
    {
    }
}
<#
.Synopsis
    Returns the federation metadata URL
.DESCRIPTION
    Returns the federation metadata URL
.EXAMPLE
    Get-FederationMetadataURL "adfs.goodworkaround.com"
#>
function Get-FederationMetadataURL
{
    [CmdletBinding()]
    [OutputType([xml])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                ValueFromPipeline=$true,
                Position=0)]
        $FQDN
    )
 
    Begin
    {
    }
    Process
    {
    return ("https://{0}/FederationMetadata/2007-06/FederationMetadata.xml" -f $FQDN)
    }
    End
    {
    }
}
Function Get-FileFromURI {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        [String] $URI,

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=1)]
        [String] $DestinationFileName
    )

    Begin
    {
    }
    Process
    {
        $_DestinationFileName = $DestinationFileName

        $wc = New-Object System.Net.WebClient
        if(!$_DestinationFileName) {
            $tempURI = $URI -replace "http://",""
            $_DestinationFileName = (Split-Path -Leaf $tempURI) -replace "%20"," "
            Write-Verbose "Setting destination file name to: $_DestinationFileName"
        }

        if($_DestinationFileName.Substring(1,1) -ne ":") {
            $_DestinationFileName = (pwd).Path + "\" + $_DestinationFileName
            Write-Verbose "Full path: $_DestinationFileName"
        }

        Write-Verbose "Downloading $uri -> $_DestinationFileName"
        $wc.DownloadFile($uri, $_DestinationFileName)
    }
    End
    {
    }

}

<#
.Synopsis
   Returns fim management agent matching pattern
.DESCRIPTION
   This method uses WMI to get and return FIM Management Agents
.EXAMPLE
   Get-FIMManagementAgent "SP - *"
#>
function Get-FIMManagementAgent
{
    [CmdletBinding()]
    [OutputType([object])]
    Param
    (
        # The Management Agent name
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$false,
                   Position=0)]
        [string] $MA
    )

    Begin
    {
        # Connect to database
        Write-Verbose ("Connecting to WMI root/MicrosoftIdentityIntegrationServer class MIIS_ManagementAgent")
        $wmi = Get-WmiObject -Namespace root/MicrosoftIdentityIntegrationServer -Class MIIS_ManagementAgent
    }
    Process
    {
        return ($wmi | where{$_.Name -like $MA})
    }
    End
    {
        
    }
}
<#
.Synopsis
   Returns fim management agent run status for all, one or some MAs
.DESCRIPTION
   Returns fim management agent run status for all, one or some MAs
.EXAMPLE
   Get-FIMManagementAgentRunStatus "SP - *"
#>
function Get-FIMManagementAgentRunStatus
{
    [CmdletBinding()]
    [OutputType([object])]
    Param
    (
        # The Management Agent name
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=0)]
        [string] $MA
    )

    Begin
    {
    }
    Process
    {
        if($MA) 
        {
            $MAs = Get-FIMManagementAgent -MA $MA
        }   
        else 
        {
            $MAs = Get-FIMManagementAgent -MA *
        }
        
        return ($MAs | foreach{New-Object -TypeName PSObject -Property @{ManagementAgent=$_.Name;RunStatus=$_.RunStatus().ReturnValue}})
    }
    End
    {   
    }
}
<#
.Synopsis
   Returns all MPRs that triggers an action workflow
.DESCRIPTION
   Returns all MPRs that triggers an action workflow
.EXAMPLE
   Get-FimWorkflow *accountname* | Get-FIMManagementPolicyRuleByActionWorkflowDefinition
#>
function Get-FIMManagementPolicyRuleByActionWorkflowDefinition
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $FIMWorkflow
    )

    Begin
    {
    }
    Process
    {
        return (Export-FimConfig -CustomConfig ("/ManagementPolicyRule[ActionWorkflowDefinition='$($FIMWorkflow.ObjectID.Replace('urn:uuid:',''))']") -OnlyBaseResources -Uri "http://localhost:5725/" | Convert-FimExportToPSObject)
    }
    End
    {
    }
}
<#
.Synopsis
   Returns object with object id
.DESCRIPTION
   Returns object with object id
.EXAMPLE
   Get-FIMObjectByObjectID "0a0b2dsa-ccccc-cccc-cccccccccccc"
#>
function Get-FIMObjectByObjectID
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        [String] $ObjectID
    )

    Begin
    {
    }
    Process
    {
        $ObjectID = $ObjectID.Replace("urn:uuid:","")
        return (Export-FimConfig -CustomConfig ("/*[ObjectID='$($ObjectID)']") -OnlyBaseResources -Uri "http://localhost:5725/" | Convert-FimExportToPSObject)
    }
    End
    {
    }
}
<#
.Synopsis
   Returns objects matching xpath
.DESCRIPTION
   Returns objects matching xpath
.EXAMPLE
   Get-FIMObjectByXPath "/testUser"
#>
function Get-FIMObjectByXPath
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $XPath
    )

    Begin
    {
    }
    Process
    {
        return (Export-FimConfig -CustomConfig $XPath -OnlyBaseResources -Uri "http://localhost:5725/" | Convert-FimExportToPSObject)
    }
    End
    {
    }
}
<#
.Synopsis
   Returns objects matching xpath
.DESCRIPTION
   Returns objects matching xpath
.EXAMPLE
   Get-FIMObjectByXPath "/testUser"
#>
function Get-FIMObjectWithLocalesByXPath
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $XPath
    )

    Begin
    {
    }
    Process
    {
        $res = Export-FimConfig -CustomConfig $XPath -Uri "http://localhost:5725/" -AllLocales
        if($res) {
            $res | Foreach {
                $PSObject = $_ | Convert-FimExportToPSObject
                $PSObject | Add-Member -MemberType NoteProperty -Name Locale -Value @{}
                if($_.ResourceManagementObject.LocalizedResourceManagementAttributes) {
                    $_.ResourceManagementObject.LocalizedResourceManagementAttributes | foreach {
                        $culture = $_.Culture
                    
                        $attributes = @{}
                        $_.ResourceManagementAttributes | Foreach {
                            if($_.IsMultiValue) {
                                $attributes[$_.AttributeName] = $_.Values
                            } else {
                                $attributes[$_.AttributeName] = $_.Value
                            }
                        }
                        $PSObject.Locale[$culture] = New-Object -TypeName PSObject -Property $attributes
                    }
                }
                return $PSObject
            }
        }
    }
    End
    {
    }
}
<#
.Synopsis
   Returns fim run history for all or one MA
.DESCRIPTION
   Returns fim run history for all or one MA
.EXAMPLE
   Get-FIMRunHistory "SharePoint Internal"
#>
function Get-FIMRunHistory
{
    [CmdletBinding()]
    [OutputType([object])]
    Param
    (
        # The Management Agent name
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=0)]
        [string] $MA,

        # Return only first match
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=1)]
        [bool] $FirstOnly = $true
    )

    Begin
    {
    }
    Process
    {
        if($MA) 
        {
            if($FirstOnly) 
            {
                return Get-WmiObject -Namespace root/MicrosoftIdentityIntegrationServer -Class MIIS_RunHistory -Filter ("MaName='{0}'" -f $MA) | select -First 1
            } else 
            {
                return Get-WmiObject -Namespace root/MicrosoftIdentityIntegrationServer -Class MIIS_RunHistory -Filter ("MaName='{0}'" -f $MA)
            }
        }
        else 
        {
            if($FirstOnly)
            {
                return Get-WmiObject -Namespace root/MicrosoftIdentityIntegrationServer -Class MIIS_RunHistory | select -First 1
            } 
            else 
            {
                return Get-WmiObject -Namespace root/MicrosoftIdentityIntegrationServer -Class MIIS_RunHistory
            }
        }
        return ($wmi | where{$_.Name -like $MA})
    }
    End
    {   
    }
}
<#
.Synopsis
   Returns all Fim workflows matching pattern
.DESCRIPTION
   Returns all Fim workflows matching pattern
.EXAMPLE
   Get-FimWorkflow *accountname*
#>
function Get-FimWorkflow
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   Position=0)]
        $Name
    )

    Begin
    {
    }
    Process
    {
        return (Export-FimConfig -CustomConfig ("/WorkflowDefinition[DisplayName='{0}']" -f $Name) -OnlyBaseResources -Uri "http://localhost:5725/" | Convert-FimExportToPSObject)
    }
    End
    {
    }
}
Function Get-HashValue {
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=1)]
        [String] $String,

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   Position=0)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("SHA1", "SHA256", "MD5")]
        [String] $Algorithm = "SHA1"
    )

    Process
    {
        if($Algorithm -eq "SHA1") {
            $hasher = new-object System.Security.Cryptography.SHA1Managed
        } elseif($Algorithm -eq "SHA256") {
            $hasher = new-object System.Security.Cryptography.SHA256Managed
        } elseif($Algorithm -eq "MD5") {
            $hasher = new-object System.Security.Cryptography.MD5CryptoServiceProvider
        }
        $toHash = [System.Text.Encoding]::UTF8.GetBytes($String)
        $hashByteArray = $hasher.ComputeHash($toHash)
        $res = ""
        foreach($byte in $hashByteArray)
        {
             $res += [System.String]::Format("{0:X2}", $byte)
        }
        return $res;
    }

}

Function Get-LevenshteinDistance {
    [CmdletBinding()]
    [OutputType([int])]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [String]$First,

        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [String]$Second,

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=2)]
        [switch]$IgnoreCase
    )

    Begin
    {
    }
    Process
    {
        $len1 = $First.length
        $len2 = $Second.length
 
        # If either string has length of zero, the # of edits/distance between them is simply the length of the other string
        if($len1 -eq 0) { return $len2 }
        if($len2 -eq 0) { return $len1 }
 
        # make everything lowercase if IgnoreCase flag is set
        if($IgnoreCase)
        {
            $first = $first.tolowerinvariant()
            $second = $second.tolowerinvariant()
        }
 
        # create 2d Array to store the "distances"
        $dist = new-object -type 'int[,]' -arg ($len1+1),($len2+1)
 
        # initialize the first row and first column which represent the 2
        # strings we're comparing
        for($i = 0; $i -le $len1; $i++) 
        {
            $dist[$i,0] = $i
        }
        for($j = 0; $j -le $len2; $j++) 
        {
            $dist[0,$j] = $j
        }
 
        $cost = 0
 
        for($i = 1; $i -le $len1;$i++)
        {
            for($j = 1; $j -le $len2;$j++)
            {
                if($second[$j-1] -ceq $first[$i-1])
                {
                    $cost = 0
                }
                else   
                {
                    $cost = 1
                }
    
                # The value going into the cell is the min of 3 possibilities:
                # 1. The cell immediately above plus 1
                # 2. The cell immediately to the left plus 1
                # 3. The cell diagonally above and to the left plus the 'cost'
                $tempmin = [System.Math]::Min(([int]$dist[($i-1),$j]+1) , ([int]$dist[$i,($j-1)]+1))
                $dist[$i,$j] = [System.Math]::Min($tempmin, ([int]$dist[($i-1),($j-1)] + $cost))
            }
        }
 
        # the actual distance is stored in the bottom right cell
        return $dist[$len1, $len2];
    }
    End
    {
    }

}

<#
.Synopsis
   Returns excel line for deployment excel file
.DESCRIPTION
   Returns excel line for deployment excel file
.EXAMPLE
   Get-FIMObjectByXPath /SynchronizationRule | Get-MIMSynchornizationRuleAsExcelLine
#>
function Get-MIMSynchornizationRuleAsExcelLine
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $SynchronizationRule
    )

    Begin
    {
    }
    Process
    {
        "{0}`t{1}`t{2}`t{3}`t{4}`t{5}`t{6}`t{7}`t{8}`t{9}`t{10}`t{11}" -f 
            #(Get-FIMObjectByXPath ("/ma-data[ObjectID=""{0}""]" -f $SynchronizationRule.ManagementAgentID -replace "urn:uuid:","")).DisplayName,
            $SynchronizationRule.DisplayName,
            $SynchronizationRule.FlowType,
            $SynchronizationRule.ConnectedObjectType,
            $SynchronizationRule.ILMObjectType,
            ($SynchronizationRule.ConnectedSystemScope -join ";;;"),
            $SynchronizationRule.CreateConnectedSystemObject,
            $SynchronizationRule.CreateILMObject,
            $SynchronizationRule.DisconnectConnectedSystemObject,
            ($SynchronizationRule.RelationshipCriteria -join ";;;"),
            ($SynchronizationRule.PersistentFlow -join ";;;"),
            ($SynchronizationRule.InitialFlow -join ";;;"),
            ($SynchronizationRule.ExistenceTest -join ";;;")

    }
    End
    {
    }
}

Function Get-OUFromDistinguishedName {
    [CmdletBinding()]
    [OutputType([string])]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $DN
    )

    Begin
    {
    }
    Process
    {
        $DN -split "[^\\],", 2 | select -last 1
    }
    End
    {
    }

}

Function Get-PowerShellProfileOneTimeScript {
    [CmdletBinding()]
    [Alias()]
    [OutputType([string])]
    Param()

    '"https://dl.dropboxusercontent.com/u/6872078/PS/365.psm1",
"https://dl.dropboxusercontent.com/u/6872078/PS/ad.psm1",
"https://dl.dropboxusercontent.com/u/6872078/PS/adfs.psm1",
"https://dl.dropboxusercontent.com/u/6872078/PS/dnvgl.psm1",
"https://dl.dropboxusercontent.com/u/6872078/PS/fim.psm1",
"https://dl.dropboxusercontent.com/u/6872078/PS/fimpsmodule.psm1",
"https://dl.dropboxusercontent.com/u/6872078/PS/gwrnd.psm1",
"https://dl.dropboxusercontent.com/u/6872078/PS/linqxml.psm1",
"https://dl.dropboxusercontent.com/u/6872078/PS/gwrnddsc.psm1",
"https://dl.dropboxusercontent.com/u/6872078/PS/tools.psm1" | foreach {
    Write-Verbose "Downloading file $($_)" -Verbose
    $wc = New-Object System.Net.WebClient
    $file = "{0}\{1}" -f $env:TEMP, ($_ -split "/" | select -last 1)
    $wc.DownloadFile($_, $file)

    Import-Module $file -DisableNameChecking
    Remove-Item $file -Force
}'

}

Function Get-RandomPassword {
    [CmdletBinding()]
    [OutputType([string])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=0)]
        $Length = 32
    )

    Begin
    {
        $possibleCharacters = "a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z","A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","1","2","3","4","5","6","7","8","9"

    }
    Process
    {
        if($Length -lt 3) {
            Write-Error "Length too small"
        }
        do {
            $password = (1..$Length | foreach{$possibleCharacters | Get-Random -Count 1}) -join ""
        } while($password -cnotmatch "[a-z]" -or $password -cnotmatch "[A-Z]" -or $password -notmatch "[1-9]")
        return $password
    }
    End
    {
    }

}

Function Get-ScheduledTask2008 {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True,Position=0)]
        [string]$Name
    )

    return schtasks.exe /Query /V /FO CSV /TN "$Name" | ConvertFrom-Csv

}

Function Group-Object2 {
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=3)]
        $Object,

        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   Position=0)]
        [String] $Property,

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=1)]
        [String] $ExpandProperty = $null,

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=2)]
        [Boolean] $AsHashTable = $false
    )

    Begin
    {
        $_workingHashmap = @{}
    }
    Process
    {
        $groupValue = $Object.$Property
        if(!$groupValue) {
            Write-Verbose "Empty groupValue"
            $groupValue = ""
        }
        
        $groupObject = $Object
        if($ExpandProperty -ne $null -and $ExpandProperty -ne "") {
            Write-Verbose "Expanding property $ExpandProperty"
            $groupObject = $Object.$ExpandProperty
        }

        if(!$_workingHashmap[$groupValue]) {
            $_workingHashmap[$groupValue] = @()
        }
        $_workingHashmap[$groupValue] += $groupObject
    }
    End
    {
        if($AsHashTable) {
            return $_workingHashmap
        } else {
            $_workingHashmap.Keys | foreach {
                $_t = @{
                    Count = $_workingHashmap[$_].Count
                    Name = $_ 
                    Group = $_workingHashmap[$_]
                }
                New-Object -TypeName PSCustomObject -Property $_t 
            }
        }
    }

}

<#
.Synopsis
   Function to invoke FIM run profiles
.DESCRIPTION
   This method uses WMI to trigger FIM run profiles.
.EXAMPLE
   Invoke-FIMRunProfile "AD" "Full import"
   
   This example trigger the "Full import" run profile on the "AD" management agent
.EXAMPLE
   The following example trigger the "Full import" run profile on the "AD" management agent

   Invoke-FIMRunProfile -MA "AD" -RunProfile "Full import"   
.EXAMPLE
   Invoke-FIMRunProfile "AD"
   
   This will trigger the "Delta import Delta sync" run profile on the "AD" management agent
#>
function Invoke-FIMRunProfile
{
    [CmdletBinding()]
    [OutputType([object])]
    Param
    (
        # The Management Agent name
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$false,
                   Position=0)]
        [string] $MA,

        # The run profile to trigger
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=1)]
        [string[]] $RunProfile = @("Delta import Delta sync"),
        
        # Only trigger RunProfile if there are something to export
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=2)]
        [boolean] $DoNotRunWhenNoExports = $false
            
        
    )

    Begin
    {
        # Connect to database
        Write-Verbose ("Connecting to WMI root/MicrosoftIdentityIntegrationServer class MIIS_ManagementAgent")
        $wmi = Get-WmiObject -Namespace root/MicrosoftIdentityIntegrationServer -Class MIIS_ManagementAgent
    }
    Process
    {
        $WMIMAs = ($wmi | where{$_.Name -like $MA})
        
        foreach($WMIMA in $WMIMAs) {
            if($DoNotRunWhenNoExports -and (([int]$WMIMA.NumExportAdd().ReturnValue + [int]$WMIMA.NumExportDelete().ReturnValue + [int]$WMIMA.NumExportUpdate().ReturnValue) -eq 0)) {
                Write-Verbose "Found nothing to export"
                $result = @{ReturnValue="Nothing to export"}
                New-Object -TypeName PSObject -Property @{"Management Agent"=$WMIMA.Name;"Run Profile"=$RunProfile;Result=$result.ReturnValue}
            } else {
                # Execute WMI query to run the run profile and store the result in $result
                Write-Verbose ("Executing run profile ""{0}""" -f $RunProfile)
                $RunProfile | Foreach {
                    $result = $WMIMA.Execute($_)
                    New-Object -TypeName PSObject -Property @{"Management Agent"=$WMIMA.Name;"Run Profile"=$_;Result=$result.ReturnValue}
                }
            }    
        }
        
        
    }
    End
    {
        
    }
}
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
Function Load-Assembly {
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Assembly
    )

    Begin
    {
    }
    Process
    {
        return [System.Reflection.Assembly]::LoadWithPartialName($Assembly)
    }
    End
    {
    }

}

Function Load-Credential {
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [String] $Name = $null
    )

    Begin
    {
    }
    Process
    {
        $_file = "$($env:APPDATA)\credentials.csv"

        if(!(Test-path $_file)) {
            Write-Error "No such file: $_file" -ErrorAction Stop
        }
        
        if(!$Name -or $Name.Length -eq 0) {
            $_Credential = Import-Csv $_file | Out-Gridview -OutputMode Single -Title "Choose credential"
        } else {
            $_Credential = Import-Csv $_file | Where{$_.Name -like $Name}
            if(($_Credential| measure).Count -gt 1) {
                $_Credential = $_Credential | Out-Gridview -OutputMode Single -Title "Choose credential"
            }
        }

        if(!$_Credential) {
            Write-Error "No such credential: $Name"
        } else {
            return New-Object System.Management.Automation.PSCredential($_Credential.Username, ($_Credential.Password | ConvertTo-SecureString))
        }
    }
    End
    {
    }

}

Function New-DirectoryFileHash {
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   Position=0)]
        [String] $OutputPath = $null,

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=1)]
        [String] $Path = "."
    )

    Begin
    {
    }
    Process
    {
        if(!(Test-path $path -PathType Container)) {
            Write-Error -ErrorAction Stop "Not a folder: $path"
        }

        $Path = (Get-Item $Path).FullName


        dir -Recurse $Path -File | Get-FileHash | Select Hash, @{Label="RelativePath"; Expression={$_.Path.Replace($path,"")}} | Export-Clixml -Path $OutputPath
        
    }
    End
    {
    }

}

<#
.Synopsis
   Copies the input workflow definition to new workflow object
.DESCRIPTION
   Copies the input workflow definition to new workflow object
.EXAMPLE
   Get-FIMWorkflow *accountname* | New-FIMWorkflowCopy
#>
function New-FIMWorkflowCopy
{
    [CmdletBinding()]
    [OutputType([int])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $Source
    )

    Begin
    {
    }
    Process
    {
        $attributes = @(
            New-FimImportChange -Operation None -AttributeName 'DisplayName' -AttributeValue "___COPY - $($Source.DisplayName)"
            New-FimImportChange -Operation None -AttributeName 'RunOnPolicyUpdate' -AttributeValue $Source.RunOnPolicyUpdate
            New-FimImportChange -Operation None -AttributeName 'RequestPhase' -AttributeValue $Source.RequestPhase
            New-FimImportChange -Operation None -AttributeName 'XOML' -AttributeValue $Source.XOML
        )

        New-FimImportObject -ObjectType "WorkflowDefinition" -State Create -Changes $attributes -ApplyNow:$true -PassThru -SkipDuplicateCheck:$true

    }
    End
    {
    }
}
Function New-ObjectFromHashmap {
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $Hashmap
    )

    Begin
    {
    }
    Process
    {
        New-Object -TypeName PSCustomObject -Property $Hashmap
    }
    End
    {
    }

}

Function New-Progressbar {
    [CmdletBinding()]
    Param
    (
        # Total count
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$false,
                   Position=0)]
        [int]$TotalCount,

        # Activity name
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=1)]
        [string]$ActivityName = "Running",

        # Time estimation
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=2)]
        [boolean]$TimeEstimationEnabled = $true
    )

    # Create new module instance   
    $m =  New-Module -ScriptBlock {
        # Internal variables
        $script:total = 1;
        $script:current = 0;
        $script:ActivityName = " ";
        $script:startTime = Get-Date;
        $script:timeEstimation = $false;
        # Functions with obvious method names
        function setActivityName($name) {$script:ActivityName = $name}
        function setTotal($tot) { $script:total = $tot}
        function getTotal($tot) { return $script:total}
        function enableTimeEstimation() {$script:timeEstimation = $true}
        function disableTimeEstimation() {$script:timeEstimation = $false}


        # Progress the progressbar one step. Optional parameter Text for defining the status message
        function Progress {
            Param
            (
                [Parameter(Mandatory=$false,
                    ValueFromPipelineByPropertyName=$false,
                    Position=0)]
                [string]$Text = ("{0}/{1}" -f $script:current, $script:total)
            )

            $params = @{
                Activity = $script:ActivityName
                Status = $Text
                PercentComplete = ($script:current / $script:total * 100)
            }

            if($script:timeEstimation) {
                if($script:current -gt 5) {
                    $params["SecondsRemaining"] = (((Get-Date) - $script:startTime).TotalSeconds / $script:current) * ($script:total - $script:current)
                }
            }

            Write-Progress @params
            
            if($script:current -lt $script:total) {
                $script:current += 1
            } else {
                Write-Warning "Progressbar incremented too far"
            }
        }
        function Complete() {Write-Progress -Activity $script:ActivityName -Status $script:total -PercentComplete 100 -Completed}
        export-modulemember -function setTotal,getTotal,Progress,Complete,setActivityName,enableTimeEstimation,disableTimeEstimation
    } -AsCustomObject

    # Set initial values
    $m.setTotal($TotalCount)
    $m.setActivityName($ActivityName)

    if($TimeEstimationEnabled) {
        $m.enableTimeEstimation()
    }

    return $m;

}

Function New-RDPFile {
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Hostname,

        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        $Gateway,

        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=2)]
        $File,

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=3)]
        $Port = 3389
    )

    Begin
    {
    }
    Process
    {
        $Content = 
"screen mode id:i:2
use multimon:i:0
desktopwidth:i:2560
desktopheight:i:1440
session bpp:i:32
winposstr:s:0,1,326,0,2560,1360
compression:i:1
keyboardhook:i:2
audiocapturemode:i:0
videoplaybackmode:i:1
connection type:i:7
networkautodetect:i:1
bandwidthautodetect:i:1
displayconnectionbar:i:1
enableworkspacereconnect:i:0
disable wallpaper:i:0
allow font smoothing:i:0
allow desktop composition:i:0
disable full window drag:i:1
disable menu anims:i:1
disable themes:i:0
disable cursor setting:i:0
bitmapcachepersistenable:i:1
full address:s:{0}:{1}
audiomode:i:0
redirectprinters:i:0
redirectcomports:i:0
redirectsmartcards:i:0
redirectclipboard:i:1
redirectposdevices:i:0
autoreconnection enabled:i:1
authentication level:i:2
prompt for credentials:i:0
negotiate security layer:i:1
remoteapplicationmode:i:0
alternate shell:s:
shell working directory:s:
gatewayhostname:s:{2}
gatewayusagemethod:i:1
gatewaycredentialssource:i:4
gatewayprofileusagemethod:i:1
promptcredentialonce:i:0
gatewaybrokeringtype:i:0
use redirection server name:i:0
rdgiskdcproxy:i:0
kdcproxyname:s:
drivestoredirect:s:" -f $Hostname, $Port, $Gateway
        Set-Content -Value $Content -Path $File
    }
    End
    {
    }

}

Function Open-IseFiles {
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [System.IO.FileInfo] $FullName
    )

    Begin
    {
    }
    Process
    {
        if(!(Test-path $FullName)) {
            Write-Error "No such file: $FullName"
            return;
        }

        $psise.CurrentPowerShellTab.Files.Add($FullName)
    }
    End
    {
    }

}

Function Out-Excel {
    param(
        $Path = "$env:temp\$(Get-Date -Format yyyyMMddHHmmss) $(Get-Random -min 1 -max 999).csv",
        $OpenExcel = $true
    )
    
    $input | Export-CSV -Path $Path -UseCulture -Encoding UTF8 -NoTypeInformation
    
    if($OpenExcel) {
        Invoke-Item -Path $Path
    }

    return $Path

}

# Set the PowerShell prompt to PS>
function prompt{
    Write-Host -ForegroundColor White -NoNewline ($env:COMPUTERNAME).ToUpper()
    Write-Host -ForegroundColor Red " PS" -NoNewline
    Write-Host -ForegroundColor White -NoNewline ">"
    return " "
}
Function Query-Sql {
    [CmdletBinding(DefaultParameterSetName='ConnectionString')]
    Param
    (
        # Query
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$false)]
        [String] $Query,

        # Connection string
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$false,
                   ParameterSetName='ConnectionString')]
        [String] $ConnectionString,

        # Server
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$false,
                   ParameterSetName='ServerDatabaseProvided')]
        [String] $Server,

        # Database
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$false,
                   ParameterSetName='ServerDatabaseProvided')]
        [String] $Database,

        # Credential
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false)]
        $Credential
    )

    Begin
    {
        if($Server -and $Database) {
            $ConnectionString = "Server={0}; Database={1}" -f $Server, $Database
            if(!$Credential) {
                $ConnectionString += "; Integrated Security=True"
            }
            Write-Verbose "Connection string: $ConnectionString"
        }

        $sqlConnection = New-Object System.Data.SqlClient.SQLConnection
        $sqlConnection.ConnectionString = $ConnectionString

        if($Credential) {
            $Credential.Password.MakeReadOnly()
            $sqlCredential = New-Object System.Data.SqlClient.SqlCredential($Credential.UserName, $Credential.Password)
            $sqlConnection.Credential = $sqlCredential
        }

        $sqlConnection.Open()

        $sqlQuery = New-Object System.Data.SqlClient.SqlCommand
        $sqlQuery.CommandText = $Query
        $sqlQuery.Connection = $sqlConnection
        
    }
    Process
    {
        $reader = $sqlQuery.ExecuteReader()
        $columns = $reader.GetSchemaTable() | Select-Object -ExpandProperty ColumnName
        while($reader.Read()) {
            $props = @{}
            foreach($column in $columns) {
                $props[$column] = $reader[$column];
            }
            New-Object -TypeName PSObject -Property $props
        }
        
        $reader.Close()
    }
    End
    {
        $sqlConnection.Close()
    }

}

Function Read-Credentials {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True,Position=0)]
        [string]$File
    )

    $obj = Import-Clixml -Path $File
    return New-Object System.Management.Automation.PSCredential($obj.username, ($obj.password | ConvertTo-SecureString))

}

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
Function Repeat-Command {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=2)]
        [scriptblock] $ScriptBlock,

        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   Position=1)]
        [Int] $Sleep,

        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   Position=0)]
        [Int] $Times
    )

    Begin
    {
    }
    Process
    {
        1..$times | foreach {
            $ScriptBlock.InvokeReturnAsIs()
            if($_ -ne $times) {
                 sleep -Milliseconds $Sleep
            }
        }
    }
    End
    {
    }

}

Function Replace-String {
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=4)]
        $InputObject,

        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   Position=0)]
        [String] $Pattern = $null,

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=1)]
        [String] $Replacement = "",

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=2)]
        [String] $Property = $null,

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=3)]
        [Boolean] $CaseSensitive = $false
    )

    Begin
    {
    }
    Process
    {
        if($Property) {
            $OutputObject = $InputObject | select -Property * 
            if($CaseSensitive) {
                $OutputObject.$Property = $OutputObject.$Property -creplace $Pattern, $Replacement
            } else {
                $OutputObject.$Property = $OutputObject.$Property -replace $Pattern, $Replacement
            }
            $OutputObject
        } else {
            if($CaseSensitive) {
                $InputObject -creplace $Pattern, $Replacement
            } else {
                $InputObject -ireplace $Pattern, $Replacement
            }
        }
    }
    End
    {
    }

}

Function Run-ScheduledTask2008 {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True,Position=0)]
        [string]$Name
    )

    $ret = schtasks.exe /Run /TN "$Name"

    return ($ret -like "SUCCESS:*") -eq $true

}

Function Save-Credential {
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [String] $Name,

        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [System.Management.Automation.PSCredential] $Credential,

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=2)]
        [String] $Description
    )

    Begin
    {
    }
    Process
    {
        $_file = "$($env:APPDATA)\credentials.csv"
        
        if(!(test-path $_file)) {
            Set-Content -Path $_file -Value "Name,Description,Username,Password"
        }

        $_Credentials = @(Import-Csv $_file | Where{$_.Name -ne $Name})

        $_Credentials += [PSCustomObject] @{
            Name = $Name 
            Description = $Description
            Username = $Credential.UserName
            Password = ($Credential.Password | ConvertFrom-SecureString)
        }
        
        $_Credentials | Export-Csv $_file -NoTypeInformation
    }
    End
    {
    }

}

Function Save-IseFiles {
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=0)]
        [string] $Path = "$($env:temp)\isefiles\",

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=1)]
        [boolean] $Confirm = $true
    )

    Begin
    {
    }
    Process
    {
        if(!(Test-path $Path)) {
            mkdir $Path | Out-Null
        }

        $psISE.PowerShellTabs | Foreach {
            $_.Files | 
            where{!$_.IsSaved} |
                foreach {
                    if((Test-Path $_.FullPath)) {
                        Write-Verbose "File already exists, so this is an unsaved file: $($_.FullPath)"
                        $result = "y"
                        if($Confirm) {
                            $result = Read-Host "Save $($_.FullPath)? (y/N)"
                        }

                        if($result -eq "y") {
                            Write-Verbose "Saving: $($_.FullPath)"
                            $_.Save()
                        }
                    } else {
                        $ActualPath = (Join-Path $Path ([guid]::newguid()).ToString()) + ".ps1"
                        Write-Verbose "File does not exist, so this is a temp file - saving to $ActualPath"
                        $_.SaveAs($ActualPath)
                    }
                }
        }
    }
    End
    {
    }

}

Function Search-IseFiles {
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$false,
                   Position=0)]
        [ScriptBlock] $Where,

        # Param1 help description
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=1)]
        [string] $Path = "$($env:temp)\isefiles\"
    )

    Begin
    {
    }
    Process
    {
        if(!(Test-path $Path)) {
            Write-Error "No such path: $Path"
            return;
        }

        dir $Path | Where {
            Get-ContentAsString -Path $_.FullName | Where -FilterScript $Where
        }
    }
    End
    {
    }

}

Function Send-SlackMessageLumagate {
    [CmdletBinding()]
    [Alias("slackluma")]
    Param
    (
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   Position=0)]
        [String] $Text = 'Finished',

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=1)]
        [String[]] $Channels = @('@mariussm'),

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=2)]
        [String] $WebhookUri = 'https://hooks.slack.com/services/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'

    )

    Process
    {
        foreach($Channel in $Channels) {
            $body = @{text = $Text; channel = $Channel} | ConvertTo-Json
            $result = Invoke-RestMethod -Method Post -Uri $WebhookUri -Body $body
            if($result -ne "ok") {
                Write-Error "Expected 'ok' back from slack, but got: $result"
            }
        }
    }

}

Function Send-SlackMessagePrivate {
    [CmdletBinding()]
    [Alias("slackpriv")]
    Param
    (
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   Position=0)]
        [String] $Text = 'Finished',

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=1)]
        [String[]] $Channels = @('#general'),

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=2)]
        [String] $WebhookUri = 'https://hooks.slack.com/services/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'

    )

    Process
    {
        foreach($Channel in $Channels) {
            $body = @{text = $Text; channel = $Channel} | ConvertTo-Json
            $result = Invoke-RestMethod -Method Post -Uri $WebhookUri -Body $body
            if($result -ne "ok") {
                Write-Error "Expected 'ok' back from slack, but got: $result"
            }
        }
    }

}

Function Send-TeamsNotification {
    [CmdletBinding()]
    [Alias("slackluma")]
    Param
    (
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   Position=0)]
        [String] $Text = 'Finished',

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=1)]
        [String[]] $Mentions = @('@marius.solbakken@lumagate.com'),

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=2)]
        [String] $Title = $null,
        
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=3)]
        [String] $WebhookUri = 'https://outlook.office365.com/webhook/23996283-9b34-4edc-8223-59a217712e96@fb6de64b-9546-4830-9741-05cb9344399d/IncomingWebhook/dce5d8591b13440784881266cb5da601/2db586cb-7a8a-448a-b2db-7995936cf808'
    )

    Process
    {
        $Json = @{text = $Text}

        if($Title -ne $null) {
            $Json["title"] = $Title
        }

        $Mentions | foreach {
            $Json.Text += " $($_)"
        }

        $body = $json | ConvertTo-Json
        $result = Invoke-RestMethod -Method Post -Uri $WebhookUri -Body $body
        if($result -ne "1") {
            Write-Error "Expected '1' back from teams, but got: $result"
        }
    }

}

Function Set-WorkingDirectoryToCurrentISEFilePath {
    [CmdletBinding()]
    [Alias("cdise")]
    Param
    ()

    Process
    {
        if($psise.CurrentFile.FullPath) {
            cd (split-path -Parent -Path $psise.CurrentFile.FullPath)
        }
    }
    

}

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
Function Split-String {
    [CmdletBinding()]
    [OutputType([string[]])]
    Param
    (
        # The input string object
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        [String] $InputObject,

        # Split delimiter
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=1)]
        [String] $Delimiter = "`n",

        # Do trimming or not
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=2)]
        [Boolean] $Trim = $true

    )

    Begin{}
    Process {
        if($Trim) {
            return $InputObject -split $Delimiter | foreach{$_.Trim()}
        } else {
            return $InputObject -split $Delimiter
        }
    } 
    End{}

}

<#
 
.SYNOPSIS
Generic wrapper script that tries to ensure that a script block successfully finishes execution in O365 against a large object count.

Works well with intense operations that may cause throttling

.DESCRIPTION
Wrapper script that tries to ensure that a script block successfully finishes execution in O365 against a large object count.

It accomplishs this by doing the following:
* Monitors the health of the Remote powershell session and restarts it as needed.
* Restarts the session every X number seconds to ensure a valid connection.
* Attempts to work past session related errors and will skip objects that it can't process.
* Attempts to calculate throttle exhaustion and sleep a sufficient time to allow throttle recovery

.PARAMETER Agree
Verifies that you have read and agree to the disclaimer at the top of the script file.

.PARAMETER AutomaticThrottle
Calculated value based on your tenants powershell recharge rate.
You tenant recharge rate can be calculated using a Micro Delay Warning message.

Look for the following line in your Micro Delay Warning Message
Balance: -1608289/2160000/-3000000 

The middle value is the recharge rate.
Divide this value by the number of milliseconds in an hour (3600000)
And subtract the result from 1 to get your AutomaticThrottle value

1 - (2160000 / 3600000) = 0.4

Default Value is .25

.PARAMETER Credential
Credential object for logging into Exchange Online Shell.
Prompts if there is non provided.

.PARAMETER IdentifyingProperty
What property of the objects we are processing that will be used to identify them in the log file and host
If the value is not set by the user the script will attempt to determine if one of the following properties is present
"DisplayName","Name","Identity","PrimarySMTPAddress","Alias","GUID"

If the value is not set and we are not able to match a well known property the script will generate an error and terminate.

.PARAMETER LogFile
Location and file name for the log file.

.PARAMETER ManualThrottle
Manual delay of X number of milliseconds to sleep between each cmdlets call.
Should only be used if the AutomaticThrottle isn't working to introduce sufficent delay to prevent Micro Delays

.PARAMETER NonInteractive
Suppresses output to the screen.  All output will still be in the log file.

.PARAMETER Recipients
Array of objects to operate on. This can be mailboxes or any other set of objects.
Input must be an array!
Anything comming in from the array can be accessed in the script block using $input.property

.PARAMETER ResetSeconds
How many seconds to run the script block before we rebuild the session with O365.

.PARAMETER ScriptBlock
The script that you want to robustly execute against the array of objects.  The Recipient objects will be provided to the cmdlets in the script block
and can be accessed with $input as if you were pipelining the object.

.LINK
http://EHLO.Link

.OUTPUTS
Creates the log file specified in -logfile.  Logfile contains a record of all actions taken by the script.

.EXAMPLE
invoke-command -scriptblock {Get-mailbox -resultsize unlimited | select-object -property Displayname,PrimarySMTPAddress,Identity} -session (get-pssession) | export-csv c:\temp\mbx.csv
$mbx = import-csv c:\temp\mbx.csv
$cred = get-Credential
.\Start-RobustCloudCommand.ps1 -Agree -Credential $cred -recipients $mbx -logfile C:\temp\out.log -ScriptBlock {Set-Clutter -identity $input.PrimarySMTPAddress.tostring() -enable:$false}

Gets all mailboxes from the service returning only Displayname,Identity, and PrimarySMTPAddress.  Exports the results to a CSV
Imports the CSV into a variable
Gets your O365 Credential
Executes the script setting clutter to off

.EXAMPLE
invoke-command -scriptblock {Get-mailbox -resultsize unlimited | select-object -property Displayname,PrimarySMTPAddress,Identity} -session (get-pssession) | export-csv c:\temp\recipients.csv
$recipients = import-csv c:\temp\recipients.csv
$cred = Get-Credential
.\Start-RobustCloudCommand.ps1 -Agree -Credential $cred -recipients $recipients -logfile C:\temp\out.log -ScriptBlock {Get-MobileDeviceStatistics -mailbox $input.PrimarySMTPAddress.tostring() | Select-Object -Property @{Name = "PrimarySMTPAddress";Expression={$input.PrimarySMTPAddress.tostring()}},DeviceType,LastSuccessSync,FirstSyncTime | Export-Csv c:\temp\stats.csv -Append }

Gets All Recipients and exports them to a CSV (for restartability)
Imports the CSV into a variable
Gets your O365 Credentials
Executs the script to gather EAS Device statistics and output them to a csv file


#>

function Start-RobustCloudCommand
{
    [CmdletBinding()]

    Param(
	    [switch]$Agree,
	    [Parameter(Mandatory=$true)]
	    [string]$LogFile,
	    [Parameter(Mandatory=$true)]
	    $Recipients,
	    [Parameter(Mandatory=$true)]
	    [ScriptBlock]$ScriptBlock,
	    $Credential,
	    [int]$ManualThrottle=0,
	    [double]$ActiveThrottle=.25,
	    [int]$ResetSeconds=870,
	    [string]$IdentifyingProperty,
	    [Switch]$NonInteractive
    )

    Process {
        # Writes output to a log file with a time date stamp
        Function Write-Log {
	        Param ([string]$string)
	
	        # Get the current date
	        [string]$date = Get-Date -Format G
		
	        # Write everything to our log file
	        ( "[" + $date + "] - " + $string) | Out-File -FilePath $LogFile -Append
	
	        # If NonInteractive true then supress host output
	        if (!($NonInteractive)){
		        ( "[" + $date + "] - " + $string) | Write-Host
	        }
        }

        # Sleeps X seconds and displays a progress bar
        Function Start-SleepWithProgress {
	        Param([int]$sleeptime)

	        # Loop Number of seconds you want to sleep
	        For ($i=0;$i -le $sleeptime;$i++){
		        $timeleft = ($sleeptime - $i);
		
		        # Progress bar showing progress of the sleep
		        Write-Progress -Activity "Sleeping" -CurrentOperation "$Timeleft More Seconds" -PercentComplete (($i/$sleeptime)*100);
		
		        # Sleep 1 second
		        start-sleep 1
	        }
	
	        Write-Progress -Completed -Activity "Sleeping"
        }

        # Setup a new O365 Powershell Session
        Function New-CleanO365Session {
	
	        # If we don't have a credential then prompt for it
	        $i = 0
	        while (($Credential -eq $Null) -and ($i -lt 5)){
		        $script:Credential = Get-Credential -Message "Please provide your Exchange Online Credentials"
		        $i++
	        }
	
	        # If we still don't have a credentail object then abort
	        if ($Credential -eq $null){
		        Write-log "[Error] - Failed to get credentials"
		        Write-Error -Message "Failed to get credentials" -ErrorAction Stop
	        }

	        Write-Log "Removing all PS Sessions"

	        # Destroy any outstanding PS Session
	        Get-PSSession | Remove-PSSession -Confirm:$false
	
	        # Force Garbage collection just to try and keep things more agressively cleaned up due to some issue with large memory footprints
	        [System.GC]::Collect()
	
	        # Sleep 15s to allow the sessions to tear down fully
	        Write-Log ("Sleeping 15 seconds for Session Tear Down")
	        Start-SleepWithProgress -SleepTime 15

	        # Clear out all errors
	        $Error.Clear()
	
	        # Create the session
	        Write-Log "Creating new PS Session"
	
	        $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $Credential -Authentication Basic -AllowRedirection
		
	        # Check for an error while creating the session
	        if ($Error.Count -gt 0){
	
		        Write-Log "[ERROR] - Error while setting up session"
		        Write-log $Error
		
		        # Increment our error count so we abort after so many attempts to set up the session
		        $ErrorCount++
		
		        # if we have failed to setup the session > 3 times then we need to abort because we are in a failure state
		        if ($ErrorCount -gt 3){
		
			        Write-log "[ERROR] - Failed to setup session after multiple tries"
			        Write-log "[ERROR] - Aborting Script"
			        exit
		
		        }
		
		        # If we are not aborting then sleep 60s in the hope that the issue is transient
		        Write-Log "Sleeping 60s so that issue can potentially be resolved"
		        Start-SleepWithProgress -sleeptime 60
		
		        # Attempt to set up the sesion again
		        New-CleanO365Session
	        }
	
	        # If the session setup worked then we need to set $errorcount to 0
	        else {
		        $ErrorCount = 0
	        }
	
	        # Import the PS session
	        $null = Import-PSSession $session -AllowClobber
	
	        # Set the Start time for the current session
	        Set-Variable -Scope script -Name SessionStartTime -Value (Get-Date)
        }

        # Verifies that the connection is healthy
        # Goes ahead and resets it every $ResetSeconds number of seconds either way
        Function Test-O365Session {
	
	        # Get the time that we are working on this object to use later in testing
	        $ObjectTime = Get-Date
	
	        # Reset and regather our session information
	        $SessionInfo = $null
	        $SessionInfo = Get-PSSession
	
	        # Make sure we found a session
	        if ($SessionInfo -eq $null) { 
		        Write-Log "[ERROR] - No Session Found"
		        Write-log "Recreating Session"
		        New-CleanO365Session
	        }	
	        # Make sure it is in an opened state if not log and recreate
	        elseif ($SessionInfo.State -ne "Opened"){
		        Write-Log "[ERROR] - Session not in Open State"
		        Write-log ($SessionInfo | fl | Out-String )
		        Write-log "Recreating Session"
		        New-CleanO365Session
	        }
	        # If we have looped thru objects for an amount of time gt our reset seconds then tear the session down and recreate it
	        elseif (($ObjectTime - $SessionStartTime).totalseconds -gt $ResetSeconds){
		        Write-Log ("Session Has been active for greater than " + $ResetSeconds + " seconds" )
		        Write-Log "Rebuilding Connection"
		
		        # Estimate the throttle delay needed since the last session rebuild
		        # Amount of time the session was allowed to run * our activethrottle value
		        # Divide by 2 to account for network time, script delays, and a fudge factor
		        # Subtract 15s from the results for the amount of time that we spend setting up the session anyway
		        [int]$DelayinSeconds = ((($ResetSeconds * $ActiveThrottle) / 2) - 15)
		
		        # If the delay is >15s then sleep that amount for throttle to recover
		        if ($DelayinSeconds -gt 0){
		
			        Write-Log ("Sleeping " + $DelayinSeconds + " addtional seconds to allow throttle recovery")
			        Start-SleepWithProgress -SleepTime $DelayinSeconds
		        }
		        # If the delay is <15s then the sleep already built into New-CleanO365Session should take care of it
		        else {
			        Write-Log ("Active Delay calculated to be " + ($DelayinSeconds + 15) + " seconds no addtional delay needed")
		        }
				
		        # new O365 session and reset our object processed count
		        New-CleanO365Session
	        }
	        else {
		        # If session is active and it hasn't been open too long then do nothing and keep going
	        }
	
	        # If we have a manual throttle value then sleep for that many milliseconds
	        if ($ManualThrottle -gt 0){
		        Write-log ("Sleeping " + $ManualThrottle + " milliseconds")
		        Start-Sleep -Milliseconds $ManualThrottle
	        }
        }

        # If the $identifyingProperty has not been set then we attempt to locate a value for tracking modified objects
        Function Get-ObjectIdentificationProperty {
	        Param($object)
	
	        Write-Log "Trying to identify a property for displaying per object progress"
	
	        # Common properties to check
	        [array]$PropertiesToCheck = "DisplayName","Name","Identity","PrimarySMTPAddress","Alias","GUID"
	
	        # Set our counter to 0
	        $i = 0
	
	        # While we haven't found an ID property continue checking
	        while ([string]::IsNullOrEmpty($IdentifyingProperty))
	        {
	
		        # If we have gone thru the list then we need to throw an error because we don't have Identity information
		        # Set the string to bogus just to ensure we will exit the while loop
		        if ($i -gt ($PropertiesToCheck.length -1))
		        {
			        Write-Log "[ERROR] - Unable to find a common identity parameter in the input object"
			
			        # Create an error message that has all of the valid property names that we are looking for
			        $PropertiesToCheck | foreach { [string]$PropertiesString = $PropertiesString + "`"" + $_ + "`", " }
			        $PropertiesString = $PropertiesString.TrimEnd(", ")
			        [string]$errorstring = "Objects does not contain a common identity parameter " + $PropertiesString + " please use -IdentifyingProperty to set the identity value"
			
			        # Throw error
			        Write-Error -Message $errorstring -ErrorAction Stop
		        }
		
		        # Get the property we are testing out of our array
		        [string]$Property = $PropertiesToCheck[$i]
	
		        # Check the properties of the object to see if we have one that matches a well known name
		        # If we have found one set the value to that property
		        if ($object.$Property -ne $null)
		        { 
			        Write-log ("Found " + $Property + " to use for displaying per object progress")
			        Set-Variable -Scope script -Name IdentifyingProperty -Value $Property
		        }
		
		        # Increment our position counter
		        $i++
		
	        }
        }

        # Gather and print out information about how fast the script is running
        Function Get-EstimatedTimeToCompletion {
	        param([int]$ProcessedCount)
	
	        # Increment our count of how many objects we have processed
	        $ProcessedCount++
	
	        # Every 100 we need to estimate our completion time and write that out
	        if (($ProcessedCount % 100) -eq 0){
	
		        # Get the current date
		        $CurrentDate = Get-Date
		
		        # Average time per object in seconds
		        $AveragePerObject = (((($CurrentDate) - $ScriptStartTime).totalseconds) / $ProcessedCount)
		
		        # Write out session stats and estimated time to completion
		        Write-Log ("[STATS] - Total Number of Objects:     " + $ObjectCount)
		        Write-Log ("[STATS] - Number of Objects processed: " + $ProcessedCount)
		        Write-Log ("[STATS] - Average seconds per object:  " + $AveragePerObject)
		        Write-Log ("[STATS] - Estimated completion time:   " + $CurrentDate.addseconds((($ObjectCount - $ProcessedCount) * $AveragePerObject)))
	        }
	
	        # Return number of objects processed so that the variable in incremented
	        return $ProcessedCount
        }

        ####################
        # Main Script
        ####################

        # Force use of at least version 3 of powershell https://technet.microsoft.com/en-us/library/hh847765.aspx
        #Requires -version 3

        # Turns on strict mode https://technet.microsoft.com/library/03373bbe-2236-42c3-bf17-301632e0c428(v=wps.630).aspx
        Set-StrictMode -Version 2

        # Write creation date of script for version information
        Write-Log "Created 05/10/2016"

        # Statement to ensure that you have looked at the disclaimer or that you have removed this line so you don't have too
        if ($Agree -ne $true){ Write-Error "Please run the script with -Agree to indicate that you have read and agreed to the sample script disclaimer at the top of the script file" -ErrorAction Stop }
        else { Write-log "Agreed to Disclaimer" }

        # Log the script block for debugging purposes
        Write-log $ScriptBlock

        # Setup our first session to O365
        $ErrorCount = 0
        New-CleanO365Session

        # Get when we started the script for estimating time to completion
        $ScriptStartTime = Get-Date

        # Get the object count and write it out to be used in esitmated time to completion + logging
        [int]$ObjectCount = $Recipients.count
        [int]$ObjectsProcessed = 0

        # If we don't have an identifying property then try to find one
        if ([string]::IsNullOrEmpty($IdentifyingProperty))
        {
	        # Call our function for finding an identifying property and pass in the first recipient object
	        Get-ObjectIdentificationProperty -object $Recipients[0]
        }

        # Go thru each recipient object and execute the script block
        foreach ($object in $Recipients)
        {
	
	        # Set our initial while statement values
	        $TryCommand = $true
	        $errorcount = 0
	
	        # Try the command 3 times and exit out if we can't get it to work
	        # Record the error and restart the session each time it errors out
	        while ($TryCommand)
	        {
		        Write-log ("Running scriptblock for " + ($object.$IdentifyingProperty).tostring())
		
		        # Clear all errors
		        $Error.Clear()
	
		        # Test our connection and rebuild if needed
		        Test-O365Session
	
		        # Invoke the script block
		        Invoke-Command -InputObject $object -ScriptBlock $ScriptBlock
		
		        # Test for errors
		        if ($Error.Count -gt 0) 
		        {
			        # Write that we failed
			        Write-Log ("[ERROR] - Failed on object " + ($object.$IdentifyingProperty).tostring())
			        write-log $Error
			
			        # Increment error count
			        $errorcount++
			
				        # Handle if we keep failing on the object
				        if ($errorcount -ge 3)
				        {
					        Write-Log ("[ERROR] - Oject has failed three times " + ($object.$IdentifyingProperty).tostring())
					        Write-Log ("[ERROR] - Skipping Object")
					
					        # Increment the object processed count / Estimate time to completion
					        $ObjectsProcessed = Get-EstimatedTimeToCompletion -ProcessedCount $ObjectsProcessed
					
					        # Set trycommand to false so we abort the while loop
					        $TryCommand = $false
				        }
				        # Otherwise try the command again
				        else 
				        {
					        Write-Log ("Rebuilding session and trying again")
					        # Create a new session in case the error was due to a session issue
					        New-CleanO365Session
				        }
		        }
		        else 
		        {
			        # Since we didn't get an error don't run again
			        $TryCommand = $false
			
			        # Increment the object processed count / Estimate time to completion
			        $ObjectsProcessed = Get-EstimatedTimeToCompletion -ProcessedCount $ObjectsProcessed
		        }
	        }
        }

        Write-Log "Script Complete Destroying PS Sessions"
        # Destroy any outstanding PS Session
        Get-PSSession | Remove-PSSession -Confirm:$false
    }
}
Function Start-WaitForCmdletOutputChange {
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   Position=0)]
        [System.Management.Automation.ScriptBlock] $Script,

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=1)]
        [int] $SleepSec = 1
    )

    Begin
    {
    }
    Process
    {
        $Orig = $Script.Invoke() | out-string
        Write-Verbose "Original output: $Orig"
        do {
            $Result = $Script.Invoke() | out-string
            Write-Verbose "Result: $Result"
            Sleep -Seconds $SleepSec
        } until ($Result -ne $Orig)
    }
    End
    {
    }

}

<#
.Synopsis
   Waits until no MAs are active (or has been within the last 30 seconds)
.DESCRIPTION
   Waits until no MAs are active (or has been within the last 30 seconds)
.EXAMPLE
   Start-WaitForMIMSyncToBeIdle
#>
function Start-WaitForMIMSyncToBeIdle
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=0)]
        [int] $Wait = 30
    )

    Begin
    {
        $wmi = Get-WmiObject -Namespace root/MicrosoftIdentityIntegrationServer -Class MIIS_ManagementAgent
    }
    Process
    {
        if($wmi) {
            $sleepTime = 0
            do {
                $inProgress = $wmi | where {
                    $value = $_.RunEndTime().ReturnValue
                    if($value -eq "in-progress"){return $true}
                    if($value -ne "") {
                        (Get-Date ($value)) -gt (Get-Date).AddSeconds(0 - $wait)
                    }
                }

                sleep -Seconds $sleepTime
                $sleepTime = 10
            } while($inProgress)
        }
    }
    End
    {
    }
}
Function Start-WaitUntil {
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=2)]
        $Object,

        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   Position=0)]
        [System.Management.Automation.ScriptBlock]
        $CheckScript,

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=1)]
        [System.Management.Automation.ScriptBlock]
        $DoneScript = {}
    )

    Begin
    {
        $BreakDone = $false
    }
    Process
    {
        if($BreakDone) {
            break
        } elseif ($Object | where -FilterScript $CheckScript) {
            $DoneScript.Invoke()
            $BreakDone = $true
            break
        }
    }
    End
    {
        
    }

}

Function Store-Credentials {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [string]$File,
	
        [Parameter(Mandatory=$True)]
        [System.Management.Automation.PSCredential]$Credentials
    )

    $info = @{username=$Credentials.UserName;password=($Credentials.Password | ConvertFrom-SecureString)}
    $obj = New-Object -TypeName PSObject -Property $info
    $obj | Export-Clixml -Path $File

}

Function Test-DirectoryFileHash {
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   Position=0)]
        [String] $XMLPath = $null,

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=1)]
        [String] $Path = "."
    )

    Begin
    {
    }
    Process
    {
        if(!(Test-path $path -PathType Container)) {
            Write-Error -ErrorAction Stop "Not a folder: $path"
        }

        $Path = (Get-Item $Path).FullName

        $FilesFromXML = Import-Clixml -Path $XMLPath | group RelativePath -AsHashTable
        $FilesInFolder = @{}
        dir -Recurse $Path -File | Get-FileHash | Select Hash, @{Label="RelativePath"; Expression={$_.Path.Replace($path,"")}} | foreach{$FilesInFolder[$_.RelativePath] = $_} 
        
        $Errors = @()
        $FilesFromXML.Values | where{!$FilesInFolder.ContainsKey($_.RelativePath)} | foreach{
            $Errors += [PSCustomObject]@{File = $_.RelativePath; Error = "Missing"}            
        }
        
        $FilesInFolder.Values | Where{$FilesFromXML.ContainsKey($_.RelativePath)} | where{$FilesFromXML[$_.RelativePath].Hash -ne $_.Hash} | foreach{
            $Errors += [PSCustomObject]@{File = $_.RelativePath; Error = "File corrupt"}   
        }

        $Errors
    }
    End
    {
    }

}

Function Trim-String {
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=1)]
        $InputObject,

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   Position=0)]
        [String] $Property = $null
    )

    Begin
    {
    }
    Process
    {
        if($Property) {
            $OutputObject = $InputObject | select -Property * 
            $OutputObject.$Property = $OutputObject.$Property.Trim()
            $OutputObject
        } else {
            $InputObject.Trim()
        }
    }
    End
    {
    }

}

Function Write-SlackErrorPrivate {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   Position=0)]
        [String] $Text = 'Finished',

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=1)]
        [String[]] $Channels = @('#general'),

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=2)]
        [String] $WebhookUri = 'https://hooks.slack.com/services/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'
    )

    Process
    {
        foreach($Channel in $Channels) {
            $body = @{attachments = @(@{color="danger" ; text = $Text ; fallback = $Text}); channel = $Channel} | ConvertTo-Json
            $result = Invoke-RestMethod -Method Post -Uri $WebhookUri -Body $body
            if($result -ne "ok") {
                Write-Error "Expected 'ok' back from slack, but got: $result"
            }
        }
    }
    

}

Function Write-SlackPrivate {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   Position=0)]
        [String] $Text = 'Finished',

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=1)]
        [String[]] $Channels = @('#general'),

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=2)]
        [String] $WebhookUri = 'https://hooks.slack.com/services/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'
    )

    Process
    {
        foreach($Channel in $Channels) {
            $body = @{attachments = @(@{color="good" ; text = $Text ; fallback = $Text}); channel = $Channel} | ConvertTo-Json
            $result = Invoke-RestMethod -Method Post -Uri $WebhookUri -Body $body
            if($result -ne "ok") {
                Write-Error "Expected 'ok' back from slack, but got: $result"
            }
        }
    }
    

}

Function Write-SlackWarningPrivate {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   Position=0)]
        [String] $Text = 'Finished',

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=1)]
        [String[]] $Channels = @('#general'),

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=2)]
        [String] $WebhookUri = 'https://hooks.slack.com/services/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'
    )

    Process
    {
        foreach($Channel in $Channels) {
            $body = @{attachments = @(@{color="warning" ; text = $Text ; fallback = $Text}); channel = $Channel} | ConvertTo-Json
            $result = Invoke-RestMethod -Method Post -Uri $WebhookUri -Body $body
            if($result -ne "ok") {
                Write-Error "Expected 'ok' back from slack, but got: $result"
            }
        }
    }
    

}

