Function Disable-ADAL {
    [CmdletBinding()]
    Param
    ()

    Set-ItemProperty -Path HKCU:\Software\Microsoft\Office\15.0\Common\Identity\ -Name EnableADAL -Type DWord -Value 0

}
