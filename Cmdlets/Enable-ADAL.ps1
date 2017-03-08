Function Enable-ADAL {
    [CmdletBinding()]
    Param
    ()

    mkdir HKCU:\Software\Microsoft\Office\15.0\Common\Identity\ -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
    Set-ItemProperty -Path HKCU:\Software\Microsoft\Office\15.0\Common\Identity\ -Name EnableADAL -Type DWord -Value 1
    mkdir HKCU:\Software\Microsoft\Office\15.0\Common\Debug -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
    Set-ItemProperty -Path HKCU:\Software\Microsoft\Office\15.0\Common\Debug -Name TCOTrace -Type DWord -Value 3

}
