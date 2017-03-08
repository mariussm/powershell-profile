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
