# Set the PowerShell prompt to PS>
function prompt{
    Write-Host -ForegroundColor White -NoNewline ($env:COMPUTERNAME).ToUpper()
    Write-Host -ForegroundColor Red " PS" -NoNewline
    Write-Host -ForegroundColor White -NoNewline ">"
    return " "
}