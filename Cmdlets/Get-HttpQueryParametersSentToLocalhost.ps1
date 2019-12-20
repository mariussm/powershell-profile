
function Get-HttpQueryParametersSentToLocalhost
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [int] $Port = 8080,

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [string] $Response = "Done"
    )

    $listener = New-Object System.Net.HttpListener
    $listener.Prefixes.Add("http://localhost:$Port/")
    Write-verbose "Waiting for request at http://localhost:$Port/"
    $listener.Start()
    $context = $listener.GetContext()
    $Content = [System.Text.Encoding]::UTF8.GetBytes($Response)
    $Context.Response.OutputStream.Write($Content, 0, $Content.Length)
    $Context.Response.Close()
    $listener.Dispose()
    $Context.Request.RawUrl -split "[?&]" -like "*=*" | foreach -Begin {$h = @{}} -Process {$h[($_ -split "=",2 | select -index 0)] = ($_ -split "=",2 | select -index 1)} -End {$h}
    
}