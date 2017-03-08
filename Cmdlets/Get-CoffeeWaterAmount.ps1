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
