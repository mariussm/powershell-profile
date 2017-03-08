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
