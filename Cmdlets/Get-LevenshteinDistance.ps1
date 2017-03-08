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
