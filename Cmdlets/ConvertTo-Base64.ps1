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
