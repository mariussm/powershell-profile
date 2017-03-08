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
