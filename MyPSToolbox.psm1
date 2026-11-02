

function Invoke-SQL {
    <#
        .SYNOPSIS
            Invoke-SQL is an alternative to Invoke-Sqlcmd that does not require the SQLServer module to be installed. It uses the .NET SQL client.

        .DESCRIPTION
            Invoke-SQL is an alternative to Invoke-Sqlcmd that does not require the SQLServer module. It uses the .NET SQL client to query SQL Server.

            Selecting -Verbose will display safe connection metadata.

        .PARAMETER ServerInstance
            The Microsoft SQL Server host or instance to query.

        .PARAMETER TrustServerCertificate
            Enables accepting a server certificate that is not trusted locally. Leave
            this switch off when certificate validation is required.

        .PARAMETER [PSCredential]Credential
            Specifies a user account that has permission to perform this action. The default
            is the current user with Integrated Security.

        .PARAMETER Database
            The database to query.  The default is the current database.

        .PARAMETER ConnectionTimeout
            The number of seconds before a connection is considered timed out. The default is 10 seconds.
            A value of 0 leaves the provider default behavior in effect.

        .PARAMETER CommandTimeout
            The number of seconds before a query is considered timed out. The default is 720 seconds.

        .PARAMETER Query
            The query text to submit to SQL Server.

        .PARAMETER QueryParameters
            Values for named SQL parameters. Use simple values such as
            @{ SearchName = '%Backup%' }, or metadata descriptors such as
            @{ SearchName = @{ Value = '%Backup%'; SqlDbType = 'NVarChar'; Size = 128 } }.
            Metadata supports Value, SqlDbType, Size, Precision, and Scale.

        .PARAMETER AdditionalText
            Additional non-security connection-string options. Supported keys are
            Application Name, Application Intent, Pooling, Max Pool Size, Min Pool Size,
            and Load Balance Timeout.

        .PARAMETER ApplicationName
            The application name sent to SQL Server.

        .PARAMETER ApplicationIntent
            The application intent: ReadWrite or ReadOnly.

        .PARAMETER Pooling
            Enables or disables connection pooling.

        .PARAMETER MaxPoolSize
            The maximum number of connections in the pool.

        .PARAMETER MinPoolSize
            The minimum number of connections in the pool.

        .PARAMETER LoadBalanceTimeout
            The number of seconds before pooled connections are removed.

        .EXAMPLE
            # Using Microsoft SQL Server credentials
                $cred = Get-Credential
                Invoke-SQL -ServerInstance "localhost" -Credential $cred -Query "SELECT @@servername as [Server]" -Verbose

        .INPUTS
            ServerInstance
            Query
            QueryParameters
            Database
            ConnectionTimeout
            CommandTimeout
            Credential
            TrustServerCertificate
            AdditionalText
            ApplicationName
            ApplicationIntent
            Pooling
            MaxPoolSize
            MinPoolSize
            LoadBalanceTimeout

        .OUTPUTS
            System.Data.DataTable

        .NOTES
            Author:  David Allen
            Email: dallenva@gmail.com
        #>
        [CmdletBinding()]
        Param([parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$ServerInstance
                ,[parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Query
                ,[ValidateNotNull()][hashtable]$QueryParameters = @{}
                ,[ValidateScript({
                    if([string]::IsNullOrWhiteSpace($_)) {
                        throw 'Database cannot be empty or whitespace.'
                    }
                    $true
                })][string]$Database
                ,[ValidateRange(0,2147483647)][int]$ConnectionTimeout=10
                ,[ValidateRange(0,2147483647)][int]$CommandTimeout=720
                ,[PSCredential]$Credential = $null
                ,[switch]$TrustServerCertificate
                ,[string]$AdditionalText
                ,[string]$ApplicationName
                ,[ValidateSet('ReadWrite','ReadOnly')][string]$ApplicationIntent
                ,[bool]$Pooling
                ,[ValidateRange(0,2147483647)][int]$MaxPoolSize
                ,[ValidateRange(0,2147483647)][int]$MinPoolSize
                ,[ValidateRange(0,2147483647)][int]$LoadBalanceTimeout)
        Process {
            if($PSBoundParameters.ContainsKey('MinPoolSize') -and
                $PSBoundParameters.ContainsKey('MaxPoolSize') -and
                $MinPoolSize -gt $MaxPoolSize) {
                throw 'MinPoolSize cannot be greater than MaxPoolSize.'
            }
    # Build Connection
            $conn = [System.Data.SqlClient.SqlConnection]::new()
            $ExecSQL = $null
            $Results = $null
    # Build Connection String
            $ConnectionStringBuilder = [System.Data.SqlClient.SqlConnectionStringBuilder]::new()
            $ConnectionStringBuilder['Data Source'] = $ServerInstance
            If($Database.Length -gt 0) {$ConnectionStringBuilder['Initial Catalog'] = $Database}
            If($ConnectionTimeout -gt 0) {$ConnectionStringBuilder['Connect Timeout'] = $ConnectionTimeout}
            If($Credential) {
                $ConnectionStringBuilder['Integrated Security'] = $false
                $ConnectionStringBuilder['User ID'] = $Credential.UserName
                $ConnectionStringBuilder['Password'] = $Credential.GetNetworkCredential().Password
            } else {
                $ConnectionStringBuilder['Integrated Security'] = $true
            }
            $ConnectionStringBuilder['Encrypt'] = $true
            $ConnectionStringBuilder['TrustServerCertificate'] = $TrustServerCertificate.IsPresent
            if($AdditionalText) {
                try {
                    $AdditionalConnectionStringBuilder = [System.Data.SqlClient.SqlConnectionStringBuilder]::new($AdditionalText)
                } catch {
                    throw "AdditionalText is not a valid connection-string fragment: $($_.Exception.Message)"
                }
                $AllowedAdditionalKeys = @(
                    'Application Name',
                    'Application Intent',
                    'Pooling',
                    'Max Pool Size',
                    'Min Pool Size',
                    'Load Balance Timeout'
                )
                foreach($Key in $AdditionalConnectionStringBuilder.Keys) {
                    if($Key -notin $AllowedAdditionalKeys) {
                        throw "AdditionalText contains unsupported connection-string key '$Key'. Use an explicit parameter instead."
                    }
                    $ConnectionStringBuilder[$Key] = $AdditionalConnectionStringBuilder[$Key]
                }
            }
            if($PSBoundParameters.ContainsKey('ApplicationName')) {$ConnectionStringBuilder['Application Name'] = $ApplicationName}
            if($PSBoundParameters.ContainsKey('ApplicationIntent')) {$ConnectionStringBuilder['ApplicationIntent'] = $ApplicationIntent}
            if($PSBoundParameters.ContainsKey('Pooling')) {$ConnectionStringBuilder['Pooling'] = $Pooling}
            if($PSBoundParameters.ContainsKey('MaxPoolSize')) {$ConnectionStringBuilder['Max Pool Size'] = $MaxPoolSize}
            if($PSBoundParameters.ContainsKey('MinPoolSize')) {$ConnectionStringBuilder['Min Pool Size'] = $MinPoolSize}
            if($PSBoundParameters.ContainsKey('LoadBalanceTimeout')) {$ConnectionStringBuilder['Load Balance Timeout'] = $LoadBalanceTimeout}
            $ConnectionString = $ConnectionStringBuilder.ConnectionString
            Write-Verbose ("SQL connection target: Server={0}; Database={1}; Encrypt={2}; TrustServerCertificate={3}" -f
                $ConnectionStringBuilder['Data Source'],
                $ConnectionStringBuilder['Initial Catalog'],
                $ConnectionStringBuilder['Encrypt'],
                $ConnectionStringBuilder['TrustServerCertificate'])
            foreach ($ParameterName in $QueryParameters.Keys) {
                if($ParameterName -notmatch '^[A-Za-z_][A-Za-z0-9_]*$') {
                    throw "Invalid SQL parameter name: $ParameterName"
                }
                $ParameterSpec = $QueryParameters[$ParameterName]
                $ParsedSqlDbType = $null
                if($ParameterSpec -is [hashtable]) {
                    $AllowedMetadataKeys = @('Value','SqlDbType','Size','Precision','Scale')
                    $UnexpectedMetadataKeys = @($ParameterSpec.Keys | Where-Object {$_ -notin $AllowedMetadataKeys})
                    if($UnexpectedMetadataKeys.Count -gt 0) {
                        throw "Unsupported SQL parameter metadata key for '$ParameterName': $($UnexpectedMetadataKeys -join ', ')"
                    }
                    if(-not $ParameterSpec.ContainsKey('Value')) {
                        throw "SQL parameter metadata for '$ParameterName' must include a Value."
                    }
                    if($ParameterSpec.ContainsKey('SqlDbType')) {
                        try {
                            $ParsedSqlDbType = [System.Enum]::Parse([System.Data.SqlDbType], [string]$ParameterSpec.SqlDbType, $true)
                        } catch {
                            throw "Invalid SqlDbType for SQL parameter '$ParameterName': $($ParameterSpec.SqlDbType)"
                        }
                    }
                    if($ParameterSpec.ContainsKey('Size') -and
                        ($ParameterSpec.Size -isnot [int] -or $ParameterSpec.Size -lt -1 -or $ParameterSpec.Size -eq 0)) {
                        throw "Invalid Size for SQL parameter '$ParameterName': $($ParameterSpec.Size)"
                    }
                    if($ParameterSpec.ContainsKey('Precision') -and
                        ($ParameterSpec.Precision -isnot [byte] -and $ParameterSpec.Precision -isnot [int] -or
                        $ParameterSpec.Precision -lt 1 -or $ParameterSpec.Precision -gt 38)) {
                        throw "Invalid Precision for SQL parameter '$ParameterName': $($ParameterSpec.Precision)"
                    }
                    if($ParameterSpec.ContainsKey('Scale') -and
                        ($ParameterSpec.Scale -isnot [byte] -and $ParameterSpec.Scale -isnot [int] -or
                        $ParameterSpec.Scale -lt 0 -or
                        ($ParameterSpec.ContainsKey('Precision') -and $ParameterSpec.Scale -gt $ParameterSpec.Precision))) {
                        throw "Invalid Scale for SQL parameter '$ParameterName': $($ParameterSpec.Scale)"
                    }
                    if($ParameterSpec.ContainsKey('Precision') -xor $ParameterSpec.ContainsKey('Scale')) {
                        throw "Precision and Scale must be provided together for SQL parameter '$ParameterName'."
                    }
                    if(($ParameterSpec.ContainsKey('Precision') -or $ParameterSpec.ContainsKey('Scale')) -and
                        $ParsedSqlDbType -notin @([System.Data.SqlDbType]::Decimal, [System.Data.SqlDbType]::Numeric, [System.Data.SqlDbType]::Time)) {
                        throw "Precision and Scale are not supported for SQL parameter '$ParameterName' with type $ParsedSqlDbType."
                    }
                    if($ParameterSpec.ContainsKey('Size') -and
                        $ParsedSqlDbType -notin @(
                            [System.Data.SqlDbType]::Char,
                            [System.Data.SqlDbType]::NChar,
                            [System.Data.SqlDbType]::VarChar,
                            [System.Data.SqlDbType]::NVarChar,
                            [System.Data.SqlDbType]::Text,
                            [System.Data.SqlDbType]::NText,
                            [System.Data.SqlDbType]::Binary,
                            [System.Data.SqlDbType]::VarBinary,
                            [System.Data.SqlDbType]::Image
                        )) {
                        throw "Size is not supported for SQL parameter '$ParameterName' with type $ParsedSqlDbType."
                    }
                }
                $Value = if($ParameterSpec -is [hashtable]) {$ParameterSpec.Value} else {$ParameterSpec}
                if($null -ne $Value -and $Value.GetType().FullName -notin @(
                    'System.String',
                    'System.Int32',
                    'System.Int64',
                    'System.Decimal',
                    'System.Double',
                    'System.Boolean',
                    'System.DateTime',
                    'System.Guid',
                    'System.Byte[]'
                )) {
                    throw "Unsupported SQL parameter value type for '$ParameterName': $($Value.GetType().FullName)"
                }
            }
    # Set Connection String and execute the query
            try {
                $conn.ConnectionString = $ConnectionString
                $conn.Open()

    # If successful open connection, run query
                if($conn.state -eq 'Open'){
                    Write-Verbose "----------------------------------------------------------"
                    Write-Verbose "Open() successful.... Connection Details:"
                    Write-Verbose ("Query Started: " + (Get-Date))
                    write-verbose ("Connection established: Server={0}; Database={1}; Encrypted={2}" -f
                        $conn.DataSource,
                        $conn.Database,
                        $ConnectionStringBuilder['Encrypt'])
                    write-verbose ("Database: " + $conn.Database)
                    write-verbose ("DataSource: " + $conn.DataSource)
                    write-verbose ("ServerVersion: " + $conn.ServerVersion)
                    write-verbose ("WorkstationId: " + $conn.WorkstationId)
                    $time = Measure-Command{$ExecSQL = [System.Data.SqlClient.SqlCommand]::new($Query, $conn)
                                            $ExecSQL.CommandTimeout = $CommandTimeout
                                            foreach ($ParameterName in $QueryParameters.Keys) {
                                                $Value = $QueryParameters[$ParameterName]
                                                $ParameterSpec = $Value
                                                if($ParameterSpec -is [hashtable]) {
                                                    $Value = $ParameterSpec.Value
                                                }
                                                $ExplicitSqlDbType = $null
                                                if($ParameterSpec -is [hashtable] -and $ParameterSpec.ContainsKey('SqlDbType')) {
                                                    $ExplicitSqlDbType = [System.Enum]::Parse([System.Data.SqlDbType], [string]$ParameterSpec.SqlDbType, $true)
                                                }
                                                if($null -ne $ExplicitSqlDbType) {
                                                    $Parameter = $ExecSQL.Parameters.Add("@$ParameterName", $ExplicitSqlDbType)
                                                    if($ParameterSpec.ContainsKey('Size')) {$Parameter.Size = $ParameterSpec.Size}
                                                    if($ExplicitSqlDbType -in @([System.Data.SqlDbType]::Decimal, [System.Data.SqlDbType]::Numeric)) {
                                                        if($ParameterSpec.ContainsKey('Precision')) {
                                                            $Parameter.Precision = $ParameterSpec.Precision
                                                            $Parameter.Scale = $ParameterSpec.Scale
                                                        } elseif($Value -is [decimal]) {
                                                            $DecimalBits = [decimal]::GetBits($Value)
                                                            $Parameter.Scale = [byte](($DecimalBits[3] -shr 16) -band 0x7f)
                                                            $Digits = $Value.ToString([Globalization.CultureInfo]::InvariantCulture).Replace('.', '').TrimStart('-').TrimStart('0')
                                                            $Parameter.Precision = [byte][Math]::Max(1, $Digits.Length)
                                                        } else {
                                                            $Parameter.Precision = 18
                                                            $Parameter.Scale = 0
                                                        }
                                                    } elseif($ParameterSpec.ContainsKey('Precision') -or $ParameterSpec.ContainsKey('Scale')) {
                                                        $Parameter.Precision = $ParameterSpec.Precision
                                                        $Parameter.Scale = $ParameterSpec.Scale
                                                    }
                                                    $Parameter.Value = if($null -eq $Value) {[DBNull]::Value} else {$Value}
                                                    continue
                                                }
                                                if($null -eq $Value) {
                                                    $Parameter = $ExecSQL.Parameters.Add("@$ParameterName", [System.Data.SqlDbType]::NVarChar, 4000)
                                                    $Parameter.Value = [DBNull]::Value
                                                    continue
                                                }

                                                switch($Value.GetType().FullName) {
                                                    'System.Int32' {
                                                        $Parameter = $ExecSQL.Parameters.Add("@$ParameterName", [System.Data.SqlDbType]::Int)
                                                    }
                                                    'System.Int64' {
                                                        $Parameter = $ExecSQL.Parameters.Add("@$ParameterName", [System.Data.SqlDbType]::BigInt)
                                                    }
                                                    'System.Decimal' {
                                                        $Parameter = $ExecSQL.Parameters.Add("@$ParameterName", [System.Data.SqlDbType]::Decimal)
                                                        $DecimalBits = [decimal]::GetBits($Value)
                                                        $Parameter.Scale = [byte](($DecimalBits[3] -shr 16) -band 0x7f)
                                                        $Digits = $Value.ToString([Globalization.CultureInfo]::InvariantCulture).Replace('.', '').TrimStart('-').TrimStart('0')
                                                        $Parameter.Precision = [byte][Math]::Max(1, $Digits.Length)
                                                    }
                                                    'System.Double' {
                                                        $Parameter = $ExecSQL.Parameters.Add("@$ParameterName", [System.Data.SqlDbType]::Float)
                                                    }
                                                    'System.Boolean' {
                                                        $Parameter = $ExecSQL.Parameters.Add("@$ParameterName", [System.Data.SqlDbType]::Bit)
                                                    }
                                                    'System.DateTime' {
                                                        $Parameter = $ExecSQL.Parameters.Add("@$ParameterName", [System.Data.SqlDbType]::DateTime2)
                                                    }
                                                    'System.Guid' {
                                                        $Parameter = $ExecSQL.Parameters.Add("@$ParameterName", [System.Data.SqlDbType]::UniqueIdentifier)
                                                    }
                                                    'System.Byte[]' {
                                                        $Parameter = $ExecSQL.Parameters.Add("@$ParameterName", [System.Data.SqlDbType]::VarBinary, -1)
                                                    }
                                                    default {
                                                        $Parameter = $ExecSQL.Parameters.Add("@$ParameterName", [System.Data.SqlDbType]::NVarChar, 4000)
                                                    }
                                                }
                                                $Parameter.Value = $Value
                                            }
    # Convert result to table
                                            $Results = [System.Data.SqlClient.SqlDataAdapter]::new($ExecSQL)
                                            $Data = [System.Data.DataSet]::new()
                                            [void]$Results.Fill($Data)}
                    $Info = "Elapsed Seconds: " + $time.TotalSeconds
                    Write-Verbose ("Query Ended: " + (Get-Date))
                    Write-Verbose $Info
                    $Info = "Record Count: " + $Data.Tables[0].Rows.Count
                    Write-Verbose $Info
                    Write-Verbose "-- End Invoke-SQL ----------------------------------------"
    # Return Results
                    Return $Data.Tables[0]
                }
            }
            finally {
                if($Results) {$Results.Dispose()}
                if($ExecSQL) {$ExecSQL.Dispose()}
                if($conn) {
                    if($conn.State -ne 'Closed') {$conn.Close()}
                    $conn.Dispose()
                }
            }
        }
    }

function Write-SQLBulkCopy {
<#
.SYNOPSIS
    Bulk-copy PowerShell objects into an existing SQL Server table.

.DESCRIPTION
    Streams an in-memory object collection to an existing SQL Server table by
    using the .NET SqlBulkCopy API. Source property names are matched to target
    column names by default. Use ColumnMapping when names differ.

    The destination table must already exist, and the source objects should
    have a consistent set of properties. BatchSize controls how many rows are
    sent per batch; it does not increase the database service tier's log-rate
    limit.

.PARAMETER InputObject
    Objects to upload. Accepts pipeline input.

.PARAMETER ServerInstance
    The Microsoft SQL Server host or instance to query.

.PARAMETER DestinationTable
    Existing destination table, optionally schema-qualified, such as
    dbo.ImportRows.

.PARAMETER Database
    The database containing the destination table.

.PARAMETER Credential
    Credentials used to connect. If omitted, integrated security is used.

.PARAMETER TrustServerCertificate
    Accepts a server certificate that is not trusted locally. Leave this
    switch off when certificate validation is required.

.PARAMETER BatchSize
    Number of rows sent in each batch. Defaults to 5,000.

.PARAMETER BulkCopyTimeout
    Number of seconds before the bulk copy times out. A value of 0 means no
    timeout. Defaults to 720 seconds.

.PARAMETER ColumnMapping
    Optional hashtable mapping source property names to target column names,
    for example @{ SourceName = 'TargetName' }.

.PARAMETER ShowProgress
    Displays a progress bar while rows are copied.

.PARAMETER ProgressInterval
    Number of copied rows between progress updates. Defaults to 1,000.
    This parameter is used only when ShowProgress is specified.

.EXAMPLE
    Get-Process | Select-Object Name, Id | Write-SQLBulkCopy `
        -ServerInstance localhost -Database Operations `
        -DestinationTable dbo.ProcessSnapshot -Credential (Get-Credential)

.EXAMPLE
    $rows | Write-SQLBulkCopy -ServerInstance localhost `
        -Database Operations -DestinationTable dbo.ImportRows `
        -ColumnMapping @{ ExternalId = 'Id'; DisplayName = 'Name' } `
        -ShowProgress

.INPUTS
    System.Object

.OUTPUTS
    PSCustomObject with DestinationTable, RowsCopied, and ElapsedSeconds.

.NOTES
    Author:  David Allen
    Email: dallenva@gmail.com
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [PSObject]$InputObject,
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ServerInstance,
        [Parameter(Mandatory)]
        [ValidatePattern('^(?:\[[^\]]+\]|[A-Za-z_][A-Za-z0-9_$#@]*)(?:\.(?:\[[^\]]+\]|[A-Za-z_][A-Za-z0-9_$#@]*))?$')]
        [string]$DestinationTable,
        [Parameter(Mandatory)]
        [ValidateScript({
            if ([string]::IsNullOrWhiteSpace($_)) {
                throw 'Database cannot be empty or whitespace.'
            }
            $true
        })]
        [string]$Database,
        [PSCredential]$Credential,
        [switch]$TrustServerCertificate,
        [ValidateRange(1, 2147483647)]
        [int]$BatchSize = 5000,
        [ValidateRange(0, 2147483647)]
        [int]$BulkCopyTimeout = 720,
        [ValidateNotNull()]
        [hashtable]$ColumnMapping = @{},
        [switch]$ShowProgress,
        [ValidateRange(1, 2147483647)]
        [int]$ProgressInterval = 1000
    )
    begin {
        $Rows = [System.Collections.Generic.List[object]]::new()
    }
    process {
        [void]$Rows.Add($InputObject)
    }
    end {
        if ($Rows.Count -eq 0) {
            throw 'InputObject did not contain any rows.'
        }

        $FirstRow = $Rows[0]
        $Properties = @($FirstRow.PSObject.Properties | Where-Object {
            $_.MemberType -in @('Property', 'NoteProperty')
        })
        if ($Properties.Count -eq 0) {
            throw 'InputObject must expose at least one property.'
        }

        $DataTable = [System.Data.DataTable]::new()
        foreach ($Property in $Properties) {
            [void]$DataTable.Columns.Add($Property.Name, [object])
        }

        foreach ($Row in $Rows) {
            $DataRow = $DataTable.NewRow()
            foreach ($Property in $Properties) {
                $SourceProperty = $Row.PSObject.Properties[$Property.Name]
                $DataRow[$Property.Name] = if ($null -eq $SourceProperty -or $null -eq $SourceProperty.Value) {
                    [DBNull]::Value
                } else {
                    $SourceProperty.Value
                }
            }
            [void]$DataTable.Rows.Add($DataRow)
        }

        foreach ($SourceName in $ColumnMapping.Keys) {
            if ($SourceName -notin $DataTable.Columns.ColumnName) {
                throw "ColumnMapping source '$SourceName' was not found on the input object."
            }
            if ([string]::IsNullOrWhiteSpace([string]$ColumnMapping[$SourceName])) {
                throw "ColumnMapping target for '$SourceName' cannot be empty."
            }
        }

        $ConnectionStringBuilder = [System.Data.SqlClient.SqlConnectionStringBuilder]::new()
        $ConnectionStringBuilder['Data Source'] = $ServerInstance
        $ConnectionStringBuilder['Initial Catalog'] = $Database
        if ($Credential) {
            $ConnectionStringBuilder['Integrated Security'] = $false
            $ConnectionStringBuilder['User ID'] = $Credential.UserName
            $ConnectionStringBuilder['Password'] = $Credential.GetNetworkCredential().Password
        } else {
            $ConnectionStringBuilder['Integrated Security'] = $true
        }
        $ConnectionStringBuilder['Encrypt'] = $true
        $ConnectionStringBuilder['TrustServerCertificate'] = $TrustServerCertificate.IsPresent

        $Connection = [System.Data.SqlClient.SqlConnection]::new($ConnectionStringBuilder.ConnectionString)
        $BulkCopy = $null
        try {
            Write-Verbose ("Bulk-copy target: Server={0}; Database={1}; Table={2}; Rows={3}; BatchSize={4}" -f
                $ServerInstance, $Database, $DestinationTable, $DataTable.Rows.Count, $BatchSize)
            $Connection.Open()
            $BulkCopy = [System.Data.SqlClient.SqlBulkCopy]::new($Connection)
            $BulkCopy.DestinationTableName = $DestinationTable
            $BulkCopy.BatchSize = $BatchSize
            $BulkCopy.BulkCopyTimeout = $BulkCopyTimeout
            $BulkCopy.EnableStreaming = $true
            $ProgressId = 1
            $RowsCopiedEventHandler = $null
            if ($ShowProgress.IsPresent) {
                $BulkCopy.NotifyAfter = [Math]::Min($ProgressInterval, $DataTable.Rows.Count)
                $RowsCopiedEventHandler = [System.Data.SqlClient.SqlRowsCopiedEventHandler]{
                    param($Sender, $EventArgs)
                    $PercentComplete = [Math]::Min(100, [Math]::Round(($EventArgs.RowsCopied / $DataTable.Rows.Count) * 100, 1))
                    Write-Progress -Id $ProgressId -Activity "Copying rows to $DestinationTable" `
                        -Status "$($EventArgs.RowsCopied) of $($DataTable.Rows.Count) rows copied" `
                        -PercentComplete $PercentComplete
                }
                $BulkCopy.add_SqlRowsCopied($RowsCopiedEventHandler)
            }
            if ($ColumnMapping.Count -gt 0) {
                foreach ($SourceName in $ColumnMapping.Keys) {
                    [void]$BulkCopy.ColumnMappings.Add($SourceName, [string]$ColumnMapping[$SourceName])
                }
            } else {
                foreach ($Property in $Properties) {
                    [void]$BulkCopy.ColumnMappings.Add($Property.Name, $Property.Name)
                }
            }
            $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            $BulkCopy.WriteToServer($DataTable)
            $Stopwatch.Stop()
            if ($ShowProgress.IsPresent) {
                Write-Progress -Id $ProgressId -Activity "Copying rows to $DestinationTable" `
                    -Status "$($DataTable.Rows.Count) of $($DataTable.Rows.Count) rows copied" `
                    -Completed
            }
            [PSCustomObject]@{
                DestinationTable = $DestinationTable
                RowsCopied       = $DataTable.Rows.Count
                ElapsedSeconds   = [Math]::Round($Stopwatch.Elapsed.TotalSeconds, 3)
            }
        } catch {
            Write-Error -ErrorRecord $_
            throw
        } finally {
            if ($BulkCopy -and $RowsCopiedEventHandler) {
                $BulkCopy.remove_SqlRowsCopied($RowsCopiedEventHandler)
            }
            if ($BulkCopy) { $BulkCopy.Close(); $BulkCopy.Dispose() }
            if ($Connection.State -ne [System.Data.ConnectionState]::Closed) { $Connection.Close() }
            $Connection.Dispose()
            $DataTable.Dispose()
        }
    }
}

function Get-EmptyDirectory {
    <#
    .SYNOPSIS
        Get empty directories using underlying Get-ChildItem cmdlet

    .NOTES
        Name: Get-EmptyDirectory
        Author: theSysadminChannel
        Version: 1.0
        DateCreated: 2021-Oct-2
        Modified David Allen 2024

    .PARAMETER Path
        Root path to inspect.

    .PARAMETER Recurse
        Recursively inspect child directories.

    .PARAMETER Depth
        Maximum recursion depth, from 1 through 15.

    .OUTPUTS
        PSCustomObject with EmptyDirectory and Path properties.
    .LINK
        https://thesysadminchannel.com/find-empty-folders-powershell/ -

    .EXAMPLE
        Get-EmptyDirectory -Path \\Server\Share\Folder -Depth 2
    #>
        [CmdletBinding()]
        param(
            [Parameter(
                Mandatory = $true,
                Position = 0
            )]
            [string]    $Path,

            [Parameter(
                Mandatory = $false,
                Position = 1
            )]
            [switch]    $Recurse,

            [Parameter(
                Mandatory = $false,
                Position = 2
            )]
            [ValidateRange(1,15)]
            [int]    $Depth
        )
        BEGIN {}
        PROCESS {
            try {
                $ItemParams = @{
                    Path      = $Path
                    Directory = $true
                }
                if ($Recurse.IsPresent) {
                    $ItemParams.Add('Recurse',$true)
                }

                if ($PSBoundParameters.ContainsKey('Depth')) {
                    $ItemParams.Add('Depth',$Depth)
                }
                $FolderList = Get-ChildItem @ItemParams

                $FolderList | ForEach-Object{
                    if ((Get-ChildItem -LiteralPath $_.FullName).count -lt 1) {
                        [PSCustomObject]@{
                            EmptyDirectory = $true
                            Path           = $_.FullName
                        }
                    } else {}
                }
            } catch {
                Write-Error -ErrorRecord $_
                throw
            }
        }
        END {}
    }
function Set-WindowTitle {
<#
.SYNOPSIS
    Change the PowerShell window title

.PARAMETER Title
    The text to display in the PowerShell window title.

.EXAMPLE
    set-WindowTitle -title "My Session"

.INPUTS
    None. Values are supplied through parameters.

.OUTPUTS
    None.
#>
    Param([string]$Title)
    Process {
        $host.ui.RawUI.WindowTitle = $Title
    }
}
function Set-Cursor {
<#
.SYNOPSIS
    Set the X and Y position of the cursor

.PARAMETER X
    The horizontal cursor position.

.PARAMETER Y
    The vertical cursor position.

.NOTES
    Author:  David Allen
    Email: dallenva@gmail.com

.EXAMPLE
    Set-Cursor -X 0 -Y 0

.INPUTS
    None. Values are supplied through parameters.

.OUTPUTS
    None.
#>
Param([int]$x, [int] $y)
Process {
    $Host.UI.RawUI.CursorPosition = [System.Management.Automation.Host.Coordinates]::new($x, $y)
}
}
function Set-ConsoleWindowSize
{
<#
.SYNOPSIS
    Set Width and Height of the console window.  This is dependent on the terminal, VS Code and Windows Terminal
    will not allow this to override the setting.

.PARAMETER Width
    The requested console width in columns.

.PARAMETER Height
    The requested console height in rows.

.NOTES
    Author:  David Allen
    Email: dallenva@gmail.com

.EXAMPLE
    Set-ConsoleWindowSize -Width 80 -Height 24

.INPUTS
    None. Values are supplied through parameters.

.OUTPUTS
    None.
#>
    param(
        [int]$Width,
        [int]$Height
        )
    $WindowSize = $Host.UI.RawUI.WindowSize
    $WindowSize.Width  = [Math]::Min($Width, $Host.UI.RawUI.BufferSize.Width)
    $WindowSize.Height = $Height
    try{$Host.UI.RawUI.WindowSize = $WindowSize}
    catch [System.Management.Automation.SetValueInvocationException] {
        $MaxValue = ($_.Exception.Message |Select-String "\d+").Matches[0].Value
        $WindowSize.Height = $MaxValue
        $Host.UI.RawUI.WindowSize = $WindowSize
    }
}
function Lock-Screen {
<#
.SYNOPSIS
    Lock the current workstation.

.DESCRIPTION
    Uses the host operating system's native screen-lock command. The command
    is supported on Windows, macOS, and Linux environments with the required
    platform utility installed.

.INPUTS
    None.

.OUTPUTS
    None.
#>
    Param()
    Process {
            if($IsMacOS) {
                if (-not (Get-Command pmset -ErrorAction SilentlyContinue)) {
                    throw 'The pmset command is required to lock the screen on macOS.'
                }
                & pmset sleepnow
            } elseif($IsLinux) {
                if (-not (Get-Command gnome-screensaver-command -ErrorAction SilentlyContinue)) {
                    throw 'The gnome-screensaver-command utility is required to lock the screen on Linux.'
                }
                & gnome-screensaver-command -l
            } elseif($IsWindows) {
                & rundll32.exe user32.dll,LockWorkStation
            } else {
                throw 'The current operating system is not supported.'
            }
        }
    }
function Get-UI {
<#
.SYNOPSIS
    Return the current PowerShell host UI object.

.INPUTS
    None.

.OUTPUTS
    System.Management.Automation.Host.PSHostRawUserInterface
#>
    Process {
        $host.UI.RawUI
    }
}
function get-whoami {
<#
.SYNOPSIS
    Display PowerShell, operating system, host, and user information.

.DESCRIPTION
    Writes a human-readable summary of the current PowerShell session to the
    host. This function writes directly to the host rather than returning an
    object.

.INPUTS
    None.

.OUTPUTS
    None. Information is written to the host.
#>
    Write-Host "====================================="
    Write-Host "PowerShell Edition: " -NoNewline -ForegroundColor Green
    Write-Host $PSVersionTable.PSEdition -ForegroundColor Cyan
    Write-Host "PowerShell Version: " -NoNewline -ForegroundColor Green
    Write-Host $PSVersionTable.PSVersion -ForegroundColor Cyan
    $os = if($IsWindows -eq $true){"Windows"}else{if($IsMacOS -eq $true){"MacOS"}else{"Linux"}}
    Write-Host "OS: " -NoNewline -ForegroundColor Green
    Write-Host $os -ForegroundColor Cyan
    Write-Host "Computer: " -NoNewline -ForegroundColor Green
    if ($os -eq "Windows"){$ComputerName = $env:COMPUTERNAME}else{$ComputerName = $(hostname)}
    Write-Host $ComputerName -ForegroundColor Cyan
    Write-Host "User Account: " -NoNewline -ForegroundColor Green
    if ($os -eq "Windows"){$UserName = $env:Username}else{$UserName = (get-item -path env:USER).value}
    Write-Host $UserName -ForegroundColor Cyan
    Write-Host "Session Name: " -NoNewline -ForegroundColor Green
    Write-Host $env:SESSIONNAME -ForegroundColor Cyan
    Write-Host "User Domain: " -NoNewline -ForegroundColor Green
    Write-Host $env:USERDNSDOMAIN -ForegroundColor Cyan
    Write-Host "Home: " -NoNewline -ForegroundColor Green
    Write-Host $home -ForegroundColor Cyan
    Write-Host "Profile: " -NoNewline -ForegroundColor Green
    Write-Host $Profile -ForegroundColor Cyan
    Write-Host "====================================="
}
function Get-MyDrives {
<#
.SYNOPSIS
    Return file-system drives available to the current session.

.DESCRIPTION
    Returns PowerShell file-system provider drives and excludes the drive
    named Temp.

.INPUTS
    None.

.OUTPUTS
    System.Management.Automation.PSDriveInfo
#>
    $return = Get-PSDrive -PSProvider FileSystem
    $return | Where-Object {$_.Name -notlike 'Temp'}
}
Function Out-ScreenGrid {
    <#
    .SYNOPSIS
        This is an alternative to Format-Table. When an object is passed, its
        text is converted into a table of up to 6 columns with alternating colors.

    .DESCRIPTION
        This was created for a few specific use cases and may not be the best solution for your case.
    .PARAMETER InputObject
        Any System.Object value. Enclose it in parentheses when needed.
    .PARAMETER TableHeader
        String to appear before the table
    .PARAMETER TableFooter
        String to appear after the table.
    .PARAMETER NoFooter
        Suppress the table footer.
    .EXAMPLE
        Out-ScreenGrid -InputObject (Get-ChildItem | Select-Object Name, Extension, CreationTime, LastAccessTime)
    .EXAMPLE
        Out-ScreenGrid -InputObject (Get-Process)
    .INPUTS
        System.Object
        [string]TableHeader
        [string]TableFooter
        [switch]NoFooter
    .OUTPUTS
        None
    .NOTES
        Author:  David Allen
        Email: dallenva@gmail.com
    #>
    Param(
        [Parameter(Mandatory,ValueFromPipeline)][System.Object]$InputObject,
        [String]$TableHeader,
        [String]$TableFooter,
        [switch]$NoFooter
        )
    Process {
# Setup Variables
        $MyWidth = $host.UI.RawUI.WindowSize.Width
        $Reset = $PSStyle.Reset
        $OddLinesFormat = $PSStyle.Background.cyan + $PSStyle.Foreground.Black
        $EvenLinesFormat = $PSStyle.Background.brightCyan + $PSStyle.Foreground.Black
        $HeaderLineFormat = $PSStyle.Bold+$PSStyle.Underline+$PSStyle.Background.blue + $PSStyle.Foreground.white
        $MaxColumns = 5
<#
Create options for
        max lines
        Header format colors
        Odd Even format colors
        Paging
#>
    # Convert Object into String CSV to format
        $StringData = ($InputObject) | ConvertTo-CSV -delimiter "|" -QuoteFields ""
    # Convert Header row into an array
        $Header = $StringData[0].split("|")
    # Calc number of columns up to 5
        $ColumnNumber = if($Header.count -le $MaxColumns){$Header.count}else{$MaxColumns}
        $Offset = $ColumnNumber + 5
    # Calc Column width for formatting
        $ColumnWidth = [math]::Round(($MyWidth - $Offset) / $ColumnNumber)
    # Convert object into formatted table
        $RowIdx = 0 # Row Index
        $OutputString = $HeaderLineFormat+$TableHeader+"`r`n"+$Reset
        do{$ColumnIdx = 0
            ;$TempRowString = $null
            ;$StringDataLine = $StringData[$RowIdx].split("|")
    # Format Columns with spacing and "|"
            ;$TempRowString += do{"|" + (($StringDataLine[$ColumnIdx]).PadRight($ColumnWidth," ")).substring(0,$ColumnWidth)
                                    ;if($ColumnIdx -eq $MaxColumns){$ColumnIdx = $ColumnNumber}else{$ColumnIdx +=1}
                                }until($ColumnIdx -eq $ColumnNumber)
    # Format Header, Odd, Even colors
            ;if($RowIdx -eq 0){$TempRowString = $HeaderLineFormat+$TempRowString+$Reset}else{
                if (($RowIdx % 2) -ne 0) {$TempRowString = $OddLinesFormat+$TempRowString+$Reset}else
                    {$TempRowString = $EvenLinesFormat+$TempRowString+$Reset}}
    # The process removes line breaks and adds a single one at the end
            ;$OutputString += [string]::join("",($TempRowString.Split("`n"))) + "`r`n"
            ;$RowIdx += 1
        }until($RowIdx -eq $InputObject.count+1)
    # Send output string to the console
        Clear-Host
        Write-Host $OutputString
        If($TableFooter.Length -gt 0){Write-Host $HeaderLineFormat$TableFooter$Reset}else{
            if(-not $NoFooter.IsPresent){Write-Host $PSStyle.Formatting.Warning"<END>"$Reset}}
    }
}

Function Get-SQLSessionInfo {
    <#
    .SYNOPSIS
        Returns a list of sessions on a Microsoft SQL Server similar to the sp_who and sp_who2 commands.

    .DESCRIPTION
        Get-SQLSessionInfo is a function that returns a list of Microsoft SQL Server sessions similar to the sp_who or sp_who2
        commands.  This will return more information from the sessions tables than the built in commands.  Additional options
        will let you filter the results by SPID and LoginName.

        ServerInstance can also be passed in the pipeline.

        Selecting -Verbose displays safe connection and query details.

    .PARAMETER ServerInstance
        The Microsoft SQL Server host or instance to query.

    .PARAMETER TrustServerCertificate
        Enables accepting a server certificate that is not trusted locally. Leave
        this switch off when certificate validation is required.

    .PARAMETER Credential
        Specifies a user account that has permission to perform this action. The default
        is the current user with Integrated Security.

    .PARAMETER IncludeMySPID
        By default, the current session is excluded. Use -IncludeMySPID to include it.

    .PARAMETER SPID
        Using -SPID filters the results to one session ID.

    .PARAMETER SearchLoginName
        -SearchLoginName allows you to filter returned sessions by login name.

    .PARAMETER Status
        Using -Status filters results to one request status.

    .PARAMETER OnlyStatementText
        Selecting -OnlyStatementText returns only sessions with non-empty statement text.

    .EXAMPLE
        # Using Microsoft SQL Server credentials
            $cred = Get-Credential
            Get-SQLSessionInfo -ServerInstance "127.0.0.1,1433"  -Credential $cred -Verbose -IncludeMySPID
    .EXAMPLE
        # Use Get-SQLSessionInfo to select from Out-GridView, then run again for the selected SPID.
            $target = Get-SQLSessionInfo -ServerInstance "localhost" -Credential $cred -TrustServerCertificate -Verbose | Select-Object SPID, Status, LoginName, HostName, LoginTime, Command, CPUTime, Reads, Writes, ElapsedSeconds | Out-ConsoleGridView
            Get-SQLSessionInfo -ServerInstance "localhost"  -Credential $cred  -spid $target.SPID -Verbose -TrustServerCertificate


    .Example
        # Every 5 seconds for one minute, get sessions and display them on the screen.
            $ctr=0
            do{
                Clear-Host
                ;Get-SQLSessionInfo -ServerInstance "localhost" -Credential $cred -TrustServerCertificate | Select-Object SPID, ElapsedSeconds, CPUTime, Reads, Writes | Format-Table
                ;Start-Sleep -Seconds 5
                ;$ctr+=1
            }until($ctr -eq 12)
    .EXAMPLE
        # ServerInstance can be passed in the pipeline but credentials must be the same for each lookup
            $cred = Get-Credential
            $ServerInstances = "localhost","localhost"
            $ServerInstances | Get-SQLSessionInfo -Credential $cred -Verbose -TrustServerCertificate -OnlyStatementText | Select-Object SPID, ServerInstance


    .INPUTS
        ServerInstance
        TrustServerCertificate
        Credential
        IncludeMySPID
        SPID
        SearchLoginName
        Status
        OnlyStatementText

    .OUTPUTS
        System.Array

    .NOTES
        Author:  David Allen
        Email: dallenva@gmail.com
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory,ValueFromPipeline)][String]$ServerInstance
        ,[switch]$TrustServerCertificate
        ,[PSCredential]$Credential
        ,[switch]$IncludeMySPID
        ,[int]$SPID
        ,[string]$SearchLoginName
        ,[string]$Status
        ,[switch]$OnlyStatementText
        )
    Process {
    # Setup parameters for invoke-sql
        $SQLParams = @{
            ServerInstance= $ServerInstance
        }
    # Add Trust Server Certificate if requested
        if($TrustServerCertificate -eq $true){$SQLParams += @{TrustServerCertificate = $true}}
    # Add Credentials if requested
        if($null -ne $Credential){$SQLParams.Credential = $Credential}

    # Get current SPID and test server connection, Exit if SQL call fails.
        Write-Verbose "Get current user SPID"
        $SQLGetMySPID = "Select @@SPID as [MYSPID]"
        try{$MYSPID = invoke-sql @SQLParams -query $SQLGetMySPID}catch{write-error ("Query Failed, check server instance, certificate, and credentials. " + $_.Exception.Message);throw}
        Write-Verbose ("Running user's SPID: "+$MYSPID.MYSPID)

    # Script to look up sessions
        $sql = "
SELECT
    s.session_id as [SPID],
    r.status as [Status],
    r.blocking_session_id as [BlockedBySID],
    r.wait_type as [WaitType],
    wait_resource as [WaitResource],
    r.wait_time / (1000.0) as [WaitSeconds],
    r.cpu_time as [CPUTime],
    r.logical_reads as [LogicalReads],
    r.reads as [Reads],
    r.writes as [Writes],
    r.total_elapsed_time / (1000.0) as [ElapsedSeconds],
    r.total_elapsed_time / (1000.0) / (60) as [ElapsedMinutes],
    r.total_elapsed_time / (1000.0) / (60) / (60) as [ElapsedHours],
    Substring(
        st.TEXT,
        (r.statement_start_offset / 2) + 1,
        (
            (
                CASE r.statement_end_offset
                    WHEN -1 THEN Datalength (st.TEXT)
                    ELSE r.statement_end_offset
                END - r.statement_start_offset
            ) / 2
        ) + 1
    ) AS [StatementText],
    Coalesce(
        QuoteName (Db_name (st.dbid)) + N'.' + QuoteName (Object_schema_name (st.objectid, st.dbid)) + N'.' + QuoteName (Object_name (st.objectid, st.dbid)),
        ''
    ) AS [CommandText],
    r.command as [Command],
    s.login_name as [LoginName],
    s.host_name as [HostName],
    s.program_name as [ProgramName],
    s.last_request_end_time as [LastRequestEndTime],
    s.login_time as [LoginTime],
    r.open_transaction_count as [OpenTransactionCount],
    --r.sql_handle,
    @@servername as [ServerInstance]
FROM
    sys.dm_exec_sessions AS s
    Left JOIN sys.dm_exec_requests AS r on r.session_id = s.session_id
    Outer APPLY sys.Dm_exec_sql_text (r.sql_handle) AS st
WHERE
    (@ExcludeSPID = 0 OR s.session_id <> @ExcludeSPID)
    AND (@SPID = 0 OR s.session_id = @SPID)
    AND (@SearchLoginName IS NULL OR s.login_name = @SearchLoginName)
    AND (@Status IS NULL OR r.status = @Status)
ORDER BY
    CASE WHEN r.session_id IS NULL THEN 1 ELSE 0 END,
    r.cpu_time desc,
    r.status,
    r.blocking_session_id,
    s.session_id"
        $SessionQueryParameters = @{
            ExcludeSPID = if($IncludeMySPID) {0} else {[int]$MYSPID.MYSPID}
            SPID = $SPID
            SearchLoginName = if($SearchLoginName) {$SearchLoginName} else {$null}
            Status = if($Status) {$Status} else {$null}
        }

    # Query SPIDs
        Write-Verbose "Get All SPIDs"
        try{$RunningSessions = @(invoke-sql @SQLParams -query $sql -QueryParameters $SessionQueryParameters)}catch{write-error ("Query Failed, check server instance, certificate, and credentials. " + $_.Exception.Message);throw }
        if ($null -eq $RunningSessions -or $RunningSessions.Count -eq 0){Write-Verbose "No SPIDs found"}else{Write-Verbose ($RunningSessions.Count.ToString() + " SPIDs found")}

    ###### Filter records based on parameters
    # Filters are applied by SQL
        if($includeMySPID -eq $true){Write-Verbose ("Including SPID "+$myspid.MYSPID+" for this session.")}else{Write-Verbose ("Excluding SPID: "+$MYSPID.MYSPID)}
        if($SPID -gt 0){Write-Verbose ("Including SPID: "+$SPID)}
        if($SearchLoginName){Write-Verbose ("Including Login: "+$SearchLoginName)}
        if($Status){Write-Verbose ("Only showing for Status: "+$Status)}
    # If OnlyCommandText then only return where CommandText has a value
        if($OnlyStatementText.IsPresent){$RunningSessions = $RunningSessions | Where-Object {-not [string]::IsNullOrWhiteSpace($_.StatementText)};Write-Verbose ("Only showing where StatementText has a value.")}
    ######
        #Output any remaining records
        return @($RunningSessions | Sort-Object -Property SPID)
    }
}



Function Get-SQLJobInfo {
    <#
    .SYNOPSIS
        Returns a list of SQL Agent Jobs on a Microsoft SQL Server

    .DESCRIPTION
        Get-SQLJobInfo is a function that returns a list of Microsoft SQL Server SQL Agent Jobs.  Additional options
        will let you filter the results.

        The output contains one row per job schedule and job step. Jobs with multiple schedules
        therefore appear in multiple rows for each configured step.

        ServerInstance can also be passed in the pipeline.

        Selecting -Verbose displays safe connection and query details.

    .PARAMETER ServerInstance
        The Microsoft SQL Server host or instance to query.

    .PARAMETER TrustServerCertificate
        Enables accepting a server certificate that is not trusted locally. Leave
        this switch off when certificate validation is required.

    .PARAMETER Credential
        Specifies a user account that has permission to perform this action. The default
        is the current user with Integrated Security.

    .PARAMETER IgnoreDisabled
        If specified, disabled jobs are excluded from the results.

    .PARAMETER SearchName
        If provided, job and step names are filtered with this SQL `LIKE` pattern.

    .EXAMPLE
        # Using Microsoft SQL Server credentials
            $cred = Get-Credential
            Get-SQLJobInfo -ServerInstance "127.0.0.1,1433"  -Credential $cred -Verbose

    .INPUTS
        ServerInstance
        TrustServerCertificate
        Credential
        IgnoreDisabled
        SearchName

    .OUTPUTS
        System.Array

    .NOTES
        Author:  David Allen
        Email: dallenva@gmail.com
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory,ValueFromPipeline)][String]$ServerInstance
        ,[switch]$TrustServerCertificate
        ,[PSCredential]$Credential
        ,[switch]$IgnoreDisabled
        ,[string]$SearchName
        )
    Process {
        # Setup parameters for invoke-sql
        $SQLParams = @{
            ServerInstance= $ServerInstance
        }
    # Add Trust Server Certificate if requested
        if($TrustServerCertificate -eq $true){$SQLParams += @{TrustServerCertificate = $true}}
    # Add Credentials if requested
        if($null -ne $Credential){$SQLParams.Credential = $Credential}

    # Script to look up sessions
        $sql = "
select
    @@ServerName as [ServerInstance]
    ,j.job_id as [JobID]
    ,j.name as [JobName]
    ,iif(j.enabled=1,'Yes','No') as [JobEnabled]
	,c.name as [JobCategory]
	,j.description as [Description]
	,j.date_created AS [JobCreatedOn]
    ,j.date_modified AS [JobLastModifiedOn]
	,iif(sSCH.schedule_uid is not NULL,'Yes','No') AS [IsScheduled]
    ,sSCH.schedule_uid AS [JobScheduleID]
    ,sSCH.name AS [JobScheduleName]

	, case sSCH.freq_type when 1 then 'One time only'
                 when 4 then 'Daily'
                 when 8 then 'Weekly'
                 when 16 then 'Monthly'
                 when 32 then 'Monthly, relative to freq_interval'
                 when 64 then 'Runs when SQL Server Agent service starts'
                 when 128 then 'Runs when computer is idle'
                 else 'Unknown'
                 end as JobFrequency
, case when sSCH.freq_type = 1 then 'Unused'
       when sSCH.freq_type = 4 then 'Every ' + convert(varchar(10), sSCH.freq_interval) + ' days'
       when sSCH.freq_type = 8 then 'Weekly: every'
                   + case when sSCH.freq_interval & 1 = 1 then ' Sunday' else '' end
                   + case when sSCH.freq_interval & 2 = 2 then ' Monday' else '' end
                   + case when sSCH.freq_interval & 4 = 4 then ' Tuesday' else '' end
                   + case when sSCH.freq_interval & 8 = 8 then ' Wednesday' else '' end
                   + case when sSCH.freq_interval & 16 = 16 then ' Thursday' else '' end
                   + case when sSCH.freq_interval & 32 = 32 then ' Friday' else '' end
                   + case when sSCH.freq_interval & 64 = 64 then ' Saturday' else '' end

       when sSCH.freq_type = 16 then 'Monthly: on the ' + convert(varchar(10), sSCH.freq_interval) + ' day of every ' + convert(varchar(10), sSCH.freq_recurrence_factor) + ' month(s)'
       when sSCH.freq_type = 32 then 'Monthly: on the ' + case when sSCH.freq_relative_interval = 0 then 'Unused'
                 when sSCH.freq_relative_interval = 1 then 'First'
                 when sSCH.freq_relative_interval = 2 then 'Second'
                 when sSCH.freq_relative_interval = 4 then 'Third'
                 when sSCH.freq_relative_interval = 8 then 'Fourth'
                 when sSCH.freq_relative_interval = 16 then 'Last'
                 else 'Unknown' end + ' ' + case when sSCH.freq_interval = 1 then  'Sunday'
                                                                when sSCH.freq_interval = 2 then  'Monday'
                                                                when sSCH.freq_interval = 3 then  'Tuesday'
                                                                when sSCH.freq_interval = 4 then  'Wednesday'
                                                                when sSCH.freq_interval = 5 then  'Thursday'
                                                                when sSCH.freq_interval = 6 then  'Friday'
                                                                when sSCH.freq_interval = 7 then  'Saturday'
                                                                when sSCH.freq_interval = 8 then  'Day'
                                                                when sSCH.freq_interval = 9 then  'Weekday'
                                                                when sSCH.freq_interval = 10 then 'Weekend day'
                                                                end + ' of every ' + convert(varchar(10), sSCH.freq_recurrence_factor) + ' month(s)'
       else 'Unused'
       end as JobFrequencyIntervalDescription
, case when sSCH.freq_type = 1 then 'At the specified time'
       when sSCH.freq_subday_type = 1 then 'At the specified time'
       when sSCH.freq_subday_type = 2 then 'Seconds'
       when sSCH.freq_subday_type = 4 then 'Minutes'
       when sSCH.freq_subday_type = 8 then 'Hours'
                 end as JobFrequencyTypeDescription
, case
when sSCH.freq_type = 1 then 'At ' + substring(convert(varchar(23), msdb.dbo.agent_datetime(sSCH.active_start_date, sSCH.active_start_time), 121),12, 12)
when sSCH.freq_subday_type = 1 then 'At ' + substring(convert(varchar(23), msdb.dbo.agent_datetime(sSCH.active_start_date, sSCH.active_start_time), 121),12, 12)
when sSCH.freq_subday_type in (2,4,8) then 'Every ' + convert(varchar(10), sSCH.freq_subday_interval) + ' ' + case sSCH.freq_subday_type when 1 then 'At the specified time'
                 when 2 then 'Seconds' + ' between ' + substring(convert(varchar(23), msdb.dbo.agent_datetime(sSCH.active_start_date, sSCH.active_start_time), 121),12, 12) + ' and '  + substring(convert(varchar(23), msdb.dbo.agent_datetime(sSCH.active_end_date, sSCH.active_end_time), 121),12, 12)
                 when 4 then 'Minutes' + ' between ' + substring(convert(varchar(23), msdb.dbo.agent_datetime(sSCH.active_start_date, sSCH.active_start_time), 121),12, 12) + ' and '  + substring(convert(varchar(23), msdb.dbo.agent_datetime(sSCH.active_end_date, sSCH.active_end_time), 121),12, 12)
                 when 8 then 'Hours' + ' between ' + substring(convert(varchar(23), msdb.dbo.agent_datetime(sSCH.active_start_date, sSCH.active_start_time), 121),12, 12) + ' and '  + substring(convert(varchar(23), msdb.dbo.agent_datetime(sSCH.active_end_date, sSCH.active_end_time), 121),12, 12)
                 end
                 else 'Unused' end as JobFrequencySubTypeDescription



    , CASE j.delete_level
        WHEN 0 THEN 'Never'
        WHEN 1 THEN 'On Success'
        WHEN 2 THEN 'On Failure'
        WHEN 3 THEN 'On Completion'
    END AS [JobDeletionCriterion]
    ,sx.step_id as [StepNumber]
    ,sx.step_name as [StepName]
	,sx.step_uid as [StepID]
    ,sx.subsystem as [StepType]
	,sx.database_name as [StepDataBase]
	,sPROX.name AS [RunAs]
    ,case
        when charindex('/FILE `"\`"',sx.command,1) > 0 then SUBSTRING(sx.command,charindex('/FILE `"',sx.command,1)+9,charindex('.dtsx',sx.command,charindex('/FILE `"',sx.command,1))-5)
        when charindex('/FILE',sx.command,1) > 0 and charindex('.dtsx`"',sx.command,1) > 0 then SUBSTRING(sx.command,charindex('/FILE `"',sx.command,1)+7,charindex('.dtsx',sx.command,charindex('/FILE `"',sx.command,1))-3)
        when charindex('/ISSERVER `"\`"',sx.command,1) > 0 then SUBSTRING(sx.command,charindex('/ISSERVER `"\`"',sx.command,1)+13,charindex('.dtsx',sx.command,charindex('/ISSERVER `"\`"',sx.command,1))-9)
        else null
    end as [StepDTSXFileName]
    ,case
        when charindex('/CONFIGFILE `"\`"',sx.command,1) > 0 then SUBSTRING(sx.command,charindex('/CONFIGFILE `"\`"',sx.command,1)+15,charindex('.dtsConfig',sx.command,1)-charindex('/CONFIGFILE `"\`"',sx.command,1)-5)
        when charindex('/CONFIGFILE',sx.command,1) > 0 then SUBSTRING(sx.command,charindex('/CONFIGFILE `"\`"',sx.command,1)+13,charindex('.dtsConfig',sx.command,charindex('/CONFIGFILE `"',sx.command,1)+charindex('/CONFIGFILE `"',sx.command,1)))
        else null
    end as [StepConfigFileName]
    ,case
        when sx.last_run_date = 0 then null
        else convert(datetime,msdb.dbo.agent_datetime(sx.last_run_date, sx.Last_run_time))
    end as [StepLastRunDateTime]
	,sx.command as [StepCommand]
		,CASE sx.on_success_action
		WHEN 1 THEN 'Quit the job reporting success'
		WHEN 2 THEN 'Quit the job reporting failure'
		WHEN 3 THEN 'Go to the next step'
		WHEN 4 THEN 'Go to Step: '
					+ QUOTENAME(CAST(sx.on_success_step_id AS VARCHAR(3)))
					+ ' '
					+ sOSSTP.step_name
		END AS [OnSuccessAction]
	,sx.retry_attempts AS [RetryAttempts]
	,sx.retry_interval AS [RetryInterval (Minutes)]
	,CASE sx.on_fail_action
		WHEN 1 THEN 'Quit the job reporting success'
		WHEN 2 THEN 'Quit the job reporting failure'
		WHEN 3 THEN 'Go to the next step'
		WHEN 4 THEN 'Go to Step: '
					+ QUOTENAME(CAST(sx.on_fail_step_id AS VARCHAR(3)))
					+ ' '
					+ sOFSTP.step_name
		END AS [OnFailureAction]
from
	msdb.dbo.sysjobsteps sx
	inner join msdb.dbo.sysjobs j on sx.job_id = j.job_id
	inner join msdb.dbo.syscategories AS C ON j.category_id = c.category_id
	left outer join msdb.dbo.sysjobschedules AS jsch ON j.job_id = jsch.job_id
    left outer join msdb.dbo.sysschedules AS sSCH ON jsch.schedule_id = sSCH.schedule_id
	left outer join msdb.dbo.sysproxies AS sPROX ON sx.proxy_id = sPROX.proxy_id
	left outer join msdb.dbo.sysjobsteps AS sOSSTP ON sx.job_id = sOSSTP.job_id AND sx.on_success_step_id = sOSSTP.step_id
	left outer join msdb.dbo.sysjobsteps AS sOFSTP ON sx.job_id = sOFSTP.job_id AND sx.on_fail_step_id = sOFSTP.step_id
where
    TRY_CONVERT(uniqueidentifier,j.Name) is null --Remove SSRS jobs
    AND (@IgnoreDisabled = 0 OR j.enabled = 1)
    AND (@SearchName = '%' OR j.name LIKE @SearchName OR sx.step_name LIKE @SearchName)
order by
    j.name,j.job_id,sx.step_id"
        $SearchPattern = if($SearchName) {$SearchName} else {'%'}
        $QueryParameters = @{
            SearchName = $SearchPattern
            IgnoreDisabled = [int]$IgnoreDisabled.IsPresent
        }

    # Query Jobs
        Write-Verbose "Attempting to retrieve jobs."
        try{$SQLAgentJobs = @(invoke-sql @SQLParams -query $sql -QueryParameters $QueryParameters)}catch{write-error ("Query Failed, check server instance, certificate, and credentials. " + $_.Exception.Message);throw}
    # If nothing returned allow for verbose message
        if ($null -eq $SQLAgentJobs -or $SQLAgentJobs.Count -eq 0){Write-Verbose "No SQL Agent jobs found"}else{Write-Verbose ($SQLAgentJobs.Count.ToString() + " job steps found")}
    # Filters are applied by SQL
        if($IgnoreDisabled -eq $true){Write-Verbose ("Ignoring disabled jobs.")}
        if($SearchName.Length -gt 0){Write-Verbose ("Including Name Search: "+$SearchName)}
    #Output any remaining records
        if ($SQLAgentJobs.count -gt 0){$Final = "Record Count after filters: "+$SQLAgentJobs.count}else{$Final = "Record Count after filters: 0"}
        Write-Verbose  ($final)
        return @($SQLAgentJobs | Sort-Object -Property JobName)
    }
}

Function Get-SQLRunningJobs {
    <#
    .SYNOPSIS
        Returns a list of currently running SQL Agent Jobs on a Microsoft SQL Server

    .DESCRIPTION
        Get-SQLRunningJobs is a function that returns a list of running Microsoft SQL Server SQL Agent Jobs.  Additional options
        will let you filter the results.

        ServerInstance can also be passed in the pipeline.

        Selecting -Verbose displays safe connection and query details.

    .PARAMETER ServerInstance
        The Microsoft SQL Server host or instance to query.

    .PARAMETER TrustServerCertificate
        Enables accepting a server certificate that is not trusted locally. Leave
        this switch off when certificate validation is required.

    .PARAMETER Credential
        Specifies a user account that has permission to perform this action. The default
        is the current user with Integrated Security.

    .PARAMETER SearchName
        If provided, job and step names are filtered with this SQL `LIKE` pattern.

    .EXAMPLE
        # Using Microsoft SQL Server credentials
            $cred = Get-Credential
            Get-SQLRunningJobs -ServerInstance "127.0.0.1,1433"  -Credential $cred -Verbose

    .INPUTS
        ServerInstance
        TrustServerCertificate
        Credential
        SearchName

    .OUTPUTS
        System.Array

    .NOTES
        Author:  David Allen
        Email: dallenva@gmail.com
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory,ValueFromPipeline)][String]$ServerInstance
        ,[switch]$TrustServerCertificate
        ,[PSCredential]$Credential
        ,[string]$SearchName
        )
    Process {
        # Setup parameters for invoke-sqlcmd
        $SQLParams = @{
            ServerInstance= $ServerInstance
        }
    # Add Trust Server Certificate if requested
        if($TrustServerCertificate -eq $true){$SQLParams += @{TrustServerCertificate = $true}}
    # Add Credentials if requested
        if($null -ne $Credential){$SQLParams.Credential = $Credential}

    # Script to look up agent jobs
        $sql = "
DECLARE @Now datetime = GETDATE();
SELECT @@servername as [ServerInstance],
    j.name AS [JobName],
    j.job_id as [JobId],
    ISNULL(last_executed_step_id,0)+1 AS [CurrentStep],
    Js.step_name as [StepName],
    CONVERT(varchar(10), DATEDIFF_BIG(second, ja.start_execution_date, @Now) / 3600)
    + ':' + RIGHT('0' + CONVERT(varchar(2), (DATEDIFF_BIG(second, ja.start_execution_date, @Now) % 3600) / 60), 2)
    + ':' + RIGHT('0' + CONVERT(varchar(2), DATEDIFF_BIG(second, ja.start_execution_date, @Now) % 60), 2)
    as [JobElapsedTime],
    ja.start_execution_date as [StartedOn]
FROM msdb.dbo.sysjobactivity ja
    JOIN msdb.dbo.sysjobs j on ja.job_id = j.job_id
    LEFT JOIN msdb.dbo.sysjobsteps js on ja.job_id = js.job_id AND ISNULL(ja.last_executed_step_id,0)+1 = js.step_id
WHERE
    ja.session_id = (SELECT TOP 1 session_id FROM msdb.dbo.syssessions ORDER BY agent_start_date DESC)
AND start_execution_date is not null
AND stop_execution_date is null
AND (@SearchName = '%' OR j.name LIKE @SearchName OR js.step_name LIKE @SearchName)"
        $SearchPattern = if($SearchName) {$SearchName} else {'%'}

    # Query jobs
        Write-Verbose "Attempting to retrieve jobs."
        try{$SQLAgentJobs = @(invoke-sql @SQLParams -query $sql -QueryParameters @{ SearchName = $SearchPattern })}catch{write-error ("Query Failed, check server instance, certificate, and credentials. " + $_.Exception.Message);throw}
    # If nothing returned allow for verbose message
        if ($null -eq $SQLAgentJobs -or $SQLAgentJobs.Count -eq 0){Write-Verbose "No SQL Agent jobs found"}else{Write-Verbose ($SQLAgentJobs.Count.ToString() + " jobs found")}
    # Search filters are applied by SQL
        if($SearchName.Length -gt 0){Write-Verbose ("Including Name Search: "+$SearchName)}
    #Output any remaining records
        return @($SQLAgentJobs)
    }
}
Function Get-SQLJobHistory {
    <#
    .SYNOPSIS
        Returns a list of SQL Agent Job History on a Microsoft SQL Server

    .DESCRIPTION
        Get-SQLJobHistory is a function that returns a list of Microsoft SQL Server SQL Agent Job History.  Additional options
        will let you filter the results.

        ServerInstance can also be passed in the pipeline.

        Selecting -Verbose displays safe connection and query details.

    .PARAMETER ServerInstance
        The Microsoft SQL Server host or instance to query.

    .PARAMETER TrustServerCertificate
        Enables accepting a server certificate that is not trusted locally. Leave
        this switch off when certificate validation is required.

    .PARAMETER Credential
        Specifies a user account that has permission to perform this action. The default
        is the current user with Integrated Security.

    .PARAMETER IgnoreDisabled
        If specified, disabled jobs are excluded from the results.

    .PARAMETER SearchName
        If provided, job names are filtered with this SQL `LIKE` pattern. If omitted, all job history is returned.

    .EXAMPLE
        # Using Microsoft SQL Server credentials
            $cred = Get-Credential
            Get-SQLJobHistory -ServerInstance "127.0.0.1,1433"  -Credential $cred -Verbose

    .INPUTS
        ServerInstance
        TrustServerCertificate
        Credential
        IgnoreDisabled
        SearchName

    .OUTPUTS
        System.Array

    .NOTES
        Author:  David Allen
        Email: dallenva@gmail.com
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory,ValueFromPipeline)][String]$ServerInstance
        ,[switch]$TrustServerCertificate
        ,[PSCredential]$Credential
        ,[switch]$IgnoreDisabled
        ,[string]$SearchName
        )
    Process {
        # Setup parameters for invoke-sqlcmd
        $SQLParams = @{
            ServerInstance= $ServerInstance
        }
    # Add Trust Server Certificate if requested
        if($TrustServerCertificate -eq $true){$SQLParams += @{TrustServerCertificate = $true}}
    # Add Credentials if requested
        if($null -ne $Credential){$SQLParams.Credential = $Credential}

    # Script to look up sessions
        $sql = "
SELECT
@@ServerName as [ServerInstance]
,J.name as [JobName]
,iif(j.enabled=1, 'Yes','No') as [JobEnabled]
,H.message as [Message]
,run_status = CASE H.run_status
                WHEN 0 THEN 'Failed'
                WHEN 1 THEN 'Succeeded'
                WHEN 2 THEN 'Retry'
                WHEN 3 THEN 'Canceled'
                WHEN 4 THEN 'In progress'
            END
, msdb.dbo.agent_datetime(h.run_date, h.run_time) as [StartDateTime]
,CONVERT(varchar(10), H.run_duration / 10000)
    + ':' + RIGHT('0' + CONVERT(varchar(2), (H.run_duration % 10000) / 100), 2)
    + ':' + RIGHT('0' + CONVERT(varchar(2), H.run_duration % 100), 2) as [Duration_H-M-S]
,convert(int, substring(right('00000000'+ltrim(rtrim(convert(varchar(10), H.run_duration))), 8), 7,2)) + convert(int, substring(right('00000000'+ltrim(rtrim(convert(varchar(10), H.run_duration))), 8), 5,2)) * 60 + convert(int, substring(right( '00000000'+ltrim(rtrim(convert(varchar(10),H.run_duration))), 8),1,4)) * 3600 as [DurationSeconds]
FROM
msdb.dbo.sysjobs J
INNER JOIN msdb.dbo.sysjobhistory H ON J.job_id = H.job_id and h.step_id = 0
where
j.name like @SearchName
ORDER BY
J.Name asc,h.run_date desc, h.run_time desc
    "
        $SearchPattern = if($SearchName) {$SearchName} else {'%'}

    # Query Jobs
        Write-Verbose "Attempting to retreive jobs."
        Write-Verbose $sql
        try{$SQLAgentJobs = @(invoke-sql @SQLParams -query $sql -QueryParameters @{ SearchName = $SearchPattern })}catch{write-error ("Query Failed, check server instance, certificate, and credentials. " + $_.Exception.Message);throw}
    # If nothing returned allow for verbose message
        if ($null -eq $SQLAgentJobs -or $SQLAgentJobs.Count -eq 0){Write-Verbose "No SQL Agent jobs found"}else{Write-Verbose ($SQLAgentJobs.Count.ToString() + " job steps found")}
    # If IgnoreDiabled remove disabled jobs
        if($IgnoreDisabled.IsPresent){$SQLAgentJobs = $SQLAgentJobs | Where-Object {$_.JobEnabled -eq "Yes"};Write-Verbose ("Ignoring disabled jobs.")}
    # Output any remaining records
        if ($SQLAgentJobs.count -gt 0){$Final = "Record Count after filters: "+$SQLAgentJobs.count}else{$Final = "Record Count after filters: 0"}
        Write-Verbose  ($final)
        return @($SQLAgentJobs | Sort-Object -Property JobName)
    }
}
function Format-ColorBar {
<#
    .SYNOPSIS
        Create a string to display a progress bar

    .DESCRIPTION
        Creates a text based percent bar as a string.

    .PARAMETER Width
        Width of the output string in characters, from 1 through 1,000.
        Defaults to 10.

    .PARAMETER Percent
        Percentage of the bar to fill, from 0 through 100.

    .PARAMETER ForegroundColor
        Color of the bar, defaults to screen font color.

    .EXAMPLE
        Format-ColorBar -Width 100 -Percent 50 -ForegroundColor Green

    .INPUTS
        Width
        Percent
        ForegroundColor

    .OUTPUTS
        String

    .NOTES
        Author:  David Allen
        Email: dallenva@gmail.com
    #>
    Param(
        [ValidateRange(1, 1000)]
        [int32]$Width = 10,
        [ValidateRange(0, 100)]
        [decimal]$Percent,
        [ValidateSet('Black','Blue','Cyan','DarkBlue','DarkCyan','DarkGray','DarkGreen','DarkMagenta','DarkRed','DarkYellow','Gray','Green','Magenta','Red','White','Yellow')]
        [string]$ForegroundColor
    )
    Process {
        $Display = [Math]::Round($width * ($percent / 100))
        $Output = ('').padright($Display,"█")
        $Output += ('').padright($Width - $Display,"░")
        If ($ForegroundColor.Length -gt 0){$Output = $PSStyle.Foreground.$ForegroundColor + $Output + $PSStyle.Reset}
        $Output
    }
}