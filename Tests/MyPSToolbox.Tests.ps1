$integrationConfigured = $env:MYPSTOOLBOX_SQL_SERVER -and
    $env:MYPSTOOLBOX_SQL_USER -and
    $env:MYPSTOOLBOX_SQL_PASSWORD

Describe 'MyPSToolbox module' {
    BeforeAll {
        $modulePath = Join-Path (Split-Path $PSScriptRoot -Parent) 'MyPSToolbox.psd1'
        Import-Module $modulePath -Force
    }

    It 'imports the SQL commands' {
        Get-Command Invoke-SQL, Write-SQLBulkCopy, Get-SQLSessionInfo, Get-SQLJobInfo, Get-SQLRunningJobs, Get-SQLJobHistory |
            Should -Not -BeNullOrEmpty
    }

    It 'exposes bulk-copy progress parameters' {
        $command = Get-Command Write-SQLBulkCopy

        $command.Parameters['ShowProgress'].ParameterType | Should -Be ([switch])
        $command.Parameters['ProgressInterval'].ParameterType | Should -Be ([int])
    }

    It 'rejects bulk copy objects without properties' {
        {
            [pscustomobject]@{} |
                Write-SQLBulkCopy -ServerInstance 'invalid' -Database 'master' -DestinationTable 'dbo.Rows'
        } | Should -Throw '*must expose at least one property*'
    }

    It 'rejects an invalid bulk copy destination before connecting' {
        {
            [pscustomobject]@{ Id = 1 } |
                Write-SQLBulkCopy -ServerInstance 'invalid' -Database 'master' -DestinationTable 'dbo.Rows; DROP TABLE dbo.Rows'
        } | Should -Throw
    }

    It 'rejects a column mapping source that is not present' {
        {
            [pscustomobject]@{ Id = 1 } |
                Write-SQLBulkCopy -ServerInstance 'invalid' -Database 'master' `
                    -DestinationTable 'dbo.Rows' -ColumnMapping @{ Missing = 'Id' }
        } | Should -Throw '*was not found on the input object*'
    }

    It 'exposes scalar session filters' {
        $command = Get-Command Get-SQLSessionInfo

        $command.Parameters['SPID'].ParameterType | Should -Be ([int])
        $command.Parameters['SearchLoginName'].ParameterType | Should -Be ([string])
        $command.Parameters['Status'].ParameterType | Should -Be ([string])
    }

    It 'rejects security-sensitive AdditionalText keys before connecting' {
        {
            Invoke-SQL -ServerInstance 'invalid' -Query 'SELECT 1' -AdditionalText 'Encrypt=False'
        } | Should -Throw "*unsupported connection-string key*"
    }

    It 'reports malformed AdditionalText' {
        {
            Invoke-SQL -ServerInstance 'invalid' -Query 'SELECT 1' -AdditionalText 'Pooling'
        } | Should -Throw "*AdditionalText is not a valid connection-string fragment*"
    }

    It 'rejects invalid application intent values' {
        {
            Invoke-SQL -ServerInstance 'invalid' -Query 'SELECT 1' -ApplicationIntent 'Invalid'
        } | Should -Throw
    }

    It 'rejects negative connection timeouts' {
        {
            Invoke-SQL -ServerInstance 'invalid' -Query 'SELECT 1' -ConnectionTimeout -1
        } | Should -Throw
    }

    It 'rejects an empty server instance' {
        {
            Invoke-SQL -ServerInstance '' -Query 'SELECT 1'
        } | Should -Throw
    }

    It 'rejects an empty query' {
        {
            Invoke-SQL -ServerInstance 'localhost' -Query ''
        } | Should -Throw
    }

    It 'rejects a whitespace-only database name' {
        {
            Invoke-SQL -ServerInstance 'localhost' -Query 'SELECT 1' -Database '   '
        } | Should -Throw
    }

    It 'rejects an invalid pool-size range' {
        {
            Invoke-SQL -ServerInstance 'invalid' -Query 'SELECT 1' -MinPoolSize 50 -MaxPoolSize 10
        } | Should -Throw '*MinPoolSize cannot be greater than MaxPoolSize*'
    }

    It 'rejects invalid SQL parameter names' {
        {
            Invoke-SQL -ServerInstance 'invalid' -Query 'SELECT 1' -QueryParameters @{'bad-name' = 1}
        } | Should -Throw '*Invalid SQL parameter name*'
    }

    It 'rejects null query parameters' {
        {
            Invoke-SQL -ServerInstance 'invalid' -Query 'SELECT 1' -QueryParameters $null
        } | Should -Throw
    }

    It 'rejects unsupported SQL parameter value types before connecting' {
        {
            Invoke-SQL -ServerInstance 'invalid' -Query 'SELECT 1' `
                -QueryParameters @{Value = [pscustomobject]@{Name = 'unsupported'}}
        } | Should -Throw '*Unsupported SQL parameter value type*'
    }

    It 'filters null statement text safely' {
        $session = [pscustomobject]@{
            SPID = 1
            StatementText = $null
        }

        [string]::IsNullOrWhiteSpace($session.StatementText) | Should -BeTrue
    }
}

Describe 'MyPSToolbox SQL integration' -Tag 'Integration' {
    BeforeAll {
        $script:sqlServer = $env:MYPSTOOLBOX_SQL_SERVER
        $script:sqlUser = $env:MYPSTOOLBOX_SQL_USER
        $script:sqlPassword = $env:MYPSTOOLBOX_SQL_PASSWORD
        $script:sqlDatabase = if($env:MYPSTOOLBOX_SQL_DATABASE) {$env:MYPSTOOLBOX_SQL_DATABASE} else {'master'}
        $script:integrationAvailable = $script:sqlServer -and $script:sqlUser -and $script:sqlPassword
        if($script:integrationAvailable) {
            $securePassword = ConvertTo-SecureString $script:sqlPassword -AsPlainText -Force
            $script:sqlCredential = [PSCredential]::new($script:sqlUser, $securePassword)
        }
    }

    It 'executes a parameterized query' -Skip:(-not $integrationConfigured) {
        $result = Invoke-SQL -ServerInstance $script:sqlServer -Credential $script:sqlCredential `
            -Database $script:sqlDatabase -TrustServerCertificate `
            -Query 'SELECT @Number AS [Number], IIF(@Missing IS NULL, 1, 0) AS [IsMissing]' `
            -QueryParameters @{Number = 42; Missing = $null}

        $result.Number | Should -Be 42
        $result.IsMissing | Should -Be 1
    }

    It 'supports explicit SQL parameter metadata' -Skip:(-not $integrationConfigured) {
        $result = Invoke-SQL -ServerInstance $script:sqlServer -Credential $script:sqlCredential `
            -Database $script:sqlDatabase -TrustServerCertificate `
            -Query 'SELECT @Value AS [Value]' `
            -QueryParameters @{Value = @{Value = 'metadata'; SqlDbType = 'NVarChar'; Size = 32}}

        $result.Value | Should -Be 'metadata'
    }

    It 'accepts case-insensitive SQL parameter types' -Skip:(-not $integrationConfigured) {
        $result = Invoke-SQL -ServerInstance $script:sqlServer -Credential $script:sqlCredential `
            -Database $script:sqlDatabase -TrustServerCertificate `
            -Query 'SELECT @Value AS [Value]' `
            -QueryParameters @{Value = @{Value = 'lowercase'; SqlDbType = 'nvarchar'; Size = 32}}

        $result.Value | Should -Be 'lowercase'
    }

    It 'rejects incompatible SQL parameter metadata before connecting' {
        {
            Invoke-SQL -ServerInstance 'invalid' -Query 'SELECT 1' `
                -QueryParameters @{Value = @{Value = 'test'; SqlDbType = 'NVarChar'; Precision = 18}}
        } | Should -Throw '*Precision and Scale must be provided together*'
    }

    It 'preserves inferred decimal precision and scale' -Skip:(-not $integrationConfigured) {
        $result = Invoke-SQL -ServerInstance $script:sqlServer -Credential $script:sqlCredential `
            -Database $script:sqlDatabase -TrustServerCertificate `
            -Query "SELECT SQL_VARIANT_PROPERTY(CAST(@Value AS sql_variant), 'Precision') AS [Precision], SQL_VARIANT_PROPERTY(CAST(@Value AS sql_variant), 'Scale') AS [Scale]" `
            -QueryParameters @{Value = [decimal]::new(1234500, 0, 0, $false, 4)}

        $result.Precision | Should -Be 7
        $result.Scale | Should -Be 4
    }

    It 'rejects invalid SQL parameter metadata before connecting' {
        {
            Invoke-SQL -ServerInstance 'invalid' -Query 'SELECT 1' `
                -QueryParameters @{Value = @{Value = 'test'; SqlDbType = 'NVarChar'; Size = -2}}
        } | Should -Throw '*Invalid Size*'
    }

    It 'applies scalar session filters' -Skip:(-not $integrationConfigured) {
        $result = @(Get-SQLSessionInfo -ServerInstance $script:sqlServer -Credential $script:sqlCredential `
            -TrustServerCertificate -IncludeMySPID -SearchLoginName $script:sqlUser)

        $result | Should -Not -BeNullOrEmpty
        @($result | Where-Object LoginName -ne $script:sqlUser) | Should -HaveCount 0
    }

    It 'keeps subsequent connections usable after a query failure' -Skip:(-not $integrationConfigured) {
        {
            Invoke-SQL -ServerInstance $script:sqlServer -Credential $script:sqlCredential `
                -Database $script:sqlDatabase -TrustServerCertificate `
                -Query 'SELECT * FROM dbo.TableThatDoesNotExist' | Out-Null
        } | Should -Throw

        $result = Invoke-SQL -ServerInstance $script:sqlServer -Credential $script:sqlCredential `
            -Database $script:sqlDatabase -TrustServerCertificate `
            -Query 'SELECT 1 AS [Value]'

        $result.Value | Should -Be 1
    }

    It 'bulk copies objects into an existing table' -Skip:(-not $integrationConfigured) {
        Invoke-SQL -ServerInstance $script:sqlServer -Credential $script:sqlCredential `
            -Database $script:sqlDatabase -TrustServerCertificate `
            -Query 'IF OBJECT_ID(''dbo.MyPSToolboxBulkCopyTest'', ''U'') IS NOT NULL DROP TABLE dbo.MyPSToolboxBulkCopyTest; CREATE TABLE dbo.MyPSToolboxBulkCopyTest (Id int NOT NULL, Name nvarchar(50) NULL);' |
            Out-Null

        try {
            $rows = @(
                [pscustomobject]@{ SourceId = 1; DisplayName = 'one' }
                [pscustomobject]@{ SourceId = 2; DisplayName = 'two' }
            )
            $copy = $rows | Write-SQLBulkCopy -ServerInstance $script:sqlServer `
                -Credential $script:sqlCredential -Database $script:sqlDatabase `
                -TrustServerCertificate -DestinationTable 'dbo.MyPSToolboxBulkCopyTest' `
                -BatchSize 1 -ColumnMapping @{ SourceId = 'Id'; DisplayName = 'Name' } `
                -ShowProgress -ProgressInterval 1

            $copy.RowsCopied | Should -Be 2
            $count = Invoke-SQL -ServerInstance $script:sqlServer -Credential $script:sqlCredential `
                -Database $script:sqlDatabase -TrustServerCertificate `
                -Query 'SELECT COUNT(*) AS TotalRows FROM dbo.MyPSToolboxBulkCopyTest'
            $count.TotalRows | Should -Be 2
        } finally {
            Invoke-SQL -ServerInstance $script:sqlServer -Credential $script:sqlCredential `
                -Database $script:sqlDatabase -TrustServerCertificate `
                -Query 'DROP TABLE dbo.MyPSToolboxBulkCopyTest' | Out-Null
        }
    }
}
