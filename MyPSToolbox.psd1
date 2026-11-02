@{
    RootModule        = 'MyPSToolbox.psm1'
    ModuleVersion     = '1.1.0'
    GUID              = '7f0d7b4e-6b3b-4a43-9dd4-72c2e0c3b9ab'
    Author            = 'David Allen'
    CompanyName       = ''
    Copyright         = '(c) David Allen. All rights reserved.'
    Description       = 'PowerShell tools for querying Microsoft SQL Server and supporting common console and file-system tasks.'
    PowerShellVersion = '7.0'

    FunctionsToExport = @(
        'Invoke-SQL'
        'Write-SQLBulkCopy'
        'Get-EmptyDirectory'
        'Set-WindowTitle'
        'Set-Cursor'
        'Set-ConsoleWindowSize'
        'Lock-Screen'
        'Get-UI'
        'get-whoami'
        'Get-MyDrives'
        'Out-ScreenGrid'
        'Get-SQLSessionInfo'
        'Get-SQLJobInfo'
        'Get-SQLRunningJobs'
        'Get-SQLJobHistory'
        'Format-ColorBar'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()

    PrivateData = @{
        PSData = @{
            Tags = @('PowerShell', 'SQL Server', 'T-SQL', 'Administration')
        }
    }
}
