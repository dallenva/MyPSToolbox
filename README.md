# MyPSToolbox
A set of PowerShell cmdlets for retrieving Microsoft SQL Server information.

## Import the module

Import the module manifest from the repository root:

```powershell
Import-Module ./MyPSToolbox.psd1
Get-Command -Module MyPSToolbox
Get-Help Invoke-SQL -Detailed
```

The manifest requires PowerShell 7.0 or later.

`Write-SQLBulkCopy` uploads pipeline objects to an existing SQL Server table
using the .NET `SqlBulkCopy` API. Use `-BatchSize` to control transaction
batching and `-ColumnMapping` when source property names differ from target
column names. Add `-ShowProgress` to display copy progress; use
`-ProgressInterval` to control how often it updates.

## SQL Server compatibility

The embedded T-SQL is tested with SQL Server 2022. The current queries require SQL Server 2016 or later because they use features including `IIF`, `TRY_CONVERT`, and `DATEDIFF_BIG`.

## Tests

Pester 5 or later is required for the test suite. Install it for the current user with:

```powershell
Install-Module Pester -Scope CurrentUser -Force
```

Run the Pester test suite from the repository root:

```powershell
Invoke-Pester -Path ./Tests/MyPSToolbox.Tests.ps1
```

SQL integration tests run when all of these environment variables are set:

```powershell
$env:MYPSTOOLBOX_SQL_SERVER = 'localhost,1433'
$env:MYPSTOOLBOX_SQL_USER = 'sa'
$env:MYPSTOOLBOX_SQL_PASSWORD = '<password>'
$env:MYPSTOOLBOX_SQL_DATABASE = 'master'
Invoke-Pester -Path ./Tests/MyPSToolbox.Tests.ps1
```

Without these variables, the SQL integration tests are skipped.

Part of my job involves using SQL Server with some light dba functions.  Wanting to improve on the functionality offered by Microsoft and inspired by the SQLTop and DBATools I set out to develop my own tools.

I've also included some additional cmdlets I use in my scripts.

Each cmdlet supports get-help for more information

Invoke-SQL - A standalone version of InvokeSQLCMD that does not require the Microsoft SQLServer module.

Get-SQLSessionInfo - Returns a list of sessions on a Microsoft SQL Server similar to the SP_who and SP_who2 commands.

Get-SQLJobInfo - Returns a list of SQL Agent Jobs on a Microsoft SQL Server

Get-SQLRunningJobs - Returns a list of currently running SQL Agent Jobs on a Microsoft SQL Server

Get-SQLJobHistory - Returns a list of SQL Agent Job History on a Microsoft SQL Server

Other cmdlets

Out-ScreenGrid - This is an alternative to Format-Table, when an object passed the text is converted into a table of up to 6 columns and colorized with alternating formatting.  Future versions will allow for variable columns based on the size of the data and header.
