# MyPSToolbox.ps1
A set of PowerShell cmdlets for retrieving Microsoft SQL Server information.

Part of my job involves using SQL Server with some light dba functions.  Wanting to improve on the functionality offered by Microsoft and inspired by the SQLtop and dbatools I set out to develop my own tools.

I've also included some additional cmdlets I use in my scripts.

Each cmdlet supports get-help for more information

Get-SQLSessionInfo - Returns a list of sessions on a Microsoft SQL Server similar to the sp_who and sp_who2 commands.

Get-SQLJobInfo - Returns a list of SQL Agent Jobs on a Microsoft SQL Server

Get-SQLRunningJobs - Returns a list of currently running SQL Agent Jobs on a Microsoft SQL Server

Get-SQLJobHistory - Returns a list of SQL Agent Job History on a Microsoft SQL Server

Other cmdlets

Out-ScreenGrid - This is an alternative to Format-Table, when an object passed the text is converted into a table of up to 6 columns and colorized with alternating formatting.  Future versions will allow for variable columns based on the size of the data and header.

This does require Microsoft's SQLServer module, in the future I intend on repacing this by calling .net core directly.
