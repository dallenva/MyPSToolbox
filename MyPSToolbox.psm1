

function Invoke-SQL {
    <#
        .SYNOPSIS
            Invoke-SQL is an alternative to Invoke-Sqlcmd that does not require the SQLServer module to be installed.  It uses the .Net SQLClient

        .DESCRIPTION
            Invoke-SQL is an alternative to Invoke-Sqlcmd that does not require the SQLServer module.  It uses the .Net SQLClient to query SQL Server.

            Selecting -Verbose will display the connection string being used.  -ShowPassword will display the password in the connection string.

        .PARAMETER ServerInstance
            The Micorosft SQL Server host to query for sessions.

        .PARAMETER TrustServerCertificate
            Newer SQL Server clients will not connect if the server certificate is not
            stored locally.  Older SQL Clients will not work with this parameter so using "-TrustServerCertificate $flase" will
            remove this parameter.

        .PARAMETER [PSCredential]Credential
            Specifies a user account that has permission to perform this action. The default
            is the current user with Intergrated Security.

        .PARAMETER Database
            The database to query.  The default is the current database.

        .PARAMETER Encryption
            Encrypt the connection to SQL Server.

        .PARAMETER ConnectionTimeout
            The number of seconds before a connection is considered timed out.  The default is 10 seconds.

        .PARAMETER Query
            The query test to submitt to SQL Server.

        .PARAMETER ShowPassword
            Will display the password in the connection string if -Verbose is used.

        .PARAMETER AdditionalText
            Any additional text to add to the connection string.

        .EXAMPLE
            # Using Microsoft SQL Server credentials
                $cred = Get-Credential
                Invoke-SQL -ServerInstance "localhost" -Credential $cred -Query "SELECT @@servername as 'Server'" -Verbose

        .INPUTS
            ServerInstance
            TrustServerCertificate
            Credential
            Database
            ConnectionTimeout
            Credentials
            ShowPassword
            AdditionalText

        .OUTPUTS
            Dataset

        .NOTES
            Author:  David Allen
            Email: dallenva@gmail.com
        #>
        [CmdletBinding()]
        Param([parameter(Mandatory)][string]$ServerInstance
                ,[parameter(Mandatory)][string]$Query
                ,[string]$Database
                ,[int]$ConnectionTimeout=10
                ,[pscredential]$Credential = $null
                ,[switch]$Encryption
                ,[switch]$TrustServerCertificate
                ,[switch]$ShowPassword
                ,[string]$AdditionalText)
        Process {
    # Build Connection
            $conn = new-object System.Data.SqlClient.SQLConnection
    # Build Connection String
            $ConnectionString = "Server='{0}'" -f $ServerInstance
            If($Database.Length -gt 0) {$ConnectionString += ";Database='{0}'" -f $Database}
            If($ConnectionTimeout -gt 0) {$ConnectionString += ";Connect Timeout={0}" -f $ConnectionTimeout}
            If($Credential.length -gt 0) {$ConnectionString += ";Integrated Security=false;User ID='{0}';Password='{1}'" -f $Credential.UserName, $Credential.GetNetworkCredential().Password}else{$ConnectionString += ";Integrated Security=true"}
            If($Encryption) {$ConnectionString += ";Encrypt='Yes'"}
            If($TrustServerCertificate-eq $true) {$ConnectionString += ";TrustServerCertificate='Yes'"}
            if($AdditionalText) {$ConnectionString += ";{0}" -f $AdditionalText}
            $InfoConStr = $ConnectionString
            if($showpassword -eq $false -and $credential.length -gt 0){$InfoConStr = $InfoConStr.Replace(($Credential.GetNetworkCredential().Password),"********")}
            $Info = ("SQL ConnectionString before trying to connect: " + $InfoConStr)
            Write-Verbose $Info
    # Set Connection String
            try{$conn.ConnectionString = $ConnectionString}catch{throw $_}
    # Open Connection
            $conn.Open()

    # If sucessful open connection, run query
            if($conn.state -eq 'Open'){
                    Write-Verbose "----------------------------------------------------------"
                    Write-Verbose "Open() successful.... Connection Details:"
                    Write-Verbose ("Query Started: " + (Get-Date))
                    write-verbose ("ConnectionString: " + $conn.ConnectionString)
                    write-verbose ("Database: " + $conn.Database)
                    write-verbose ("DataSource: " + $conn.DataSource)
                    write-verbose ("ServerVersion: " + $conn.ServerVersion)
                    write-verbose ("WorkstationId: " + $conn.WorkstationId)
                    $time = measure-command{$ExecSQL = New-Object system.Data.SqlClient.SqlCommand($Query, $conn)
                                            $ExecSQL.CommandTimeout = 720
    # Convert result to table
                                            $Results = New-Object System.Data.SqlClient.SqlDataAdapter($ExecSQL)
                                            $Data = New-Object System.Data.DataSet
                                            [void]$Results.Fill($Data)
    # Close SQL Connection
                                            $conn.Close()}
                    $Info = "Elapsed Seconds: " + $time.TotalSeconds
                    Write-Verbose ("Query Ended: " + (Get-Date))
                    Write-Verbose $Info
                    $Info = "Record Count: " + $Data.Tables[0].Rows.Count
                    Write-Verbose $Info
                    Write-Verbose "-- End Invoke-SQL ----------------------------------------"
    # Return Results
                    Return $Data.Tables[0]}
        }
    }

Function Get-EmptyDirectory {
    <#
    .SYNOPSIS
        Get empty directories using underlying Get-ChildItem cmdlet

    .NOTES
        Name: Get-EmptyDirectory
        Author: theSysadminChannel
        Version: 1.0
        DateCreated: 2021-Oct-2
        Modified David Allen 2024
q
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
                if ($PSBoundParameters.ContainsKey('Recurse')) {
                    $ItemParams.Add('Recurse',$true)
                }

                if ($PSBoundParameters.ContainsKey('Depth')) {
                    $ItemParams.Add('Depth',$Depth)
                }
                $FolderList = Get-ChildItem @ItemParams

                $folderlist | ForEach-Object{
                    if ((Get-ChildItem -literalpath $_.FullName).count -lt 1) {
                        [PSCustomObject]@{
                            EmtpyDirectory = $true
                            Path           = $_.FullName
                        }
                    } else {}
                }
            } catch {
                Write-Error $_.Exception.Message
            }
        }
        END {}
    }
function Set-WindowTitle {
<#
.SYNOPSIS
    Change the PowerShell window title

.NOTES
    Change the PowerShell window title

.EXAMPLE
    set-windowtitle -title "My Session"
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

.NOTES
    Author:  David Allen
    Email: dallenva@gmail.com

.EXAMPLE
    Set-Cursor -X 0 -Y 0
#>
Param([int]$x, [int] $y)
Process {
    $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates $x , $y
}
}
function Set-ConsoleWindowSize
{
<#
.SYNOPSIS
    Set Width and Height of the console window.  This is dependent on the terminal, VS Code and Windows Terminal
    will not allow this to override the setting.

.NOTES
    Author:  David Allen
    Email: dallenva@gmail.com

.EXAMPLE
    Set-ConsoleWindowSize -Width 80 -Height 24
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
        $Maxvalue = ($_.Exception.Message |Select-String "\d+").Matches[0].Value
        $WindowSize.Height = $Maxvalue
        $Host.UI.RawUI.WindowSize = $WindowSize
    }
}
function Lock-Screen {
    Param()
    Process {
            if($ismacos){pmset sleepnow}
            if($islinux){gnome-screensaver-command -l}
            if($iswindows){$xCmdString = {rundll32.exe user32.dll,LockWorkStation};Invoke-Command $xCmdString}
        }
    }
function Get-UI {
    Process {
        $host.ui.rawui #Because I'm tired of googleing to find this variable
    }
}
function get-whoami {
    Write-Host "====================================="
    write-host "PowerShell Edition: " -NoNewline -ForegroundColor Green
    write-host $PSVersionTable.PSEdition -ForegroundColor Cyan
    write-host "Powershell Version: " -NoNewline -ForegroundColor Green
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
    write-host $home -ForegroundColor Cyan
    Write-Host "Profile: " -NoNewline -ForegroundColor Green
    write-host $Profile -ForegroundColor Cyan
    Write-Host "====================================="
}
function Get-MyDrives {
    $return = Get-PSDrive -PSProvider FileSystem
    $return = $return | where-object  {$_.name -notlike "Temp"}
    return $return
}
Function Out-ScreenGrid {
    <#
    .SYNOPSIS
        This is an alternative to Format-Table, when an object passed the text is converted into a table
        of up to 6 columns and colorized with alternating formatting.

    .DESCRIPTION
        This was created for a few specific use cases and may not be the best solution for your case.
    .PARAMETER InputObject
        Any System Object. Enclose in () if needed
    .PARAMETER TableHeader
        String to appear before the table
    .PARAMETER ShowFooter
        Show table footer
    .EXAMPLE
        Out-ScreenGrid -InputObject (Get-ChildItem | select-object name, extension, CreationTime,LastAccessTime)
    .EXAMPLE
        Out-ScreenGrid -DataObject (Get-process)
    .INPUTS
        System.Object
        [string]TableHeader
        [Boolean]ShowFooter
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
        $MyWidth = $host.UI.RawUI.Windowsize.Width
        $Reset = $PSStyle.Reset
        $OddLinesFormat = $psstyle.Background.cyan + $PSStyle.Foreground.Black
        $EvenLinesFormat = $PSStyle.Background.brightCyan + $psstyle.Foreground.Black
        $HeaderLineFormat = $psstyle.Bold+$psstyle.Underline+$PSStyle.Background.blue + $psstyle.Foreground.white
        $MaxColumns = 5
<#
Create options for
        max lines
        Header format colors
        Odd Even format colors
        Paging
#>
    # Convert Object into String CSV to format
        $StringData = ($InputObject) | Convertto-CSV -delimiter "|" -QuoteFields ""
    # Convert Header row into an array
        $Header = $StringData[0].split("|")
    # Calc number of columns up to 5
        $ColumnNumber = if($Header.count -le $MaxColumns){$Header.count}else{$MaxColumns}
        $Offset = $ColumnNumber + 5
    # Calc Column width for formating
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
    # The process leave line breaks, remove them and add a single one at the end
            ;$OutputString += [string]::join("",($TempRowString.Split("`n"))) + "`r`n"
            ;$RowIdx += 1
        }until($RowIdx -eq $InputObject.count+1)
    # Send output string to the console
        clear-host
        write-host $OutputString
        If($TableFooter.length -gt 0){Write-host $HeaderLineFormat$TableFooter$Reset}else{
            if($nofooter -eq $false){write-host $psstyle.Formatting.Warning"<END>"$Reset}}
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

        The Module SQLServer must be installed, the function will test each time it's run and provide the command to
        install if needed.

        ServerInstance can also be passed in the pipeline.

        Selecting -Verbose will display details like the SPID used by the current session.

    .PARAMETER ServerInstance
        The Micorosft SQL Server host to query for sessions.

    .PARAMETER TrustServerCertificate
        Newer SQL Server clients will not connect if the server certificate is not
        stored locally.  Older SQL Clients will not work with this parameter so using "-TrustServerCertificate $flase" will
        remove this parameter.

    .PARAMETER Credential
        Specifies a user account that has permission to perform this action. The default
        is the current user with Intergrated Security.

    .PARAMETER IncludeMySPID
        If -IncludeMySPID is used the Comandlet will filter out the current SPID's session from the returned sessions.

    .PARAMETER SPID
        Using -SPID will let you send an array of SPIDs for the Commandlet to filter to only the sessions you are intrested.

    .PARAMETER SearchLoginName
        -SearchLoginName allows you to filter returned sessions by login name.

    .PARAMETER Status
        Using -Status you can filter on one or more status.

    .PARAMETER OnlyStatementText
        Selecting -OnlyStatementText will filter to records with an entry in the sys.Dm_exec_sql_text table.

    .EXAMPLE
        # Using Microsoft SQL Server credentials
            $cred = Get-Credential
            Get-SQLSessionInfo -ServerInstance "127.0.0.1,1433"  -Credential $cred -Verbose -IncludeMySPID
    .EXAMPLE
        #Use Get-SQLSessionInfo to select from an out-gridview then run again for only the selected SPID(s)
            $target = Get-SQLSessionInfo -ServerInstance "localhost"  -Credential $cred -TrustServerCertificate -Verbose| Select-Object SPID,Status,LoginName,HostName,LoginTime,Command,CPUTime,Reads,Writes,ElapsedSeconds| out-consolegridview
            Get-SQLSessionInfo -ServerInstance "localhost"  -Credential $cred  -spid $target.SPID -Verbose -TrustServerCertificate


    .Example
        # Every 5 second for one minute get sessions and display on the screen
            $ctr=0
            do{
                Clear-Host
                ;Get-SQLSessionInfo -ServerInstance "localhost"  -Credential $cred -TrustServerCertificate | select-object SPID,ElapsedSeconds,CPUTime,Reads,Writes | Format-Table
                ;Start-Sleep -Seconds 5
                ;$ctr+=1
            }until($ctr -eq 12)
    .EXAMPLE
        # ServerInstance can be passed in the pipeline but credentials must be the same for each lookup
            $cred = Get-Credential
            $ServerInstances = "localhost","localhost"
            $ServerInstances | Get-SQLSessionInfo  -Credential $cred -Verbose  -TrustServerCertificate -OnlyStatementText | select-object spid,ServerInstance


    .INPUTS
        ServerInstance
        TrustServerCertificate
        Credential
        IncludeMySPID
        OnlyMySPID
        SPID
        SearchLoginName
        Status

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
        ,[pscredential]$Credential
        ,[switch]$IncludeMySPID
        ,[array]$SPID
        ,[array]$SearchLoginName
        ,[array]$Status
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
        if($Credential.count -gt 0){$SQLParams += @{Credential = $Credential}}

    # Get current SPID and test server connection, Exit if SQL call fails.
        Write-Verbose "Get current user SPID"
        $SQLGetMySPID = "Select convert(varchar(10),@@SPID) as 'MYSPID'"
        try{$MYSPID = invoke-sql @SQLParams -query $SQLGetMySPID -ErrorAction silentlycontinue}catch{write-error "Query Failed, check server instance, certificate, and credentials.";throw}
        Write-Verbose ("Running user's SPID: "+$MYSPID.MYSPID)

    # Script to look up sessions
        $sql = "
USE msdb
SELECT
    s.session_id as 'SPID',
    r.status as 'Status',
    r.blocking_session_id as 'BlockedBySID',
    r.wait_type as 'WaitType',
    wait_resource as 'WaitResource',
    r.wait_time / (1000.0) as 'WaitSeconds',
    r.cpu_time as 'CPUTime',
    r.logical_reads as 'LogicalReads',
    r.reads as 'Reads',
    r.writes as 'Writes',
    r.total_elapsed_time / (1000.0) as 'ElapsedSeconds',
    r.total_elapsed_time / (1000.0) / (60) as 'ElapsedMinutes',
    r.total_elapsed_time / (1000.0) / (60) / (60) as 'ElapsedHours',
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
    ) AS 'StatementText',
    Coalesce(
        Quotename (Db_name (st.dbid)) + N'.' + Quotename (Object_schema_name (st.objectid, st.dbid)) + N'.' + Quotename (Object_name (st.objectid, st.dbid)),
        ''
    ) AS 'CommandText',
    r.command as 'Command',
    s.login_name as 'LoginName',
    s.host_name as 'HostName',
    s.program_name as 'ProgramName',
    s.last_request_end_time as 'LastRequestEndTime',
    s.login_time as 'LoginTime',
    r.open_transaction_count as 'OpenTransactionCount',
    --r.sql_handle,
    @@servername as 'ServerInstance'
FROM
    sys.dm_exec_sessions AS s
    Inner JOIN sys.dm_exec_requests AS r on r.session_id = s.session_id
    Outer APPLY sys.Dm_exec_sql_text (r.sql_handle) AS st
ORDER BY
    r.cpu_time desc,
    r.status,
    r.blocking_session_id,
    s.session_id"

    # Query SPIDs
        Write-Verbose "Get All SPIDs"
        try{$RunningSessions = invoke-sql @SQLParams -query $sql -ErrorAction silentlycontinue}catch{write-error "Query Failed, check server instance, certificate, and credentials.";throw }
        if ($RunningSessions.count -lt 0){Write-Verbose "No SPID(s) found"}else{Write-Verbose ($RunningSessions.count.ToString() + " SPID(s) Found")}

    ###### Filter records based on parameters
    # If IncludeMySPID is false then remove this session's SPID
        if($includeMySPID -eq $true){Write-Verbose ("Including SPID "+$myspid.MYSPID+" for this session.")}else{$RunningSessions = $RunningSessions | where-object {$_.SPID -ne $MYSPID.MYSPID};Write-Verbose ("Excluding SPID: "+$MYSPID.MYSPID)}
    # If only looking at specific SPIDs remove all but those SPIDs
        if($SPID -gt 0){$RunningSessions = $RunningSessions | where-object {$_.SPID -in $SPID};Write-Verbose ("Including SPID(s): "+$SPID)}
    # If only looking at specific login names remove all but those login names
        if($SearchLoginName -gt 0){$RunningSessions = $RunningSessions | where-object {$_.LoginName -in $SearchLoginName};Write-Verbose ("Including Login(s): "+$SearchLoginName)}
    # If SPIDStatus only return selected status
        if($Status.count -gt 0){$RunningSessions = $RunningSessions | where-object {$_.Status -in $Status};Write-Verbose ("Only showing for Status: "+$Status)}
    # If OnlyCommandText then only return where CommandText has a value
        if($OnlyStatementText -eq $true){$RunningSessions = $RunningSessions | where-object {$_.StatementText.length -gt 1};Write-Verbose ("Only showing where StatementText has a value.")}
    ######
        #Ouput any remaining records
        return $RunningSessions | Sort-Object -Property SPID
    }
}



Function Get-SQLJobInfo {
    <#
    .SYNOPSIS
        Returns a list of SQL Agent Jobs on a Microsoft SQL Server

    .DESCRIPTION
        Get-SQLJobInfo is a function that returns a list of Microsoft SQL Server SQL Agent Jobs.  Additional options
        will let you filter the results.

        The Module SQLServer must be installed, the function will test each time it's run and provide the command to
        install if needed.

        ServerInstance can also be passed in the pipeline.

        Selecting -Verbose will display details like the SPID used by the current session.

    .PARAMETER ServerInstance
        The Micorosft SQL Server host to query for sessions.

    .PARAMETER TrustServerCertificate
        Newer SQL Server clients will not connect if the server certificate is not
        stored locally.  Older SQL Clients will not work with this parameter so using "-TrustServerCertificate $flase" will
        remove this parameter.

    .PARAMETER Credential
        Specifies a user account that has permission to perform this action. The default
        is the current user with Intergrated Security.

    .PARAMETER IgnoreDisabled
        If IgnoreDisabled is true then disabled jobs will be excluded from the results

    .PARAMETER SearchName
        If a string is provided Job Names and Step Names will be searched for the string.

    .EXAMPLE
        # Using Microsoft SQL Server credentials
            $cred = Get-Credential
            Get-SQLJobInfo -ServerInstance "127.0.0.1,1433"  -Credential $cred -Verbose

    .INPUTS
        ServerInstance
        TrustServerCertificate
        Credential
        IgnoreDisabled
        SearchJobName

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
        ,[pscredential]$Credential
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
        if($Credential.count -gt 0){$SQLParams += @{Credential = $Credential}}

    # Script to look up sessions
        $sql = "
USE msdb
select
    @@ServerName as 'ServerInstance'
    ,j.job_id as 'JobID'
    ,j.name as 'JobName'
    ,iif(j.enabled=1,'Yes','No') as 'JobEnabled'
	,c.name as 'JobCategory'
	,j.description as 'Description'
	,j.date_created AS 'JobCreatedOn'
    ,j.date_modified AS 'JobLastModifiedOn'
	,iif(sSCH.schedule_uid is not NULL,'Yes','No') AS 'IsScheduled'
    ,sSCH.schedule_uid AS 'JobScheduleID'
    ,sSCH.name AS 'JobScheduleName'
    , CASE j.delete_level
        WHEN 0 THEN 'Never'
        WHEN 1 THEN 'On Success'
        WHEN 2 THEN 'On Failure'
        WHEN 3 THEN 'On Completion'
    END AS 'JobDeletionCriterion'
    ,sx.step_id as 'StepNumber'
    ,sx.step_name as 'StepName'
	,sx.step_uid as 'StepID'
    ,sx.subsystem as 'StepType'
	,sx.database_name as 'StepDataBase'
	,sPROX.name AS 'RunAs'
    ,case
        when charindex('/FILE `"\`"',sx.command,1) > 0 then SUBSTRING(sx.command,charindex('/FILE `"',sx.command,1)+9,charindex('.dtsx',sx.command,charindex('/FILE `"',sx.command,1))-5)
        when charindex('/FILE',sx.command,1) > 0 then SUBSTRING(sx.command,charindex('/FILE `"',sx.command,1)+7,charindex('.dtsx',sx.command,charindex('/FILE `"',sx.command,1))-3)
        when charindex('/ISSERVER `"\`"',sx.command,1) > 0 then SUBSTRING(sx.command,charindex('/ISSERVER `"\`"',sx.command,1)+13,charindex('.dtsx',sx.command,charindex('/ISSERVER `"\`"',sx.command,1))-9)
        else null
    end as 'StepDTSXFileName'
    ,case
        when charindex('/CONFIGFILE `"\`"',sx.command,1) > 0 then SUBSTRING(sx.command,charindex('/CONFIGFILE `"\`"',sx.command,1)+15,charindex('.dtsConfig',sx.command,1)-charindex('/CONFIGFILE `"\`"',sx.command,1)-5)
        --when charindex('/CONFIGFILE',sx.command,1) > 0 then SUBSTRING(sx.command,charindex('/CONFIGFILE `"\`"',sx.command,1)+13,charindex('.dtsConfig',sx.command,charindex('/CONFIGFILE `"',sx.command,1)+charindex('/CONFIGFILE `"',sx.command,1)))
        else null
    end as 'StepConfigFileName'
    ,case
        when sx.last_run_date = 0 then null
        else convert(datetime,msdb.dbo.agent_datetime(sx.last_run_date, sx.Last_run_time))
    end as 'StepLastRunDateTime'
	,sx.command as 'StepCommand'
		,CASE sx.on_success_action
		WHEN 1 THEN 'Quit the job reporting success'
		WHEN 2 THEN 'Quit the job reporting failure'
		WHEN 3 THEN 'Go to the next step'
		WHEN 4 THEN 'Go to Step: '
					+ QUOTENAME(CAST(sx.on_success_step_id AS VARCHAR(3)))
					+ ' '
					+ sOSSTP.step_name
		END AS 'OnSuccessAction'
	,sx.retry_attempts AS 'RetryAttempts'
	,sx.retry_interval AS 'RetryInterval (Minutes)'
	,CASE sx.on_fail_action
		WHEN 1 THEN 'Quit the job reporting success'
		WHEN 2 THEN 'Quit the job reporting failure'
		WHEN 3 THEN 'Go to the next step'
		WHEN 4 THEN 'Go to Step: '
					+ QUOTENAME(CAST(sx.on_fail_step_id AS VARCHAR(3)))
					+ ' '
					+ sOFSTP.step_name
		END AS 'OnFailureAction'
from
	msdb.dbo.sysjobsteps sx
	inner join msdb.dbo.sysjobs j on sx.job_id = j.job_id
	inner join msdb.dbo.syscategories AS C ON j.category_id = c.category_id
	left outer join msdb.dbo.sysjobschedules AS jsch ON j.job_id = jsch.job_id
    left outer join msdb.dbo.sysschedules AS sSCH ON jsch.schedule_id = sSCH.schedule_id
	left outer join msdb.dbo.sysproxies AS sPROX ON sx.proxy_id = sPROX.proxy_id
	left outer join msdb.dbo.sysjobsteps AS sOSSTP ON sx.job_id = sOSSTP.job_id AND sx.on_success_step_id = sOSSTP.step_id
	left outer join msdb.dbo.sysjobsteps AS sOFSTP ON sx.job_id = sOFSTP.job_id AND sx.on_success_step_id = sOFSTP.step_id
where
    TRY_CONVERT(uniqueidentifier,j.Name) is null --Remove SSRS jobs
order by
    j.name,j.job_id,sx.step_id"

    # Query Jobs
        Write-Verbose "Attempting to retreive jobs."
        try{$SQLAgentJobs = invoke-sql @SQLParams -query $sql -ErrorAction silentlycontinue}catch{write-error "Query Failed, check server instance, certificate, and credentials.";throw}
    # If nothing returned allow for verbose message
        if ($SQLAgentJobs.count -lt 0){Write-Verbose "No SQL Agent Job(s) found"}else{Write-Verbose ($SQLAgentJobs.count.ToString() + " Jobs(s)/Step(s) Found")}
    # If IgnoreDiabled remove disabled jobs
        if($IgnoreDisabled -eq $true){$SQLAgentJobs = $SQLAgentJobs | where-object {$_.JobEnabled -eq "Yes"};Write-Verbose ("Ignoreing disabled jobs.")}
    # If only looking at specific login names remove all but those login names
        if($SearchName.Length -gt 0){$SQLAgentJobs = $SQLAgentJobs | where-object {$_.JobName -like $Searchname -or $_.stepName -like $Searchname} ;Write-Verbose ("Including Name Search: "+$SearchName)}
    #Ouput any remaining records
        if ($SQLAgentJobs.count -gt 0){$Final = "Record Count after filters: "+$SQLAgentJobs.count}else{$Final = "Record Count after filters: 0"}
        Write-Verbose  ($final)
        return $SQLAgentJobs | Sort-Object -Property JobName
    }
}

Function Get-SQLRunningJobs {
    <#
    .SYNOPSIS
        Returns a list of currently running SQL Agent Jobs on a Microsoft SQL Server

    .DESCRIPTION
        Get-SQLRunningJobs is a function that returns a list of running Microsoft SQL Server SQL Agent Jobs.  Additional options
        will let you filter the results.

        The Module SQLServer must be installed, the function will test each time it's run and provide the command to
        install if needed.

        ServerInstance can also be passed in the pipeline.

        Selecting -Verbose will display details like the SPID used by the current session.

    .PARAMETER ServerInstance
        The Micorosft SQL Server host to query for sessions.

    .PARAMETER TrustServerCertificate
        Newer SQL Server clients will not connect if the server certificate is not
        stored locally.  Older SQL Clients will not work with this parameter so using "-TrustServerCertificate $flase" will
        remove this parameter.

    .PARAMETER Credential
        Specifies a user account that has permission to perform this action. The default
        is the current user with Intergrated Security.

    .PARAMETER SearchName
        If a string is provided Job Names and Step Names will be searched for the string.

    .EXAMPLE
        # Using Microsoft SQL Server credentials
            $cred = Get-Credential
            Get-SQLRunningJobs -ServerInstance "127.0.0.1,1433"  -Credential $cred -Verbose

    .INPUTS
        ServerInstance
        TrustServerCertificate
        Credential
        SearchJobName

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
        ,[pscredential]$Credential
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
        if($Credential.count -gt 0){$SQLParams += @{Credential = $Credential}}

    # Script to look up agent jobs
        $sql = "
USE msdb
SELECT @@servername as 'ServerInstance',
    j.name AS 'JobName',
    j.job_id as 'JobId',
    ISNULL(last_executed_step_id,0)+1 AS 'CurrentStep',
    Js.step_name as 'StepName',
    (cast(
        (cast(cast(getdate() as float) - cast(ja.start_execution_date as float) as int) * 24) /* hours over 24 */
        + datepart(hh, getdate() - ja.start_execution_date) /* hours */
        as varchar(10))
    + ':' + right('0' + cast(datepart(mi, getdate() - ja.start_execution_date) as varchar(2)), 2) /* minutes */
    + ':' + right('0' + cast(datepart(ss, getdate() - ja.start_execution_date) as varchar(2)), 2) /* seconds */
    ) as 'JobElapsedTime',
    ja.start_execution_date as 'StartedOn'
FROM msdb.dbo.sysjobactivity ja
    LEFT JOIN msdb.dbo.sysjobhistory jh on ja.job_history_id = jh.instance_id
    JOIN msdb.dbo.sysjobs j on ja.job_id = j.job_id
    JOIN msdb.dbo.sysjobsteps js on ja.job_id = js.job_id AND ISNULL(ja.last_executed_step_id,0)+1 = js.step_id
WHERE
    ja.session_id = (SELECT TOP 1 session_id FROM msdb.dbo.syssessions ORDER BY agent_start_date DESC)
AND start_execution_date is not null
AND stop_execution_date is null"

    # Query jobss
        Write-Verbose "Attempting to retreive jobs."
        try{$SQLAgentJobs = invoke-sql @SQLParams -query $sql -ErrorAction silentlycontinue}catch{write-error "Query Failed, check server instance, certificate, and credentials.";throw}
    # If nothing returned allow for verbose message
        if ($SQLAgentJobs.count -lt 0){Write-Verbose "No SQL Agent Job(s) found"}else{Write-Verbose ($SQLAgentJobs.count.ToString() + " Jobs(s) Found")}
    # If only looking at specific login names remove all but those login names
        if($SearchName.Length -gt 0){$SQLAgentJobs = $SQLAgentJobs | where-object {$_.JobName -like $Searchname -or $_.stepName -like $Searchname} ;Write-Verbose ("Including Name Search: "+$SearchName)}
    #Ouput any remaining records
        return $SQLAgentJobs
    }
}
Function Get-SQLJobHistory {
    <#
    .SYNOPSIS
        Returns a list of SQL Agent Job History on a Microsoft SQL Server

    .DESCRIPTION
        Get-SQLJobHistory is a function that returns a list of Microsoft SQL Server SQL Agent Job Hisotry.  Additional options
        will let you filter the results.

        The Module SQLServer must be installed, the function will test each time it's run and provide the command to
        install if needed.

        ServerInstance can also be passed in the pipeline.

        Selecting -Verbose will display details like the SPID used by the current session.

    .PARAMETER ServerInstance
        The Micorosft SQL Server host to query for sessions.

    .PARAMETER TrustServerCertificate
        Newer SQL Server clients will not connect if the server certificate is not
        stored locally.  Older SQL Clients will not work with this parameter so using "-TrustServerCertificate $flase" will
        remove this parameter.

    .PARAMETER Credential
        Specifies a user account that has permission to perform this action. The default
        is the current user with Intergrated Security.

    .PARAMETER IgnoreDisabled
        If IgnoreDisabled is true then disabled jobs will be excluded from the results

    .PARAMETER SearchName
        If a string is provided Job Names and Step Names will be searched for the string.

    .EXAMPLE
        # Using Microsoft SQL Server credentials
            $cred = Get-Credential
            Get-SQLJobHistory -ServerInstance "127.0.0.1,1433"  -Credential $cred -Verbose

    .INPUTS
        ServerInstance
        TrustServerCertificate
        Credential
        IgnoreDisabled
        SearchJobName

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
        ,[pscredential]$Credential
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
        if($Credential.count -gt 0){$SQLParams += @{Credential = $Credential}}

    # Script to look up sessions
        $sql = "
USE msdb
SELECT top 100000
    @@ServerName as 'ServerInstance'
    ,J.Job_id as 'JobId'
    ,J.name as 'JobName'
    ,iif(j.enabled=1,'Yes','No') as 'JobEnabled'
    ,S.step_id as 'StepNumber'
	,s.step_uid as 'StepID'
    ,S.step_name as 'StepName'
    ,H.message as 'Message'
    ,CASE H.run_status
        WHEN 0 THEN 'Failed'
        WHEN 1 THEN 'Succeeded'
        WHEN 2 THEN 'Retry'
        WHEN 3 THEN 'Canceled'
        WHEN 4 THEN 'In progress'
    END as RunStatus
    ,case
        when s.last_run_date = 0 then null
        else msdb.dbo.agent_datetime(h.run_date, h.run_time)
    end as 'StepLastRunDateTime'
    ,H.run_duration as 'DurationSeconds'
FROM
    sysjobhistory H
    INNER JOIN sysjobsteps S ON H.step_id = S.step_id AND H.job_id = S.job_id
    INNER JOIN sysjobs J ON J.job_id = H.job_id
where H.run_status != 4
ORDER BY
    msdb.dbo.agent_datetime(h.run_date, h.run_time)
    "

    # Query Jobs
        Write-Verbose "Attempting to retreive jobs."
        try{$SQLAgentJobs = invoke-sql @SQLParams -query $sql -ErrorAction silentlycontinue}catch{write-error "Query Failed, check server instance, certificate, and credentials.";throw}
    # If nothing returned allow for verbose message
        if ($SQLAgentJobs.count -lt 0){Write-Verbose "No SQL Agent Job(s) found"}else{Write-Verbose ($SQLAgentJobs.count.ToString() + " Jobs(s)/Step(s) Found")}
    # If IgnoreDiabled remove disabled jobs
        if($IgnoreDisabled -eq $true){$SQLAgentJobs = $SQLAgentJobs | where-object {$_.JobEnabled -eq "Yes"};Write-Verbose ("Ignoreing disabled jobs.")}
    # If only looking at specific login names remove all but those login names
        if($SearchName.Length -gt 0){$SQLAgentJobs = $SQLAgentJobs | where-object {$_.JobName -like $Searchname -or $_.stepName -like $Searchname} ;Write-Verbose ("Including Name Search: "+$SearchName)}
    #Ouput any remaining records
        if ($SQLAgentJobs.count -gt 0){$Final = "Record Count after filters: "+$SQLAgentJobs.count}else{$Final = "Record Count after filters: 0"}
        Write-Verbose  ($final)
        return $SQLAgentJobs | Sort-Object -Property JobName
    }
}