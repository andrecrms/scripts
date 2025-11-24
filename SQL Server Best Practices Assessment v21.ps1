# Define log file path with timestamp
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFilePath = "C:\temp\SQL_Server_Best_Practices_Assessment_Execution_Log_$timestamp.txt"

# Start capturing console output
Start-Transcript -Path $logFilePath -Append

Write-Host @"
#============================================================================================================================================================================================
#                                                        # Welcome to SQL Server Best Practices Assessment Script v21!
#============================================================================================================================================================================================
# This script checks some SQL Server best practices, running it and understanding the results is for those who have been working with the product for some time.
# It can be run locally or remotely against server names inside a serverlist txt file, establishes remote sessions and executes several SQL queries to evaluate the following best practices:
# Instance settings, checkdb history, backup history, VLFs, autogrowth, trace flags, tempdb file checks, database options and compatibility levels.
# This script uses some queries from the BPCheck script to evaluate some best practices: https://github.com/microsoft/tigertoolbox/tree/master/BPCheck
# A quick summary of the findings will be presented at the end of the script, but all detailed results will be placed in a CSV file.
# Each status column will be either OK or REVIEW, when marked as REVIEW, navigate to the right of the sheet to understand why it was marked that way.
# A log of the execution will also be generated in the same directory as the csv file mentioned above.
# Tested on: SQL Server 2014 to 2022.
#
# Author: Andre Cesar Rodrigues
# LinkedIn: https://www.linkedin.com/in/andre-c-rodrigues
# Blog: http://sqlmagu.blogspot.com.br
# GitHub: https://github.com/andrecrms
# Last modified: 11/24/2025.
=============================================================================================================================================================================================
"@ -ForegroundColor Yellow
Write-Host @"
DISCLAIMER: This script should be tested in an appropriate environment before running in production. Additionally, properly validate your results as each environment may
have its own characteristics. This script will not change nothing in the environment, it will just run some SQL Queries to collect all necessary information.
"@ -ForegroundColor Red

# Prompt user for server input
$ServerName = Read-Host "Enter the server name (or press Enter to use the list from C:\temp\serverlist.txt)" # make sure to have just one server per row inside this file!

# Ensure '.' is treated as 'localhost'
if ($ServerName -eq ".") {
    $ServerName = "localhost"
}

# Determine server(s) to process
if ($ServerName) {
    Write-Host "Running for a single server: $ServerName"
    $serverEntries = @($ServerName)
}
else {
    Try {
        $serverListPath = "C:\temp\serverlist.txt"
        if (-Not (Test-Path $serverListPath)) {
            Throw "Server list file not found: $serverListPath"
        }
        $serverEntries = Get-Content $serverListPath
    }
    Catch {
        Write-Host "ERROR: Could not find the server list file at '$serverListPath'."
        Stop-Transcript
        Exit 1
    }
}


# Ask for FQDN usage with validation loop
while ($true) {
    $UseFQDN = (Read-Host "Do you want to use the full domain name (yes/no)?").ToLower()

    if ($UseFQDN -in @("yes", "no")) {
        break
    }

    Write-Host "Invalid input. Please type 'yes' or 'no'." -ForegroundColor Red
}

$DomainName = ""

if ($UseFQDN -eq "yes") {
    $DomainName = Read-Host "Enter the domain name (e.g., contoso.com)"
}

# Initialize arrays
$Jobs = @()
$Results = @()
$totalServers = $serverEntries.Count
$processedServers = 0

# Start processing servers
foreach ($entry in $serverEntries) {
    $processedServers++
    $progressPercent = [math]::Round(($processedServers / $totalServers) * 100)
    Write-Progress -Activity "Processing SQL Servers" -Status "Processing $processedServers of $totalServers servers ($progressPercent%)" -PercentComplete $progressPercent

    $BaseServerName = ($entry -split ',')[0]
    $VMs = if ($UseFQDN -eq "yes" -and $BaseServerName -ne "localhost") { "$BaseServerName.$DomainName" } else { $BaseServerName }
    $VMName = ($BaseServerName -split '\.')[0]

    Try {
        $Session = New-PSSession -ComputerName $VMs -ErrorAction Stop
        Write-Host "Successfully connected to $VMs"

        # Start job for SQL commands execution
        $Job = Invoke-Command -Session $Session -ScriptBlock {
            param($vmName)

            Try {
                Import-Module sqlps -DisableNameChecking -ErrorAction SilentlyContinue

                # Get SQL Server instance names
                Try {
                    $instanceNames = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server' -ErrorAction Stop).InstalledInstances
                }
                Catch {
                    Write-Host "SQL Server not installed on this server: $vmName"
                    return @()
                }

                $jobResults = @()

# Main config query
$mainQuery = @"
SELECT
    SERVERPROPERTY('ServerName') AS [Server Name],
    SERVERPROPERTY('ProductVersion') AS [SQL Build Number],
    SERVERPROPERTY('Edition') AS [SQL Edition],
    (SELECT total_physical_memory_kb / 1024 FROM sys.dm_os_sys_memory) AS [Total Server Memory (MB)],
    (SELECT cpu_count FROM sys.dm_os_sys_info) AS [Total Visible Processors],
    MAX(CASE WHEN name = 'min server memory (MB)' THEN value_in_use END) AS [Min Server Memory (MB)],
    MAX(CASE WHEN name = 'max server memory (MB)' THEN value_in_use END) AS [Max Server Memory (MB)],
    MAX(CASE WHEN name = 'optimize for ad hoc workloads' THEN value_in_use END) AS [Optimize for Ad Hoc Workloads],
    MAX(CASE WHEN name LIKE '%backup compression%' THEN value_in_use END) AS [Backup Compression Default],
    MAX(CASE WHEN name = 'remote admin connections' THEN value_in_use END) AS [Remote Admin Connections]
FROM sys.configurations
WHERE name IN (
    'min server memory (MB)',
    'max server memory (MB)',
    'optimize for ad hoc workloads',
    'remote admin connections'
) OR name LIKE '%backup compression%'
"@

# Compatibility level query
$compatQuery = @"
SELECT
    d.name AS [Database Name],
    d.compatibility_level AS [Compatibility Level],
    d.is_auto_update_stats_on AS [Auto Update Stats],
    d.is_auto_create_stats_on AS [Auto Create Stats],
    d.is_auto_shrink_on AS [Auto Shrink],
    d.page_verify_option_desc AS [Page Verify]
FROM sys.databases d
WHERE d.state_desc = 'ONLINE' AND d.name NOT IN ('tempdb', 'model', 'msdb')
"@

# AutoGrow query
$autoGrowQuery = @"
SELECT
    mf.name AS [File Name],
    mf.physical_name AS [Physical Name],
    mf.size * 8 / 1024 AS [Size (MB)],
    mf.growth * 8 / 1024 AS [AutoGrow Increment (MB)],
    CASE
        WHEN mf.growth = 0 THEN 'No AutoGrow'
        WHEN mf.is_percent_growth = 1 THEN 'Percentage-based growth'
        WHEN mf.is_percent_growth = 0 THEN 'Size-based growth'
        ELSE 'Unknown'
    END AS [Growth Type],
    mf.max_size AS [Max Size]
FROM sys.master_files mf
"@

# Trace Flag query
$traceflagquery = @"
-- Create a temporary table with the appropriate columns for storing trace status
CREATE TABLE #TraceStatus (
    TraceFlag INT,
    Status BIT,
	Global BIT,
    Session BIT
);

-- Insert the trace flag status into the temp table by executing DBCC TRACESTATUS
INSERT INTO #TraceStatus (TraceFlag, Status, Global, Session)
EXEC ('DBCC TRACESTATUS');

-- Query the temp table to view the active trace flags
SELECT TraceFlag FROM #TraceStatus;

-- Clean up the temporary table
DROP TABLE #TraceStatus;
GO
"@

# CHECKDB query
$checkDBQuery = @"
DECLARE @dbname NVARCHAR(256), @sql NVARCHAR(MAX);

-- Create a temporary table to store CHECKDB results
CREATE TABLE #CheckDBInfo (
    DatabaseName NVARCHAR(256),
    LastCheckDB DATETIME NULL
);

DECLARE db_cursor CURSOR FOR
SELECT name FROM sys.databases WHERE state_desc = 'ONLINE' AND name not in ('tempdb');

OPEN db_cursor;
FETCH NEXT FROM db_cursor INTO @dbname;

WHILE @@FETCH_STATUS = 0
BEGIN
    -- Create a temporary table to store DBCC DBINFO results
    CREATE TABLE #DBInfo (
        ParentObject NVARCHAR(255),
        Object NVARCHAR(255),
        Field NVARCHAR(255),
        Value NVARCHAR(255)
    );

    -- Build the dynamic SQL safely using QUOTENAME()
    SET @sql = 'DBCC DBINFO (' + QUOTENAME(@dbname) + ') WITH TABLERESULTS;';

    -- Insert DBCC DBINFO results into the temporary table
    INSERT INTO #DBInfo
    EXEC sp_executesql @sql;

    -- Extract the Last Known Good DBCC CHECKDB execution date
    INSERT INTO #CheckDBInfo (DatabaseName, LastCheckDB)
    SELECT @dbname,
           MAX(CASE WHEN Field = 'dbi_dbccLastKnownGood' THEN TRY_CAST(Value AS DATETIME) END)
    FROM #DBInfo;

    -- Drop temporary table for this iteration
    DROP TABLE #DBInfo;

    FETCH NEXT FROM db_cursor INTO @dbname;
END

CLOSE db_cursor;
DEALLOCATE db_cursor;

-- Retrieve the last CHECKDB execution date per database
SELECT DatabaseName,
    ISNULL(LastCheckDB, '1900-01-01') AS LastCheckDB
FROM #CheckDBInfo
ORDER BY LastCheckDB ASC;

-- Drop the final temporary table
DROP TABLE #CheckDBInfo;
"@

# VLF Query
$vlfQuery = @"
-- Create a temporary table to store VLF counts per database
CREATE TABLE #VLFInfo (
    DatabaseName SYSNAME,
    VLFCount INT
);

DECLARE @dbName SYSNAME, @sql NVARCHAR(MAX);

DECLARE db_cursor CURSOR FOR
SELECT name
FROM sys.databases
WHERE state_desc = 'ONLINE' AND name not in ('master','model','msdb','tempdb');

OPEN db_cursor;
FETCH NEXT FROM db_cursor INTO @dbName;

WHILE @@FETCH_STATUS = 0
BEGIN
    -- Build dynamic SQL to count VLFs per database using sys.dm_db_log_info
    SET @sql = N'
        USE ' + QUOTENAME(@dbName) + N';
        INSERT INTO #VLFInfo (DatabaseName, VLFCount)
        SELECT ''' + @dbName + N''', COUNT(*)
        FROM sys.dm_db_log_info(DB_ID());';

    EXEC sp_executesql @sql;

    FETCH NEXT FROM db_cursor INTO @dbName;
END

CLOSE db_cursor;
DEALLOCATE db_cursor;

-- Retrieve databases with more than 1000 VLFs
SELECT DatabaseName, VLFCount
FROM #VLFInfo
WHERE VLFCount > 1000
ORDER BY VLFCount DESC;

-- Drop temporary table after use
DROP TABLE #VLFInfo;
"@

# Backup query
$BkpQuery = @"
WITH BackupData AS (
    -- Get latest full backup and log backup for each database
    SELECT
        database_name,
        MAX(CASE WHEN type = 'D' THEN backup_finish_date ELSE NULL END) AS LastFullBackup,
        MAX(CASE WHEN type = 'L' THEN backup_finish_date ELSE NULL END) AS LastLogBackup
    FROM msdb.dbo.backupset
    GROUP BY database_name
)
SELECT
    d.name AS DatabaseName,
    d.recovery_model_desc AS RecoveryModel,
    ISNULL(bd.LastFullBackup, '1900-01-01 00:00:00.000') AS LastFullBackup,
    ISNULL(bd.LastLogBackup, '1900-01-01 00:00:00.000') AS LastLogBackup
FROM sys.databases d
LEFT JOIN BackupData bd ON d.name = bd.database_name
WHERE d.name NOT IN ('tempdb', 'model')
ORDER BY d.name;
"@

# MaxDop query
$maxdopquery = @"
DECLARE @sqlmajorver INT;
DECLARE @numa INT;
DECLARE @affined_cpus INT;
DECLARE @cpucount INT;
DECLARE @cpuaffin_fixed VARCHAR(300);
DECLARE @cpuaffin VARCHAR(300);
DECLARE @affinity64mask NVARCHAR(1024);
DECLARE @affinitymask NVARCHAR(1024);
DECLARE @recommended_maxdop INT;
DECLARE @current_maxdop INT;

-- Get the SQL Server version
SELECT @sqlmajorver = CONVERT(int, (@@microsoftversion / 0x1000000) & 0xff);

-- Count of processors (schedulers) for the current instance
SELECT @cpucount = COUNT(cpu_id) FROM sys.dm_os_schedulers WHERE scheduler_id < 255 AND parent_node_id < 64;

-- Count of NUMA nodes
SELECT @numa = COUNT(DISTINCT parent_node_id) FROM sys.dm_os_schedulers WHERE scheduler_id < 255 AND parent_node_id < 64;

-- Fetch the CPU affinity mask for current configuration
SELECT @cpuaffin = CASE WHEN @cpucount > 32 THEN @affinity64mask ELSE @affinitymask END;
SET @cpuaffin_fixed = @cpuaffin;

-- Fetch the number of CPUs available for SQL (online schedulers)
SELECT @affined_cpus = COUNT(cpu_id)
FROM sys.dm_os_schedulers
WHERE is_online = 1 AND scheduler_id < 255 AND parent_node_id < 64;

-- Select Recommended MaxDOP based on conditions
SELECT @recommended_maxdop =
    CASE
        -- If not NUMA, and up to 8 @affined_cpus then MaxDOP up to 8
        WHEN @numa = 1 AND @affined_cpus <= 8 THEN @affined_cpus
        -- If not NUMA, and more than 8 @affined_cpus then MaxDOP 8
        WHEN @numa = 1 AND @affined_cpus > 8 THEN 8
        -- If SQL 2016 or higher and has NUMA and # logical CPUs per NUMA up to 15, then MaxDOP is set as # logical CPUs per NUMA, up to 15
        WHEN @sqlmajorver >= 13 AND @numa > 1 AND CEILING(@cpucount*1.00/@numa) <= 15 THEN CEILING((@cpucount*1.00)/@numa)
        -- If SQL 2016 or higher and has NUMA and # logical CPUs per NUMA > 15, then MaxDOP is set as 1/2 of # logical CPUs per NUMA
        WHEN @sqlmajorver >= 13 AND @numa > 1 AND CEILING(@cpucount*1.00/@numa) > 15 THEN
            CASE WHEN CEILING(@cpucount*1.00/@numa/2) > 16 THEN 16 ELSE CEILING(@cpucount*1.00/@numa/2) END
        -- If up to SQL 2016 and has NUMA and # logical CPUs per NUMA up to 8, then MaxDOP is set as # logical CPUs per NUMA
        WHEN @sqlmajorver < 13 AND @numa > 1 AND CEILING(@cpucount*1.00/@numa) < 8 THEN CEILING(@cpucount*1.00/@numa)
        -- If up to SQL 2016 and has NUMA and # logical CPUs per NUMA > 8, then MaxDOP 8
        WHEN @sqlmajorver < 13 AND @numa > 1 AND CEILING(@cpucount*1.00/@numa) >= 8 THEN 8
        ELSE 0
    END;

-- Get the Current MaxDOP from SQL configuration, explicitly convert the value from sql_variant to int
SELECT @current_maxdop = CONVERT(INT, value)
FROM sys.configurations
WHERE name = 'max degree of parallelism';

-- Display the Recommended MaxDOP and Current MaxDOP
SELECT
    @recommended_maxdop AS [Recommended_MaxDOP],
    @current_maxdop AS [Current_MaxDOP]
"@

# TempDB Query
$tempDBFileSizeQuery = @"
SELECT
    name AS FileName,
    type_desc AS FileType,
    size * 8 / 1024 AS SizeMB,
    physical_name AS PhysicalPath
FROM sys.master_files
WHERE database_id = DB_ID('tempdb');
"@

# Query Store Query
$QueryStoreQuery = @"
DECLARE @major_version INT = CAST(LEFT(CAST(SERVERPROPERTY('ProductVersion') AS VARCHAR(20)),
                               CHARINDEX('.', CAST(SERVERPROPERTY('ProductVersion') AS VARCHAR(20))) - 1) AS INT);

IF @major_version >= 13
BEGIN
    IF OBJECT_ID('tempdb..#QueryStoreStatus') IS NOT NULL
        DROP TABLE #QueryStoreStatus;

    CREATE TABLE #QueryStoreStatus (
        database_name SYSNAME,
        query_store_status NVARCHAR(60),
        query_capture_mode INT,
        query_capture_mode_desc NVARCHAR(20),
        sql_major_version INT
    );

    DECLARE @sql NVARCHAR(MAX) = N'';

    SELECT @sql += '
    IF DB_ID(N''' + name + ''') IS NOT NULL
    BEGIN
        INSERT INTO #QueryStoreStatus (database_name, query_store_status, query_capture_mode, query_capture_mode_desc, sql_major_version)
        SELECT
            ''' + name + ''' AS database_name,
            actual_state_desc,
            query_capture_mode,
            CASE query_capture_mode
                WHEN 0 THEN ''OFF''
                WHEN 1 THEN ''ALL''
                WHEN 2 THEN ''AUTO''
                WHEN 3 THEN ''CUSTOM''
                ELSE ''UNKNOWN''
            END,
            ' + CAST(@major_version AS NVARCHAR) + ' AS sql_major_version
        FROM [' + name + '].sys.database_query_store_options;
    END;
    '
    FROM sys.databases
    WHERE database_id NOT IN (1,2,3);  -- exclude system DBs

    EXEC sp_executesql @sql;

    SELECT * FROM #QueryStoreStatus ORDER BY database_name;
END
ELSE
BEGIN
    SELECT 'There is no query store feature available on this SQL version' AS message;
END
"@

# ADR check
$ADRQuery = @"
DECLARE @major_version INT = CAST(LEFT(CAST(SERVERPROPERTY('ProductVersion') AS VARCHAR(20)),
                               CHARINDEX('.', CAST(SERVERPROPERTY('ProductVersion') AS VARCHAR(20))) - 1) AS INT);

IF @major_version >= 15
BEGIN
    DECLARE @sql NVARCHAR(MAX) = N'
    SELECT
        name AS database_name,
        is_accelerated_database_recovery_on,
        CASE is_accelerated_database_recovery_on
            WHEN 1 THEN ''ENABLED''
            WHEN 0 THEN ''DISABLED''
            ELSE ''UNKNOWN''
        END AS status
    FROM sys.databases
    WHERE database_id NOT IN (1,2,3,4);  -- exclude system DBs
    ';
    EXEC sp_executesql @sql;
END
ELSE
BEGIN
    SELECT 'This SQL version does not have Accelerated Database Recovery feature available' AS message;
END
"@

# Login/User Test query
$loginUserTestQuery = @"
DECLARE @dbname sysname;

CREATE TABLE #TestPrincipals
(
    PrincipalType NVARCHAR(10),
    PrincipalName sysname,
    DatabaseName  sysname NULL
);

-- 1. Logins with 'test' in the name
INSERT INTO #TestPrincipals (PrincipalType, PrincipalName, DatabaseName)
SELECT
    'Login' AS PrincipalType,
    sp.name AS PrincipalName,
    NULL    AS DatabaseName
FROM sys.server_principals AS sp
WHERE sp.name LIKE '%test%'
  AND sp.type IN ('S','U','G')       -- SQL login, Windows login, Windows group
  AND sp.name NOT LIKE '##%';        -- skip system auto-created logins

-- 2. Users with 'test' in the name across databases
DECLARE db_cursor CURSOR FAST_FORWARD FOR
    SELECT name
    FROM sys.databases
    WHERE state_desc = 'ONLINE'
      AND database_id > 4;           -- exclude system dbs; adjust if you want master/msdb etc.

OPEN db_cursor;
FETCH NEXT FROM db_cursor INTO @dbname;

WHILE @@FETCH_STATUS = 0
BEGIN
    DECLARE @sql NVARCHAR(MAX);

    SET @sql = N'
        INSERT INTO #TestPrincipals (PrincipalType, PrincipalName, DatabaseName)
        SELECT
            ''User'' AS PrincipalType,
            dp.name  AS PrincipalName,
            N''' + @dbname + N''' AS DatabaseName
        FROM ' + QUOTENAME(@dbname) + N'.sys.database_principals AS dp
        WHERE dp.name LIKE ''%test%''
          AND dp.type IN (''S'',''U'',''G'')
          AND dp.principal_id > 4;   -- skip dbo, guest, sys users
    ';

    EXEC (@sql);

    FETCH NEXT FROM db_cursor INTO @dbname;
END

CLOSE db_cursor;
DEALLOCATE db_cursor;

SELECT
    PrincipalType,
    PrincipalName,
    DatabaseName
FROM #TestPrincipals;

DROP TABLE #TestPrincipals;
"@


# Loop through each instance and execute the queries
foreach ($instanceName in $instanceNames) {

# Inline logic to retrieve SQL Server port from registry (no function)
try {
    # Define the registry base path for SQL Server
    $basePath = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server"

    # Initialize the port variable
    $port = "1433" # Default port

    # Dynamically find the instance ID that matches the given instance name
    $instanceID = Get-ChildItem -Path $basePath -ErrorAction Stop |
                  Where-Object { $_.PSChildName -match "^MSSQL.*\.$instanceName$" } |
                  Select-Object -ExpandProperty PSChildName -ErrorAction Stop

    if ($instanceID) {
        # Construct the full registry path for the TCP port
        $tcpKeyPath = "$basePath\$instanceID\MSSQLServer\SuperSocketNetLib\Tcp\IPAll"
        try {
            # Read the port from the registry
            $port = (Get-ItemProperty -Path $tcpKeyPath -Name TcpPort -ErrorAction Stop).TcpPort
            if (-not $port) {
                #Write-Host "No specific port found in the registry for instance $instanceName, assuming default port 1433."
                $port = "1433"
            }
            #Write-Host "SQL Server port for instance $instanceName is $port."
        } catch {
            #Write-Host "Error retrieving SQL Server port from registry path $tcpKeyPath. Assuming default port 1433."
            $port = "1433"
        }
    } else {
        #Write-Host "No instance ID found matching the name $instanceName. Assuming default port 1433."
        $port = "1433"
    }
} catch {
    #Write-Host "Error retrieving SQL Server port from registry for instance $instanceName. Assuming default port 1433."
    $port = "1433"
}

                    $sqlInstance = if ($instanceName -eq "MSSQLSERVER") { $vmName } else { "$vmName\$instanceName" }
                    if ($port -ne "1433") {
                        $sqlInstance = "$sqlInstance,$port"
                    }
                    #Write-Host "Connecting to SQL Server instance: $sqlInstance"

                    Try {
                        #Write-Host "Running queries on: $sqlInstance"

                        # Execute main config query
                        $currentQuery = "Main Query"
                        $mainResult = Invoke-Sqlcmd -ServerInstance $sqlInstance -Query $mainQuery -QueryTimeout 65535 -ErrorAction Stop

                        #Execute the compatibility level query
                        $currentQuery = "Compatibility Level Query"
                        $compatResult = Invoke-Sqlcmd -ServerInstance $sqlInstance -Query $compatQuery -QueryTimeout 65535 -ErrorAction Stop

                        # Execute the AutoGrow query
                        $currentQuery = "AutoGrow Query"
                        $autoGrowResult = Invoke-Sqlcmd -ServerInstance $sqlInstance -Query $autoGrowQuery -QueryTimeout 65535 -ErrorAction Stop

                        # Execute Trace Flag query
                        $currentQuery = "Trace Flag Query"
                        $enabledFlags = Invoke-Sqlcmd -ServerInstance $sqlInstance -Query $traceflagquery -QueryTimeout 65535 -ErrorAction Stop

                        # Execute CheckDB query
                        $currentQuery = "CheckDB Query"
                        $checkDBResult = Invoke-Sqlcmd -ServerInstance $sqlInstance -Query $checkDBQuery -QueryTimeout 65535 -ErrorAction Stop

                        # Execute VLFs query
                        $currentQuery = "VLFs Query"
                        $vlfResult = Invoke-Sqlcmd -ServerInstance $sqlInstance -Query $vlfQuery -QueryTimeout 65535 -ErrorAction Stop

                        # Execute Backup query
                        $currentQuery = "Backup Query"
                        $backupResult = Invoke-Sqlcmd -ServerInstance $sqlInstance -Query $BkpQuery -QueryTimeout 65535 -ErrorAction Stop

                        # Execute MaxDop Query
                        $currentQuery = "MaxDop Query"
                        $maxdopresult = Invoke-Sqlcmd -ServerInstance $sqlInstance -Query $maxdopquery -QueryTimeout 65535 -ErrorAction Stop

                        # Execute TempDB Query
                        $currentQuery = "TempDB Query"
                        $tempDBFileSizeResult = Invoke-Sqlcmd -ServerInstance $sqlInstance -Query $tempDBFileSizeQuery -QueryTimeout 65535 -ErrorAction Stop

                        # Execute Query Store Query
                        $currentQuery = "Query Store Query"
                        $QueryStoreResults = Invoke-Sqlcmd -ServerInstance $sqlInstance -Query $QueryStoreQuery -QueryTimeout 65535 -ErrorAction Stop

                        # Execute ADR Query
                        $currentQuery = "ADR Query"
                        $ADRQueryResults = Invoke-Sqlcmd -ServerInstance $sqlInstance -Query $ADRQuery -QueryTimeout 65535 -ErrorAction Stop

						# Execute Login/User Test Query
						$currentQuery = "Login/User Test Query"
						$loginUserTestResult = Invoke-Sqlcmd -ServerInstance $sqlInstance -Query $loginUserTestQuery -QueryTimeout 65535 -ErrorAction Stop

                    }
                    Catch
                    {
                        Write-Host "Error executing '$currentQuery' on SQL instance '$sqlInstance'. Error: $_"

                        # Assign empty arrays to prevent script failure due to missing results
                        $mainResult = @()
                        $compatResult = @()
                        $autoGrowResult = @()
                        $enabledFlags = @()
                        $checkDBResult = @()
                        $vlfResult = @()
                        $backupResult = @()
                        $maxdopresult = @()
                        $tempDBFileSizeResult = @()
                        $QueryStoreResults = @()
                        $ADRQueryResults = @()
                    }

                    # Process MaxDOP Logic
                    foreach ($result in $maxdopresult) {
                        # Ensure both values are available
                        $recommendedMaxDop = $result.Recommended_MaxDOP
                        $currentMaxDop = $result.Current_MaxDOP

                        # Compare recommended and current MaxDop values
                        if ($currentMaxDop -and $recommendedMaxDop) {
                            # Check if currentMaxDop is 0 or differs from recommendedMaxDop
                            $maxDopStatus = if ($currentMaxDop -eq 0 -or $currentMaxDop -ne $recommendedMaxDop) {
                                'REVIEW'
                            }
                            else {
                                'OK'
                            }
                        }
                        else {
                            # If no MaxDop values are present, mark as REVIEW
                            $maxDopStatus = 'REVIEW'
                        }
                    }

                    # Process Backups Logic
                    $databasesWithoutFullBackup = @()
                    $databasesWithoutLogBackup = @()

                    # Loop through each result from the SQL query
                    foreach ($database in $backupResult) {
                        # Check if LastFullBackup is valid and convert to DateTime
                        $lastFullBackup = if ($database.LastFullBackup) {
                            try {
                                [datetime]$database.LastFullBackup
                            }
                            catch {
                                $null
                            }
                        }
                        else {
                            $null
                        }

                        # Check if LastLogBackup is valid and convert to DateTime
                        $lastLogBackup = if ($database.LastLogBackup) {
                            try {
                                [datetime]$database.LastLogBackup
                            }
                            catch {
                                $null
                            }
                        }
                        else {
                            $null
                        }

                        # Full Backup Status logic: If LastFullBackup is older than 7 days or NULL
                        $fullBackupStatus = if ($lastFullBackup -eq $null -or $lastFullBackup -lt (Get-Date).AddDays(-7)) {
                            'REVIEW'
                        }
                        else {
                            'OK'
                        }

                        # Log Backup Status logic: For FULL recovery model, if LastLogBackup is NULL or older than 24 hours
                        $logBackupStatus = if (($database.RecoveryModel -eq 'FULL' -or $database.RecoveryModel -eq 'BULK_LOGGED') -and
                        ($lastLogBackup -eq $null -or $lastLogBackup -lt (Get-Date).AddHours(-24))) {
                            'REVIEW'
                        }
                        else {
                            'OK'
                        }

                        # If the full backup status is REVIEW, add it to the list of databases needing a full backup
                        if ($fullBackupStatus -eq 'REVIEW') {
                            $databasesWithoutFullBackup += $database.DatabaseName
                        }

                        # If the log backup status is REVIEW, add it to the list of databases needing a log backup
                        if ($logBackupStatus -eq 'REVIEW') {
                            $databasesWithoutLogBackup += $database.DatabaseName
                        }

                        # Add FullBackupStatus and LogBackupStatus to the result object
                        $database | Add-Member -MemberType NoteProperty -Name "FullBackupStatus" -Value $fullBackupStatus
                        $database | Add-Member -MemberType NoteProperty -Name "LogBackupStatus" -Value $logBackupStatus
                    }

                    # Determine overall status for reporting
                    $fullBackupStatusMessage = if ($databasesWithoutFullBackup.Count -gt 0) {
                        "DBs: " + ($databasesWithoutFullBackup -join ', ')
                    }
                    else {
                        "All databases have recent full backups."
                    }

                    # Check if all databases are in SIMPLE recovery model
                    $nonSimpleDatabases = $backupResult | Where-Object { $_.RecoveryModel -ne 'SIMPLE' }

                    # If all databases are SIMPLE, set a special message
                    if ($databasesWithoutLogBackup.Count -eq 0 -and $nonSimpleDatabases.Count -eq 0) {
                        $logBackupStatusMessage = "Not applicable because the databases are using the SIMPLE recovery model."
                    }
                    elseif ($databasesWithoutLogBackup.Count -gt 0) {
                        $logBackupStatusMessage = "DBs: " + ($databasesWithoutLogBackup -join ', ')
                    }
                    else {
                        $logBackupStatusMessage = "All databases have recent log backups."
                    }


                    # Process VLF results
                    $databasesWithHighVLFs = @($vlfResult | Where-Object { $_.VLFCount -gt 1000 } | Select-Object -ExpandProperty DatabaseName)

                    # Determine VLF Status
                    $vlfStatus = if ($databasesWithHighVLFs.Count -gt 0) { "REVIEW" } else { "OK" }
                    $vlfMessage = if ($databasesWithHighVLFs.Count -eq 0) { "All databases are OK" } else { $databasesWithHighVLFs -join ', ' }

                    # Determine databases missing CHECKDB within the last 7 days
                    $sevenDaysAgo = (Get-Date).AddDays(-7)
                    $missingCheckDB = @($checkDBResult | Where-Object { $_.LastCheckDB -lt $sevenDaysAgo -or $_.LastCheckDB -eq $null } | Select-Object -ExpandProperty DatabaseName)

                    # Determine CHECKDB Status
                    $checkDBStatus = if ($missingCheckDB.Count -gt 0) { "REVIEW" } else { "OK" }
                    $missingCheckDBMessage = if ($missingCheckDB.Count -eq 0) { "All databases are OK" } else { $missingCheckDB -join ', ' }

					# ------------------------------------------
					# Login/User 'test' validation logic
					# ------------------------------------------
					$testLogins = @(
						$loginUserTestResult |
						Where-Object { $_.PrincipalType -eq 'Login' } |
						Select-Object -ExpandProperty PrincipalName -ErrorAction SilentlyContinue
					)

					$testUsersFormatted = @(
						$loginUserTestResult |
						Where-Object { $_.PrincipalType -eq 'User' } |
						ForEach-Object {
							# Format as dbname_users: username
							"$($_.DatabaseName)_users: $($_.PrincipalName)"
						}
					)

					if ($testLogins.Count -gt 0 -or $testUsersFormatted.Count -gt 0) {
						$loginUserTestStatus = "REVIEW"
					} else {
						$loginUserTestStatus = "OK"
					}

					# Build the combined descriptive string
					$loginUserTestValidation = ""

					if ($testLogins.Count -gt 0) {
						$loginUserTestValidation += "Logins: " + ($testLogins -join ', ')
					}

					if ($testUsersFormatted.Count -gt 0) {
						if (-not [string]::IsNullOrEmpty($loginUserTestValidation)) {
							$loginUserTestValidation += " | "
						}
						$loginUserTestValidation += "Users: " + ($testUsersFormatted -join '; ')
					}

					if ([string]::IsNullOrEmpty($loginUserTestValidation)) {
						$loginUserTestValidation = "No logins or users with 'test' found"
					}
					
                     # Process the main query results
                    if (($mainResult -ne $null -and $mainResult.Count -gt 0) -or
                        ($compatResult -ne $null -and $compatResult.Count -gt 0) -or
                        ($autoGrowResult -ne $null -and $autoGrowResult.Count -gt 0)) {
                        foreach ($result in $mainResult) {
                            $serverNameWithInstance = $result.'Server Name'  # Example: ServerName\INSTANCE_NAME
                            $serverName = $serverNameWithInstance -replace '\\.*$', ''  # Keep only the server name part
                            $instanceNameOnly = $serverNameWithInstance -replace '^[^\\]*\\', ''  # Extract instance name

                            # If the instance name is the same as the server name, set it to "DEFAULT"
                            if ($instanceNameOnly -eq $serverName) {
                                $instanceNameOnly = 'DEFAULT'
                            }

                            $resultObject = [PSCustomObject]@{
                                "Server Name"                                 = $serverName
                                "SQL Instance Name"                           = $instanceNameOnly
                                "SQL Server Version"                          = switch -Wildcard ($result.'SQL Build Number') {
                                    "11*" { "SQL Server 2012" }
                                    "12*" { "SQL Server 2014" }
                                    "13*" { "SQL Server 2016" }
                                    "14*" { "SQL Server 2017" }
                                    "15*" { "SQL Server 2019" }
                                    "16*" { "SQL Server 2022" }
                                    "17*" { "SQL Server 2025" }
                                    default { "Unknown Version" }
                                }
                                "SQL Build Number"                            = $result.'SQL Build Number'
                                "SQL Edition"                                 = $result.'SQL Edition'
                                "Total Server Memory (MB)"                    = $result.'Total Server Memory (MB)'
                                "Current Min Server Memory (MB)"              = $result.'Min Server Memory (MB)'
                                "Current Max Server Memory (MB)"              = $result.'Max Server Memory (MB)'
                                "Total Visible Processors"                    = $result.'Total Visible Processors'
                                "Optimize for Ad Hoc Workloads"               = $result.'Optimize for Ad Hoc Workloads'
                                "Backup Compression Default"                  = if ([string]::IsNullOrEmpty($result.'Backup Compression Default')) {
                                    "Backup compression not available, check SQL server edition"
                                }
                                else {
                                    $result.'Backup Compression Default'
                                }
                                "Remote Admin Connections"                    = $result.'Remote Admin Connections'
                                "TempDB Data Files"                           = $result.'TempDB Data Files'
                                "DBs with missing CHECKDB in the last 7 days" = $missingCheckDBMessage
                                "CHECKDB Status"                              = $checkDBStatus
								"Login or User Test Status"                   = $loginUserTestStatus
								"Login and Users validation"                  = $loginUserTestValidation
                            }

                            # Trace Flag logic
                            $traceFlagsByVersion = @{
                                '11' = @(4199, 1118)                              # SQL Server 2012
                                '12' = @(4199, 1118)                              # SQL Server 2014
                                '13' = @(4199, 7745)                              # SQL Server 2016
                                '14' = @(4199, 7745, 12310)                       # SQL Server 2017
                                '15' = @(4199, 7745, 12310)                       # SQL Server 2019
                                '16' = @(4199, 7745, 12656, 12618)                # SQL Server 2022 and above
                            }

                            # Extract the major version number from the SQL version string
                            $majorVersion = ($result.'SQL Build Number' -split '\.')[0]
                            #Write-Host "Detected SQL Server Major Version: $majorVersion"

                            # Initialize trace flag status
                            $traceFlagStatus = "OK"
                            $enabledFlagNames = @()  # Initialize as empty

                            # Assume $enabledFlags is already populated from the trace flag query
                            if ($enabledFlags -and $enabledFlags.TraceFlag) {
                                $enabledFlagNames = $enabledFlags.TraceFlag | ForEach-Object { [int]$_ }  # Ensure trace flags are integers
                            }

                            # Determine the required trace flags for the current SQL version
                            $traceFlagList = $null
                            $majorVersionInt = [int]$majorVersion

                            if ($majorVersionInt -ge 16) {
                                $traceFlagList = $traceFlagsByVersion['16']  # Use version 16 flags for 16 and above
                            } elseif ($traceFlagsByVersion.ContainsKey($majorVersion)) {
                                $traceFlagList = $traceFlagsByVersion[$majorVersion]
                            } else {
                                #Write-Host "No trace flags are defined for version $majorVersion" -ForegroundColor Red
                                $traceFlagStatus = "REVIEW"
                            }

                            # If trace flags for the current SQL version are found
                            if ($traceFlagList) {
                                # Check if any required trace flags are missing
                                $missingFlags = $traceFlagList | Where-Object { [int]$_ -notin $enabledFlagNames }

                                # If there are missing flags
                                if ($missingFlags.Count -gt 0) {
                                    $traceFlagStatus = "REVIEW"
                                    #Write-Host "Missing Trace Flags: $([string]::Join(', ', $missingFlags))" -ForegroundColor Yellow
                                } else {
                                    #Write-Host "All required trace flags are present: $([string]::Join(', ', $traceFlagList))" -ForegroundColor Green
                                }
                            } else {
                                # If no trace flags are defined for the current version
                                $traceFlagStatus = "REVIEW"
                                #Write-Host "No trace flags are defined for version $majorVersion" -ForegroundColor Red
                            }

                            # Set the result for Trace Flag List
                            $traceFlagListString = if ($enabledFlagNames.Count -gt 0) {
                                [string]::Join(", ", $enabledFlagNames)
                            } else {
                                "No trace flags enabled"
                            }


                            # Calculate recommended max memory (75% of total server memory)
                            $recommendedMaxMemory = [math]::Round($result.'Total Server Memory (MB)' * 0.75, 0)

                            # Initialize memoryStatus as "REVIEW"
                            $memoryStatus = "REVIEW"

                            # Check if Max Server Memory is unconfigured (2147483647)
                            if ($result.'Max Server Memory (MB)' -eq 2147483647) {
                                $memoryStatus = "REVIEW"
                            }
                            # Check if Max Server Memory is greater than or equal to the total memory
                            elseif ($result.'Max Server Memory (MB)' -ge $result.'Total Server Memory (MB)') {
                                $memoryStatus = "REVIEW"
                            }
                            elseif ($result.'Min Server Memory (MB)' -eq 1024) {
                                # If Min Server Memory is 1024, check Max Server Memory (should be at least 75% of total memory)
                                if ($result.'Max Server Memory (MB)' -ge $recommendedMaxMemory) {
                                    $memoryStatus = "OK"
                                }
                            }
                            else {
                                # If Min Server Memory is not 1024, set to REVIEW
                                $memoryStatus = "REVIEW"
                            }

                            # Check if any configurations are not optimal, skipping 'Backup Compression Default' if it's null or empty
                            $configStatus = if (
                                $result.'Optimize for Ad Hoc Workloads' -eq 0 -or
                                $result.'Remote Admin Connections' -eq 0 -or
                                # If Backup Compression is not null or empty and not set to the expected value, flag as REVIEW
                                (![string]::IsNullOrEmpty($result.'Backup Compression Default') -and $result.'Backup Compression Default' -lt 1)
                            ) {
                                "REVIEW"
                            }
                            elseif ([string]::IsNullOrEmpty($result.'Backup Compression Default') -and
                                $result.'Optimize for Ad Hoc Workloads' -ne 0 -and
                                $result.'Remote Admin Connections' -ne 0) {
                                # If Backup Compression is null or empty and other settings are good, status is OK
                                "OK"
                            }
                            else {
                                # Default to OK if all other configurations are correct
                                "OK"
                            }

                            # Compatibility Level logic (filter out native compatibility level)
                            $nativeCompatibilityLevel = switch ($majorVersion) {
                                '11' { 110 }  # SQL Server 2012
                                '12' { 120 }  # SQL Server 2014
                                '13' { 130 }  # SQL Server 2016
                                '14' { 140 }  # SQL Server 2017
                                '15' { 150 }  # SQL Server 2019
                                '16' { 160 }  # SQL Server 2022
                                '17' { 170 }  # SQL Server 2025
                                default { 0 } # Unknown version
                            }
                            $compatibilityLevels = $compatResult | Where-Object {
                                $_.'Compatibility Level' -ne $nativeCompatibilityLevel
                            }
                            $compatLevels = $compatibilityLevels | ForEach-Object {
                                "$($_.'Database Name') (Level $($_.'Compatibility Level'))"
                            }

                            # Set message for databases out of native compatibility level
                            $compatLevelsMessage = if ($compatLevels.Count -eq 0) {
                                "All databases are in native compatibility level!"
                            }
                            else {
                                $compatLevels -join ', '
                            }

                            # Checks for Auto Create Stats, Auto Update Stats, and Page Verify
                            $autoUpdateStatsDatabases = $compatResult | Where-Object { $_.'Auto Update Stats' -eq 0 }
                            $autoCreateStatsDatabases = $compatResult | Where-Object { $_.'Auto Create Stats' -eq 0 }
			    $autoShrink = $compatResult | Where-Object { $_.'Auto Shrink' -eq 1 }
                            $pageVerifyDatabases = $compatResult | Where-Object { $_.'Page Verify' -ne "CHECKSUM" }

                            # List databases without proper settings
                            $divergentDatabases = @()
                            $divergentDatabases += $autoUpdateStatsDatabases | ForEach-Object { "$($_.'Database Name') (Auto Update Stats OFF)" }
                            $divergentDatabases += $autoCreateStatsDatabases | ForEach-Object { "$($_.'Database Name') (Auto Create Stats OFF)" }
			    $divergentDatabases += $autoShrink | ForEach-Object { "$($_.'Database Name') (Auto Shrink ON)" }
                            $divergentDatabases += $pageVerifyDatabases | ForEach-Object { "$($_.'Database Name') (Page Verify NOT CHECKSUM)" }

                            $divergentDatabasesMessage = if ($divergentDatabases.Count -eq 0) {
                                "All settings are OK!"
                            }
                            else {
                                $divergentDatabases -join ', '
                            }

                            # Database Options Status logic
                            $databaseOptionsStatus = if ($divergentDatabases.Count -gt 0) {
                                "REVIEW"
                            }
                            else {
                                "OK"
                            }

                            # Compatibility Level Status logic
                            $compatibilityLevelStatus = if ($compatLevels.Count -gt 0) {
                                "REVIEW"
                            }
                            else {
                                "OK"
                            }

                            # Process the result
                            $unlimitedAutoGrowFiles = @()
                            $percentageAutoGrowFiles = @()
                            $largeIncrementAutoGrowFiles = @()


                            foreach ($file in $autoGrowResult) {
                                # Check for unlimited AutoGrow
                                if ($file.'Max Size' -eq -1) {
                                    $unlimitedAutoGrowFiles += $file.'File Name'
                                }

                                # Check for AutoGrow by percentage
                                if ($file.'Growth Type' -eq 'Percentage-based growth') {
                                    $percentageAutoGrowFiles += $file.'File Name'
                                }

                                # Check for large increments (growth > 1024MB)
                                if ($file.'AutoGrow Increment (MB)' -gt 1024) {
                                    $largeIncrementAutoGrowFiles += $file.'File Name'
                                }
                            }

                            # Format the messages for AutoGrow issues
                            $unlimitedAutoGrowMessage = if ($unlimitedAutoGrowFiles.Count -eq 0) {
                                "No files have unlimited AutoGrow."
                            }
                            else {
                                "Unlimited AutoGrow: " + ($unlimitedAutoGrowFiles -join ', ') + ","
                            }

                            $percentageAutoGrowMessage = if ($percentageAutoGrowFiles.Count -eq 0) {
                                "No files are using AutoGrow by percentage."
                            }
                            else {
                                "AutoGrow by percentage: " + ($percentageAutoGrowFiles -join ', ') + ","
                            }

                            $largeIncrementAutoGrowMessage = if ($largeIncrementAutoGrowFiles.Count -eq 0) {
                                "No files have large increments for AutoGrow."
                            }
                            else {
                                "Large Increment AutoGrow: " + ($largeIncrementAutoGrowFiles -join ', ') + ","
                            }

                            # Check if any issues exist for AutoGrowth and set AutoGrowth Status
                            $autoGrowthIssues = $unlimitedAutoGrowFiles.Count + $percentageAutoGrowFiles.Count + $largeIncrementAutoGrowFiles.Count
                            $autoGrowthStatus = if ($autoGrowthIssues -gt 0) { "REVIEW" } else { "OK" }

                            # Query Store Checks
                            if ($QueryStoreResults -and $QueryStoreResults.Count -gt 0) {
                                if ($QueryStoreResults[0].PSObject.Properties.Name -contains "message" -and
                                    $QueryStoreResults[0].message -like "*no query store feature*") {

                                    $QueryStoreStatus = "OK"
                                    $QueryStoreDatabases = "There is no query store feature available on this SQL version"
                                }
                                else {
                                    $QueryStoreEnabled = $QueryStoreResults | Where-Object { $_.query_store_status -ne $null }

                                    if ($QueryStoreEnabled.Count -gt 0) {
                                        # Format as: dbname: status / capture_mode
                                        $QueryStoreDatabases = ($QueryStoreEnabled | ForEach-Object {
                                            "$($_.database_name): $($_.query_store_status) / $($_.query_capture_mode_desc)"
                                        }) -join ", "

                                        # If any are OFF, mark status as REVIEW
                                        if ($QueryStoreEnabled | Where-Object { $_.query_store_status -eq "OFF" }) {
                                            $QueryStoreStatus = "REVIEW"
                                        }
                                        else {
                                            $QueryStoreStatus = "OK"
                                        }
                                    }
                                    else {
                                        $QueryStoreStatus = "REVIEW"
                                        $QueryStoreDatabases = "No databases with Query Store enabled"
                                    }
                                }
                            }
                            else {
                                $QueryStoreStatus = "OK"
                                $QueryStoreDatabases = "This SQL version don't have Query Store feature available"
                            }

                            # ADR Check
                            if ($ADRQueryResults -and $ADRQueryResults.Count -gt 0) {
                                # Case: SQL version does not support ADR
                                if ($ADRQueryResults[0].PSObject.Properties.Name -contains "message" -and
                                    $ADRQueryResults[0].message -like "*feature available*") {

                                    $ADRStatus = "OK"
                                    $ADRDetails = "This SQL version does not have Accelerated Database Recovery feature available"
                                }
                                else {
                                    # Format all database ADR statuses
                                    $ADRDetails = ($ADRQueryResults | ForEach-Object {
                                        "$($_.database_name): $($_.status)"
                                    }) -join ", "

                                    # REVIEW if any database has DISABLED status
                                    if ($ADRQueryResults | Where-Object { $_.status -eq "DISABLED" }) {
                                        $ADRStatus = "REVIEW"
                                    }
                                    else {
                                        $ADRStatus = "OK"
                                    }
                                }
                            }
                            else {
                                $ADRStatus = "OK"
                                $ADRDetails = "This SQL version does not have Accelerated Database Recovery feature available"
                            }

                            #TempDB checks
                            # Extract number of processors
                            $totalProcessors = [int]$result.'Total Visible Processors'

                            # Extract TempDB data file count
                            $tempDBDataFiles = @($tempDBFileSizeResult | Where-Object { $_.FileType -eq "ROWS" })
                            $totalTempDBDataFiles = $tempDBDataFiles.Count

                            # Extract TempDB file sizes and check for uniformity
                            $tempDBFileSizes = $tempDBDataFiles | Select-Object -ExpandProperty SizeMB | Sort-Object -Unique
                            $allFilesSameSize = ($tempDBFileSizes.Count -eq 1)

                            # Check if TempDB data files count is a multiple of 4
                            $tempDBMultipleOf4 = ($totalTempDBDataFiles % 4 -eq 0)

                            # Determine recommended TempDB file count range based on processors
                            if ($totalProcessors -gt 8) {
                                $recommendedTempDBFilesMin = 4
                                $recommendedTempDBFilesMax = 8
                            } elseif ($totalProcessors -eq 8) {
                                $recommendedTempDBFilesMin = 4
                                $recommendedTempDBFilesMax = 8
                            } elseif ($totalProcessors -eq 4) {
                                $recommendedTempDBFilesMin = 2
                                $recommendedTempDBFilesMax = 4
                            } else {
                                $recommendedTempDBFilesMin = 1
                                $recommendedTempDBFilesMax = 1
                            }

                            # Evaluate TempDB status
                            $tempDBStatus = "OK"

                            # SQL Server 2022+ (version 16+) special case: 1 file is acceptable, even if below min
                            $singleFileOkay = ($majorVersion -ge 16 -and $totalTempDBDataFiles -eq 1)

                            if (
                                (-not $singleFileOkay -and (
                                    $totalTempDBDataFiles -lt $recommendedTempDBFilesMin -or
                                    $totalTempDBDataFiles -gt $recommendedTempDBFilesMax
                                )) -or
                                -not $allFilesSameSize -or
                                (-not $tempDBMultipleOf4 -and $totalTempDBDataFiles -gt 1)
                            ) {
                                $tempDBStatus = "REVIEW"
                            }

                            # Determine uniform size description
                            if ($totalTempDBDataFiles -eq 1) {
                                $tempDBUniformSize = "Nothing to compare because TempDB has only one data file"
                            }
                            elseif (-not $allFilesSameSize) {
                                $tempDBFileDetails = $tempDBDataFiles | ForEach-Object { "$($_.FileName): $($_.SizeMB) MB" }
                                $tempDBUniformSize = $tempDBFileDetails -join ", "
                            }
                            else {
                                $tempDBUniformSize = "All data files have the same size"
                            }

                            # Add new properties to result object
                            $resultObject | Add-Member -MemberType NoteProperty -Name "Memory Status" -Value $memoryStatus
                            $resultObject | Add-Member -MemberType NoteProperty -Name "Config Status" -Value $configStatus
                            $resultObject | Add-Member -MemberType NoteProperty -Name "MaxDop Status" -Value $maxDopStatus
                            $resultObject | Add-Member -MemberType NoteProperty -Name "Databases out of native compatibility" -Value $compatLevelsMessage
                            $resultObject | Add-Member -MemberType NoteProperty -Name "Unlimited AutoGrow" -Value $unlimitedAutoGrowMessage
                            $resultObject | Add-Member -MemberType NoteProperty -Name "AutoGrow by Percentage" -Value $percentageAutoGrowMessage
                            $resultObject | Add-Member -MemberType NoteProperty -Name "Large Increment AutoGrow" -Value $largeIncrementAutoGrowMessage
                            $resultObject | Add-Member -MemberType NoteProperty -Name "Auto Growth Status" -Value $autoGrowthStatus
                            $resultObject | Add-Member -MemberType NoteProperty -Name "Compatibility Level Status" -Value $compatibilityLevelStatus
                            $resultObject | Add-Member -MemberType NoteProperty -Name "Database Options Status" -Value $databaseOptionsStatus
                            $resultObject | Add-Member -MemberType NoteProperty -Name "Database Options Divergence" -Value $divergentDatabasesMessage
                            $resultObject | Add-Member -MemberType NoteProperty -Name "Trace Flag List" -Value $traceFlagListString
                            $resultObject | Add-Member -MemberType NoteProperty -Name "Trace Flag Status" -Value $traceFlagStatus
                            $resultObject | Add-Member -MemberType NoteProperty -Name "DBs with too many VLFs" -Value $vlfMessage
                            $resultObject | Add-Member -MemberType NoteProperty -Name "VLF Status" -Value $vlfStatus
                            $resultObject | Add-Member -MemberType NoteProperty -Name "Full Backup Status" -Value $fullBackupStatus
                            $resultObject | Add-Member -MemberType NoteProperty -Name "Log Backup Status" -Value $logBackupStatus
                            $resultObject | Add-Member -MemberType NoteProperty -Name "DBs missing Full Backup in the last 7 days" -Value $fullBackupStatusMessage
                            $resultObject | Add-Member -MemberType NoteProperty -Name "DBs missing Log Backup (With full rec model)" -Value $logBackupStatusMessage
                            $resultObject | Add-Member -MemberType NoteProperty -Name "Recommended Min Server Memory (MB)" -Value 1024
                            $resultObject | Add-Member -MemberType NoteProperty -Name "Recommended Max Server Memory (MB)" -Value $recommendedMaxMemory
                            $resultObject | Add-Member -MemberType NoteProperty -Name "Current Max Dop" -Value $currentMaxDop
                            $resultObject | Add-Member -MemberType NoteProperty -Name "Recommended Max Dop" -Value $recommendedMaxDop
                            $resultObject | Add-Member -MemberType NoteProperty -Name "Query Store Status" -Value $QueryStoreStatus
                            $resultObject | Add-Member -MemberType NoteProperty -Name "Query Store Details" -Value $QueryStoreDatabases
                            $resultObject | Add-Member -MemberType NoteProperty -Name "ADR Status" -Value $ADRStatus
                            $resultObject | Add-Member -MemberType NoteProperty -Name "ADR Details" -Value $ADRDetails
                            $resultObject | Add-Member -MemberType NoteProperty -Name "TempDB Status" -Value $tempDBStatus
                            $resultObject | Add-Member -MemberType NoteProperty -Name "TempDB Data Files Count" -Value $totalTempDBDataFiles
                            $resultObject | Add-Member -MemberType NoteProperty -Name "TempDB Data Files Size" -Value $tempDBUniformSize

                            # Add to jobResults
                            $jobResults += $resultObject
                        }
                    }
                }

                # Return all results
                return $jobResults
            }
            Catch {
                Write-Host "Error during SQL query execution: $_"
                return @()
            }
        } -ArgumentList $VMName -AsJob

        $Jobs += $Job
    }
    Catch {
        Write-Host "Failed to connect to $($VMs): $_"
    }
}


# Ensure the progress bar reaches 100% when completed
Write-Progress -Activity "Processing SQL Servers" -Status "Completed" -PercentComplete 100 -Completed

# Wait for all jobs to complete (only if jobs exist)
if ($Jobs.Count -gt 0) {
    Write-Host "Waiting for all jobs to finish..."
    Wait-Job -Job $Jobs
}
else {
    Write-Host "No jobs were created. Exiting..."
    # Stop capturing console output
    Stop-Transcript
    Exit 1  # Exit with an error code to indicate failure
}

# Receive results and store them
$Jobs | ForEach-Object {
    $JobResult = Receive-Job -Job $_
    if ($JobResult) {
        # Add results to the array
        $Results += $JobResult
    }
}

# Remove completed jobs
$Jobs | Remove-Job

# Prepare the column order to bring the status columns to the front
$columnOrder = @(
    "Server Name",
    "SQL Instance Name",
    "SQL Server Version",
    "SQL Build Number",
    "SQL Edition",
    "Memory Status",
    "Config Status",
    "MaxDop Status",
    "Query Store Status",
    "ADR Status",
    "TempDB Status",
    "Auto Growth Status",
    "Database Options Status",
    "Compatibility Level Status",
    "Trace Flag Status",
    "CHECKDB Status",
    "VLF Status",
    "Full Backup Status",
    "Log Backup Status",
	"Login or User Test Status",  
    "Total Server Memory (MB)",
    "Current Min Server Memory (MB)",
    "Recommended Min Server Memory (MB)",
    "Current Max Server Memory (MB)",
    "Recommended Max Server Memory (MB)",
    "Total Visible Processors",
    "Current Max Dop",
    "Recommended Max Dop",
    "Optimize for Ad Hoc Workloads",
    "Backup Compression Default",
    "Remote Admin Connections",
    "Databases out of native compatibility",
    "Database Options Divergence",
    "Unlimited AutoGrow",
    "AutoGrow by Percentage",
    "Large Increment AutoGrow",
    "Trace Flag List",
    "DBs with missing CHECKDB in the last 7 days",
    "DBs with too many VLFs",
    "DBs missing Full Backup in the last 7 days",
    "DBs missing Log Backup (With full rec model)",
	"Login and Users validation",  
    "Query Store Details",
    "ADR Details",
    "TempDB Data Files Count",
    "TempDB Data Files Size"
)

# Remove unwanted columns from the results before exporting to CSV
$filteredResults = $Results | Select-Object -Property $columnOrder -ExcludeProperty PSComputerName, RunspaceId, PSShowComputerName

# Remove duplicate rows based on "Server Name" and "SQL Instance Name"
$uniqueResults = $filteredResults | Group-Object -Property "Server Name", "SQL Instance Name" | ForEach-Object { $_.Group | Select-Object -First 1 }

# Generate Summary of OK and REVIEW Counts by Status Column
# Prepare CSV Output
$executionDate = Get-Date -Format "yyyyMMdd_HHmmss"
$filePath = "C:\temp\SQL_Server_Best_Practices_Assessment_Results_$executionDate.csv"

# Export filtered results to CSV with the dynamic filename
$uniqueResults | Export-Csv -Path $filePath -NoTypeInformation

# Generate Summary of OK and REVIEW Counts by Status Column
Write-Host "Generating Summary of Assessment Results..."

# Define columns that contain status checks
$statusColumns = @("Memory Status", "Config Status", "MaxDop Status", "Query Store Status", "ADR Status", "TempDB Status", "Auto Growth Status",
    "Database Options Status", "Compatibility Level Status", "Trace Flag Status", "CHECKDB Status",
    "VLF Status", "Full Backup Status", "Log Backup Status", "Login or User Test Status")

# Initialize counters per column
$statusSummary = @{}
foreach ($column in $statusColumns) {
    $statusSummary[$column] = @{OK = 0; REVIEW = 0 }
}

# Iterate over each result to count OKs and REVIEWs by column
foreach ($row in $uniqueResults) {
    foreach ($column in $statusColumns) {
        if ($row.$column -eq "OK") {
            $statusSummary[$column]["OK"]++
        }
        elseif ($row.$column -eq "REVIEW") {
            $statusSummary[$column]["REVIEW"]++
        }
    }
}

# Determine maximum column width for alignment
$maxStatusLength = ($statusColumns | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum

# Print Summary Table
Write-Host "----------------------------------------------------------" -ForegroundColor Cyan
Write-Host " ===== SQL Server Best Practices Assessment Summary ===== " -ForegroundColor Cyan
Write-Host "----------------------------------------------------------" -ForegroundColor Cyan
Write-Host ("{0,-$maxStatusLength} | {1,5} | {2,5}" -f "STATUS", "OK", "REVIEW") -ForegroundColor Yellow
Write-Host ("-" * ($maxStatusLength + 20))

foreach ($column in $statusColumns) {
    $okCount = $statusSummary[$column]["OK"]
    $reviewCount = $statusSummary[$column]["REVIEW"]
    Write-Host ("{0,-$maxStatusLength} | {1,5} | {2,5}" -f $column, $okCount, $reviewCount)
}

Write-Host ("=========================================================") -ForegroundColor Cyan

Write-Host "CSV containing all findings details is located in: $filePath" -ForegroundColor Green

# Stop capturing console output
Stop-Transcript
