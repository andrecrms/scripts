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
# Instance settings, checkdb history, backup history, VLFs, autogrowth, trace flags, tempdb file checks,logins and users with word "test", database options, Always On and compatibility levels.
# This script uses some queries from the BPCheck script to evaluate some best practices: https://github.com/microsoft/tigertoolbox/tree/master/BPCheck
# A quick summary of the findings will be presented at the end of the script, but all detailed results will be placed in a CSV file.
# Each status column will be either OK or REVIEW, when marked as REVIEW, navigate to the right of the sheet to understand why it was marked that way.
# A log of the execution will also be generated in the same directory as the csv file mentioned above.
# Tested on: SQL Server 2019 to 2025.
#
# Author: Andre Cesar Rodrigues
# LinkedIn: https://www.linkedin.com/in/andre-c-rodrigues
# Blog: http://sqlmagu.blogspot.com.br
# GitHub: https://github.com/andrecrms
# Last modified: 05/11/2026.
=============================================================================================================================================================================================
"@ -ForegroundColor Yellow
Write-Host @"
DISCLAIMER: This script should be tested in an appropriate environment before running in production. Additionally, properly validate your results as each environment may
have its own characteristics. This script will not change nothing in the environment, it will just run some SQL Queries to collect all necessary information.
"@ -ForegroundColor Red

$ServerName = Read-Host "Enter the server name (or press Enter to use the list from C:\temp\serverlist.txt)"

if ($ServerName -eq ".") {
    $ServerName = "localhost"
}

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

while ($true) {
    $UseFQDN = (Read-Host "Do you want to use the full domain name (yes/no)?").ToLower()
    if ($UseFQDN -in @("yes", "no")) { break }
    Write-Host "Invalid input. Please type 'yes' or 'no'." -ForegroundColor Red
}

$DomainName = ""
if ($UseFQDN -eq "yes") {
    $DomainName = Read-Host "Enter the domain name (e.g., contoso.com)"
}

function IsLocalMachine {
    param([string]$Name)
    $localNames = @("localhost", "127.0.0.1", ".", $env:COMPUTERNAME)
    return ($localNames -icontains $Name)
}

$CoreScriptBlock = {
    param($vmName, $baseServerName)

    Try {
        # Prefer the modern SqlServer module; fall back to legacy sqlps
        Import-Module SqlServer -DisableNameChecking -ErrorAction SilentlyContinue
        if (-not (Get-Module -Name SqlServer -ErrorAction SilentlyContinue)) {
            Import-Module sqlps -DisableNameChecking -ErrorAction SilentlyContinue
        }

        Try {
            $regKey        = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server' -ErrorAction Stop
            $instanceNames = @($regKey.InstalledInstances)   # wrap in @() — single instance returns a plain string
            if ($instanceNames.Count -eq 0 -or ($instanceNames.Count -eq 1 -and [string]::IsNullOrEmpty($instanceNames[0]))) {
                Write-Host "No SQL Server instances found in registry on: $vmName"
                return @()
            }
            Write-Host "Found $($instanceNames.Count) instance(s) on ${vmName}: $($instanceNames -join ', ')"
        }
        Catch {
            Write-Host "SQL Server registry key not found on: $vmName — $_"
            return @()
        }

        Try {
            $null = Get-Command Invoke-Sqlcmd -ErrorAction Stop
            Write-Host "Invoke-Sqlcmd is available."
        }
        Catch {
            Write-Host "ERROR: Invoke-Sqlcmd not found. Run: Install-Module SqlServer -Force -AllowClobber"
            return @()
        }

        $jobResults = @()

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
    MAX(CASE WHEN name = 'remote admin connections' THEN value_in_use END) AS [Remote Admin Connections],
    MAX(CASE WHEN name = 'remote access' THEN value_in_use END) AS [Remote Access],
    MAX(CASE WHEN name = 'xp_cmdshell' THEN value_in_use END) AS [xp_cmdshell]
FROM sys.configurations
WHERE name IN (
    'min server memory (MB)', 'max server memory (MB)', 'optimize for ad hoc workloads',
    'remote admin connections', 'remote access', 'xp_cmdshell'
) OR name LIKE '%backup compression%'
"@

$compatQuery = @"
SELECT
    d.name AS [Database Name],
    d.compatibility_level AS [Compatibility Level],
    d.is_auto_update_stats_on AS [Auto Update Stats],
    d.is_auto_create_stats_on AS [Auto Create Stats],
    d.page_verify_option_desc AS [Page Verify],
    d.is_auto_shrink_on AS [Auto Shrink]
FROM sys.databases d
WHERE d.state_desc = 'ONLINE' AND d.name NOT IN ('master', 'tempdb', 'model', 'msdb')
"@

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

$traceflagquery = @"
CREATE TABLE #TraceStatus (TraceFlag INT, Status BIT, Global BIT, Session BIT);
INSERT INTO #TraceStatus (TraceFlag, Status, Global, Session) EXEC ('DBCC TRACESTATUS');
SELECT TraceFlag FROM #TraceStatus;
DROP TABLE #TraceStatus;
GO
"@

$checkDBQuery = @"
DECLARE @dbname NVARCHAR(256), @sql NVARCHAR(MAX);
CREATE TABLE #CheckDBInfo (DatabaseName NVARCHAR(256), LastCheckDB DATETIME NULL);
DECLARE db_cursor CURSOR FOR
    SELECT name FROM sys.databases WHERE state_desc = 'ONLINE' AND name NOT IN ('tempdb');
OPEN db_cursor; FETCH NEXT FROM db_cursor INTO @dbname;
WHILE @@FETCH_STATUS = 0
BEGIN
    CREATE TABLE #DBInfo (ParentObject NVARCHAR(255), Object NVARCHAR(255), Field NVARCHAR(255), Value NVARCHAR(255));
    SET @sql = 'DBCC DBINFO (' + QUOTENAME(@dbname) + ') WITH TABLERESULTS;';
    INSERT INTO #DBInfo EXEC sp_executesql @sql;
    INSERT INTO #CheckDBInfo (DatabaseName, LastCheckDB)
    SELECT @dbname, MAX(CASE WHEN Field = 'dbi_dbccLastKnownGood' THEN TRY_CAST(Value AS DATETIME) END) FROM #DBInfo;
    DROP TABLE #DBInfo;
    FETCH NEXT FROM db_cursor INTO @dbname;
END
CLOSE db_cursor; DEALLOCATE db_cursor;
SELECT DatabaseName, ISNULL(LastCheckDB, '1900-01-01') AS LastCheckDB FROM #CheckDBInfo ORDER BY LastCheckDB ASC;
DROP TABLE #CheckDBInfo;
"@

$vlfQuery = @"
CREATE TABLE #VLFInfo (DatabaseName SYSNAME, VLFCount INT);
DECLARE @dbName SYSNAME, @sql NVARCHAR(MAX);
DECLARE db_cursor CURSOR FOR
    SELECT name FROM sys.databases
    WHERE state_desc = 'ONLINE' AND name NOT IN ('master','model','msdb','tempdb')
      AND DATABASEPROPERTYEX(name, 'Updateability') = 'READ_WRITE';
OPEN db_cursor; FETCH NEXT FROM db_cursor INTO @dbName;
WHILE @@FETCH_STATUS = 0
BEGIN
    SET @sql = N'USE ' + QUOTENAME(@dbName) + N'; INSERT INTO #VLFInfo SELECT ''' + @dbName + N''', COUNT(*) FROM sys.dm_db_log_info(DB_ID());';
    EXEC sp_executesql @sql;
    FETCH NEXT FROM db_cursor INTO @dbName;
END
CLOSE db_cursor; DEALLOCATE db_cursor;
SELECT DatabaseName, VLFCount FROM #VLFInfo WHERE VLFCount > 1000 ORDER BY VLFCount DESC;
DROP TABLE #VLFInfo;
"@

$BkpQuery = @"
WITH BackupData AS (
    SELECT database_name,
        MAX(CASE WHEN type = 'D' THEN backup_finish_date ELSE NULL END) AS LastFullBackup,
        MAX(CASE WHEN type = 'L' THEN backup_finish_date ELSE NULL END) AS LastLogBackup
    FROM msdb.dbo.backupset GROUP BY database_name
)
SELECT d.name AS DatabaseName, d.recovery_model_desc AS RecoveryModel,
    ISNULL(bd.LastFullBackup, '1900-01-01 00:00:00.000') AS LastFullBackup,
    ISNULL(bd.LastLogBackup,  '1900-01-01 00:00:00.000') AS LastLogBackup
FROM sys.databases d
LEFT JOIN BackupData bd ON d.name = bd.database_name
WHERE d.name NOT IN ('tempdb', 'model') ORDER BY d.name;
"@

$maxdopquery = @"
DECLARE @sqlmajorver INT, @numa INT, @affined_cpus INT, @cpucount INT, @recommended_maxdop INT, @current_maxdop INT;
SELECT @sqlmajorver = CONVERT(int,(@@microsoftversion/0x1000000)&0xff);
SELECT @cpucount    = COUNT(cpu_id) FROM sys.dm_os_schedulers WHERE scheduler_id < 255 AND parent_node_id < 64;
SELECT @numa        = COUNT(DISTINCT parent_node_id) FROM sys.dm_os_schedulers WHERE scheduler_id < 255 AND parent_node_id < 64;
SELECT @affined_cpus= COUNT(cpu_id) FROM sys.dm_os_schedulers WHERE is_online=1 AND scheduler_id<255 AND parent_node_id<64;
SELECT @recommended_maxdop =
    CASE
        WHEN @numa=1 AND @affined_cpus<=8 THEN @affined_cpus
        WHEN @numa=1 AND @affined_cpus>8  THEN 8
        WHEN @sqlmajorver>=13 AND @numa>1 AND CEILING(@cpucount*1.00/@numa)<=15 THEN CEILING(@cpucount*1.00/@numa)
        WHEN @sqlmajorver>=13 AND @numa>1 AND CEILING(@cpucount*1.00/@numa)>15  THEN CASE WHEN CEILING(@cpucount*1.00/@numa/2)>16 THEN 16 ELSE CEILING(@cpucount*1.00/@numa/2) END
        WHEN @sqlmajorver<13  AND @numa>1 AND CEILING(@cpucount*1.00/@numa)<8  THEN CEILING(@cpucount*1.00/@numa)
        WHEN @sqlmajorver<13  AND @numa>1 AND CEILING(@cpucount*1.00/@numa)>=8 THEN 8
        ELSE 0
    END;
SELECT @current_maxdop = CONVERT(INT,value) FROM sys.configurations WHERE name='max degree of parallelism';
SELECT @recommended_maxdop AS [Recommended_MaxDOP], @current_maxdop AS [Current_MaxDOP];
"@

$tempDBFileSizeQuery = @"
SELECT name AS FileName, type_desc AS FileType, size*8/1024 AS SizeMB, physical_name AS PhysicalPath
FROM sys.master_files WHERE database_id = DB_ID('tempdb');
"@

$QueryStoreQuery = @"
DECLARE @major_version INT = CAST(LEFT(CAST(SERVERPROPERTY('ProductVersion') AS VARCHAR(20)),
    CHARINDEX('.',CAST(SERVERPROPERTY('ProductVersion') AS VARCHAR(20)))-1) AS INT);
IF @major_version >= 13
BEGIN
    IF OBJECT_ID('tempdb..#QueryStoreStatus') IS NOT NULL DROP TABLE #QueryStoreStatus;
    CREATE TABLE #QueryStoreStatus (database_name SYSNAME, query_store_status NVARCHAR(60), query_capture_mode INT, query_capture_mode_desc NVARCHAR(20), sql_major_version INT);
    DECLARE @sql NVARCHAR(MAX)=N'';
    SELECT @sql += 'IF DB_ID(N''' + name + ''') IS NOT NULL BEGIN INSERT INTO #QueryStoreStatus SELECT ''' + name + ''', actual_state_desc, query_capture_mode, CASE query_capture_mode WHEN 0 THEN ''OFF'' WHEN 1 THEN ''ALL'' WHEN 2 THEN ''AUTO'' WHEN 3 THEN ''CUSTOM'' ELSE ''UNKNOWN'' END, ' + CAST(@major_version AS NVARCHAR) + ' FROM [' + name + '].sys.database_query_store_options; END;'
    FROM sys.databases WHERE database_id NOT IN(1,2,3) AND state_desc='ONLINE' AND DATABASEPROPERTYEX(name,'Updateability')='READ_WRITE';
    EXEC sp_executesql @sql;
    SELECT * FROM #QueryStoreStatus ORDER BY database_name;
END
ELSE SELECT 'There is no query store feature available on this SQL version' AS message;
"@

$ADRQuery = @"
DECLARE @major_version INT = CAST(LEFT(CAST(SERVERPROPERTY('ProductVersion') AS VARCHAR(20)), CHARINDEX('.',CAST(SERVERPROPERTY('ProductVersion') AS VARCHAR(20)))-1) AS INT);
IF @major_version >= 15
    SELECT name AS database_name, CASE is_accelerated_database_recovery_on WHEN 1 THEN 'ON' WHEN 0 THEN 'OFF' ELSE 'UNKNOWN' END AS status FROM sys.databases WHERE database_id > 4 ORDER BY name;
ELSE
    SELECT 'ADR does not exist in SQL Server versions earlier than 15 (SQL Server 2019).' AS message;
"@

$loginUserTestQuery = @"
DECLARE @dbname sysname;
CREATE TABLE #TestPrincipals (PrincipalType NVARCHAR(10), PrincipalName sysname, DatabaseName sysname NULL);
INSERT INTO #TestPrincipals SELECT 'Login', sp.name, NULL FROM sys.server_principals AS sp WHERE sp.name LIKE '%test%' AND sp.type IN ('S','U','G') AND sp.name NOT LIKE '##%';
DECLARE db_cursor CURSOR FAST_FORWARD FOR SELECT name FROM sys.databases WHERE state_desc='ONLINE' AND database_id>4 AND DATABASEPROPERTYEX(name,'Updateability')='READ_WRITE';
OPEN db_cursor; FETCH NEXT FROM db_cursor INTO @dbname;
WHILE @@FETCH_STATUS=0
BEGIN
    DECLARE @sql NVARCHAR(MAX);
    SET @sql=N'INSERT INTO #TestPrincipals SELECT ''User'',dp.name,N''' + @dbname + N''' FROM ' + QUOTENAME(@dbname) + N'.sys.database_principals AS dp WHERE dp.name LIKE ''%test%'' AND dp.type IN (''S'',''U'',''G'') AND dp.principal_id>4;';
    EXEC(@sql); FETCH NEXT FROM db_cursor INTO @dbname;
END
CLOSE db_cursor; DEALLOCATE db_cursor;
SELECT PrincipalType, PrincipalName, DatabaseName FROM #TestPrincipals;
DROP TABLE #TestPrincipals;
"@

$detectsqlaccounts = @"
SELECT servicename, service_account FROM sys.dm_server_services
WHERE servicename LIKE 'SQL Server%' OR servicename LIKE 'SQL Server Agent%';
"@

$alwaysonquery = @"
SELECT ISNULL((SELECT STUFF((SELECT ' | ' + ag.name + ': ' + STUFF((SELECT ', ' + ar.replica_server_name FROM sys.availability_replicas ar WHERE ar.group_id=ag.group_id ORDER BY ar.replica_server_name FOR XML PATH(''),TYPE).value('.','NVARCHAR(MAX)'),1,2,'') FROM sys.availability_groups ag ORDER BY ag.name FOR XML PATH(''),TYPE).value('.','NVARCHAR(MAX)'),1,3,'')),'NO') AS AlwaysOnDetails;
"@

        foreach ($instanceName in $instanceNames) {

            # Read port from registry
            try {
                $basePath = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server"; $port = "1433"
                $instanceID = Get-ChildItem -Path $basePath -ErrorAction Stop |
                              Where-Object { $_.PSChildName -match "^MSSQL.*\.$instanceName$" } |
                              Select-Object -ExpandProperty PSChildName -ErrorAction Stop
                if ($instanceID) {
                    $tcpKeyPath = "$basePath\$instanceID\MSSQLServer\SuperSocketNetLib\Tcp\IPAll"
                    try { $port = (Get-ItemProperty -Path $tcpKeyPath -Name TcpPort -ErrorAction Stop).TcpPort; if (-not $port) { $port="1433" } } catch { $port="1433" }
                }
            } catch { $port = "1433" }

            $effectiveName = if ($baseServerName -eq "localhost") { "localhost" } else { $vmName }
            $sqlInstance   = if ($instanceName -eq "MSSQLSERVER") { $effectiveName } else { "$effectiveName\$instanceName" }
            if ($port -ne "1433") { $sqlInstance = "$sqlInstance,$port" }
            Write-Host "Attempting connection to SQL instance: $sqlInstance"

            Try {
                $invk       = Get-Command Invoke-Sqlcmd -ErrorAction Stop
                $hasTSC     = $invk.Parameters.ContainsKey('TrustServerCertificate')
                $hasEncrypt = $invk.Parameters.ContainsKey('Encrypt')

                $RunQuery = {
                    param([string]$ServerInstance, [string]$Query)
                    $p = @{ ServerInstance=$ServerInstance; Query=$Query; QueryTimeout=65535; ErrorAction='Stop' }
                    if ($hasEncrypt) { $p['Encrypt']                = 'Optional' }
                    if ($hasTSC)     { $p['TrustServerCertificate'] = $true }
                    return Invoke-Sqlcmd @p
                }

                $currentQuery="Main Query";                $mainResult           = & $RunQuery $sqlInstance $mainQuery
                $currentQuery="Compatibility Level Query"; $compatResult         = & $RunQuery $sqlInstance $compatQuery
                $currentQuery="AutoGrow Query";            $autoGrowResult       = & $RunQuery $sqlInstance $autoGrowQuery
                $currentQuery="Trace Flag Query";          $enabledFlags         = & $RunQuery $sqlInstance $traceflagquery
                $currentQuery="CheckDB Query";             $checkDBResult        = & $RunQuery $sqlInstance $checkDBQuery
                $currentQuery="VLFs Query";                $vlfResult            = & $RunQuery $sqlInstance $vlfQuery
                $currentQuery="Backup Query";              $backupResult         = & $RunQuery $sqlInstance $BkpQuery
                $currentQuery="MaxDop Query";              $maxdopresult         = & $RunQuery $sqlInstance $maxdopquery
                $currentQuery="TempDB Query";              $tempDBFileSizeResult = & $RunQuery $sqlInstance $tempDBFileSizeQuery
                $currentQuery="Query Store Query";         $QueryStoreResults    = & $RunQuery $sqlInstance $QueryStoreQuery
                $currentQuery="ADR Query";                 $ADRQueryResults      = & $RunQuery $sqlInstance $ADRQuery
                $currentQuery="Login/User Test Query";     $loginUserTestResult  = & $RunQuery $sqlInstance $loginUserTestQuery
                $currentQuery="Detect SQL service accts";  $sqlaccountsresult    = & $RunQuery $sqlInstance $detectsqlaccounts
                $currentQuery="Check Always On";           $alwaysonresult       = & $RunQuery $sqlInstance $alwaysonquery
            }
            Catch {
                Write-Host "Error executing '$currentQuery' on '$sqlInstance': $($_.Exception.Message)"
                $mainResult=$compatResult=$autoGrowResult=$enabledFlags=$checkDBResult=$vlfResult=@()
                $backupResult=$maxdopresult=$tempDBFileSizeResult=$QueryStoreResults=$ADRQueryResults=@()
                $loginUserTestResult=$sqlaccountsresult=$alwaysonresult=@()
            }

            # MaxDOP
            $maxDopStatus='REVIEW'; $recommendedMaxDop=$null; $currentMaxDop=$null
            foreach ($r in $maxdopresult) {
                $recommendedMaxDop=$r.Recommended_MaxDOP; $currentMaxDop=$r.Current_MaxDOP
                $maxDopStatus = if ($currentMaxDop -and $recommendedMaxDop -and $currentMaxDop -ne 0 -and $currentMaxDop -eq $recommendedMaxDop) {'OK'} else {'REVIEW'}
            }

            # Backups
            $databasesWithoutFullBackup=@(); $databasesWithoutLogBackup=@()
            foreach ($database in $backupResult) {
                $lastFullBackup = try {[datetime]$database.LastFullBackup} catch {$null}
                $lastLogBackup  = try {[datetime]$database.LastLogBackup}  catch {$null}
                $fullBackupStatus = if ($lastFullBackup -eq $null -or $lastFullBackup -lt (Get-Date).AddDays(-7)) {'REVIEW'} else {'OK'}
                $logBackupStatus  = if (($database.RecoveryModel -in @('FULL','BULK_LOGGED')) -and ($lastLogBackup -eq $null -or $lastLogBackup -lt (Get-Date).AddHours(-24))) {'REVIEW'} else {'OK'}
                if ($fullBackupStatus -eq 'REVIEW') {$databasesWithoutFullBackup += $database.DatabaseName}
                if ($logBackupStatus  -eq 'REVIEW') {$databasesWithoutLogBackup  += $database.DatabaseName}
                $database | Add-Member -MemberType NoteProperty -Name "FullBackupStatus" -Value $fullBackupStatus -Force
                $database | Add-Member -MemberType NoteProperty -Name "LogBackupStatus"  -Value $logBackupStatus  -Force
            }
            $fullBackupStatusMessage = if ($databasesWithoutFullBackup.Count -gt 0) {"DBs: " + ($databasesWithoutFullBackup -join ', ')} else {"All databases have recent full backups."}
            $nonSimpleDatabases = $backupResult | Where-Object {$_.RecoveryModel -ne 'SIMPLE'}
            $logBackupStatusMessage = if ($databasesWithoutLogBackup.Count -eq 0 -and $nonSimpleDatabases.Count -eq 0) {"Not applicable because the databases are using the SIMPLE recovery model."} elseif ($databasesWithoutLogBackup.Count -gt 0) {"DBs: " + ($databasesWithoutLogBackup -join ', ')} else {"All databases have recent log backups."}

            # VLFs
            $databasesWithHighVLFs = @($vlfResult | Where-Object {$_.VLFCount -gt 1000} | Select-Object -ExpandProperty DatabaseName)
            $vlfStatus  = if ($databasesWithHighVLFs.Count -gt 0) {"REVIEW"} else {"OK"}
            $vlfMessage = if ($databasesWithHighVLFs.Count -eq 0)  {"All databases are OK"} else {$databasesWithHighVLFs -join ', '}

            # CHECKDB
            $sevenDaysAgo = (Get-Date).AddDays(-7)
            $missingCheckDB = @($checkDBResult | Where-Object {$_.LastCheckDB -lt $sevenDaysAgo -or $_.LastCheckDB -eq $null} | Select-Object -ExpandProperty DatabaseName)
            $checkDBStatus  = if ($missingCheckDB.Count -gt 0) {"REVIEW"} else {"OK"}
            $missingCheckDBMessage = if ($missingCheckDB.Count -eq 0) {"All databases are OK"} else {$missingCheckDB -join ', '}

            # Login/User test
            $testLogins         = @($loginUserTestResult | Where-Object {$_.PrincipalType -eq 'Login'} | Select-Object -ExpandProperty PrincipalName -ErrorAction SilentlyContinue)
            $testUsersFormatted = @($loginUserTestResult | Where-Object {$_.PrincipalType -eq 'User'}  | ForEach-Object {"$($_.DatabaseName)_users: $($_.PrincipalName)"})
            $loginUserTestStatus = if ($testLogins.Count -gt 0 -or $testUsersFormatted.Count -gt 0) {"REVIEW"} else {"OK"}
            $loginUserTestValidation = ""
            if ($testLogins.Count -gt 0)        {$loginUserTestValidation += "Logins: " + ($testLogins -join ', ')}
            if ($testUsersFormatted.Count -gt 0) {if ($loginUserTestValidation) {$loginUserTestValidation += " | "}; $loginUserTestValidation += "Users: " + ($testUsersFormatted -join '; ')}
            if (-not $loginUserTestValidation)    {$loginUserTestValidation = "No logins or users with 'test' found"}

            $mainResult = @($mainResult)   # ensure array — single-row result is not an array
            Write-Host "Main query returned $($mainResult.Count) row(s)."
            if ($mainResult.Count -gt 0) {
                foreach ($result in $mainResult) {
                    $serverNameWithInstance = $result.'Server Name'
                    $serverName       = $serverNameWithInstance -replace '\\.*$',''
                    $instanceNameOnly = $serverNameWithInstance -replace '^[^\\]*\\',''
                    if ($instanceNameOnly -eq $serverName) {$instanceNameOnly = 'DEFAULT'}

                    $resultObject = [PSCustomObject]@{
                        "Server Name"                                 = $serverName
                        "SQL Instance Name"                           = $instanceNameOnly
                        "SQL Server Version"                          = switch -Wildcard ($result.'SQL Build Number') {"11*"{"SQL Server 2012"}"12*"{"SQL Server 2014"}"13*"{"SQL Server 2016"}"14*"{"SQL Server 2017"}"15*"{"SQL Server 2019"}"16*"{"SQL Server 2022"}"17*"{"SQL Server 2025"}default{"Unknown Version"}}
                        "SQL Build Number"                            = $result.'SQL Build Number'
                        "SQL Edition"                                 = $result.'SQL Edition'
                        "Total Server Memory (MB)"                    = $result.'Total Server Memory (MB)'
                        "Current Min Server Memory (MB)"              = $result.'Min Server Memory (MB)'
                        "Current Max Server Memory (MB)"              = $result.'Max Server Memory (MB)'
                        "Total Visible Processors"                    = $result.'Total Visible Processors'
                        "Optimize for Ad Hoc Workloads"               = $result.'Optimize for Ad Hoc Workloads'
                        "Backup Compression Default"                  = if ([string]::IsNullOrEmpty($result.'Backup Compression Default')) {"Backup compression not available, check SQL server edition"} else {$result.'Backup Compression Default'}
                        "Remote Admin Connections"                    = $result.'Remote Admin Connections'
                        "Remote Access"                               = $result.'Remote Access'
                        "xp_cmdshell"                                 = $result.'xp_cmdshell'
                        "TempDB Data Files"                           = $result.'TempDB Data Files'
                        "DBs with missing CHECKDB in the last 7 days" = $missingCheckDBMessage
                        "CHECKDB Status"                              = $checkDBStatus
                        "Login or User Test Status"                   = $loginUserTestStatus
                        "Login and Users validation"                  = $loginUserTestValidation
                    }

                    # Trace flags
                    $traceFlagsByVersion = @{'11'=@(4199,1118);'12'=@(4199,1118);'13'=@(4199,7745);'14'=@(4199,7745,12310);'15'=@(4199,7745,12310);'16'=@(4199,7745,12656,12618)}
                    $majorVersion    = ($result.'SQL Build Number' -split '\.')[0]
                    $majorVersionInt = [int]$majorVersion
                    $enabledFlagNames = if ($enabledFlags -and $enabledFlags.TraceFlag) {$enabledFlags.TraceFlag | ForEach-Object {[int]$_}} else {@()}
                    $traceFlagList   = if ($majorVersionInt -ge 16) {$traceFlagsByVersion['16']} elseif ($traceFlagsByVersion.ContainsKey($majorVersion)) {$traceFlagsByVersion[$majorVersion]} else {$null}
                    $traceFlagStatus = if (-not $traceFlagList) {"REVIEW"} elseif (($traceFlagList | Where-Object {[int]$_ -notin $enabledFlagNames}).Count -gt 0) {"REVIEW"} else {"OK"}
                    $traceFlagListString = if ($enabledFlagNames.Count -gt 0) {$enabledFlagNames -join ", "} else {"No trace flags enabled"}

                    # Memory
                    $recommendedMaxMemory = [math]::Round($result.'Total Server Memory (MB)' * 0.75, 0)
                    $memoryStatus = if ($result.'Max Server Memory (MB)' -eq 2147483647 -or $result.'Max Server Memory (MB)' -ge $result.'Total Server Memory (MB)') {"REVIEW"} elseif ($result.'Min Server Memory (MB)' -eq 1024 -and $result.'Max Server Memory (MB)' -ge $recommendedMaxMemory) {"OK"} else {"REVIEW"}

                    # Config
                    $configStatus = if ($result.'Optimize for Ad Hoc Workloads' -eq 0 -or $result.'Remote Admin Connections' -eq 0 -or (![string]::IsNullOrEmpty($result.'Backup Compression Default') -and $result.'Backup Compression Default' -lt 1) -or $result.'Remote Access' -eq 1 -or $result.'xp_cmdshell' -eq 1) {"REVIEW"} else {"OK"}

                    # Compatibility level
                    $nativeCompatibilityLevel = switch ($majorVersion) {'11'{110}'12'{120}'13'{130}'14'{140}'15'{150}'16'{160}'17'{170}default{0}}
                    $compatLevels = $compatResult | Where-Object {$_.'Compatibility Level' -ne $nativeCompatibilityLevel} | ForEach-Object {"$($_.'Database Name') (Level $($_.'Compatibility Level'))"}
                    $compatLevelsMessage      = if ($compatLevels.Count -eq 0) {"All databases are in native compatibility level!"} else {$compatLevels -join ', '}
                    $compatibilityLevelStatus = if ($compatLevels.Count -gt 0) {"REVIEW"} else {"OK"}

                    # Database options
                    $divergentDatabases  = @()
                    $divergentDatabases += $compatResult | Where-Object {$_.'Auto Update Stats' -eq 0} | ForEach-Object {"$($_.'Database Name') (Auto Update Stats OFF)"}
                    $divergentDatabases += $compatResult | Where-Object {$_.'Auto Create Stats' -eq 0} | ForEach-Object {"$($_.'Database Name') (Auto Create Stats OFF)"}
                    $divergentDatabases += $compatResult | Where-Object {$_.'Auto Shrink'       -eq 1} | ForEach-Object {"$($_.'Database Name') (Auto Shrink ON)"}
                    $divergentDatabases += $compatResult | Where-Object {$_.'Page Verify'  -ne "CHECKSUM"} | ForEach-Object {"$($_.'Database Name') (Page Verify NOT CHECKSUM)"}
                    $divergentDatabasesMessage = if ($divergentDatabases.Count -eq 0) {"All settings are OK!"} else {$divergentDatabases -join ', '}
                    $databaseOptionsStatus     = if ($divergentDatabases.Count -gt 0) {"REVIEW"} else {"OK"}

                    # AutoGrow
                    $unlimitedAutoGrowFiles      = @($autoGrowResult | Where-Object {$_.'Max Size' -eq -1}                           | Select-Object -ExpandProperty 'File Name')
                    $percentageAutoGrowFiles     = @($autoGrowResult | Where-Object {$_.'Growth Type' -eq 'Percentage-based growth'} | Select-Object -ExpandProperty 'File Name')
                    $largeIncrementAutoGrowFiles = @($autoGrowResult | Where-Object {$_.'AutoGrow Increment (MB)' -gt 1024}          | Select-Object -ExpandProperty 'File Name')
                    $unlimitedAutoGrowMessage     = if ($unlimitedAutoGrowFiles.Count    -eq 0) {"No files have unlimited AutoGrow."}          else {"Unlimited AutoGrow: "     + ($unlimitedAutoGrowFiles    -join ', ') + ","}
                    $percentageAutoGrowMessage    = if ($percentageAutoGrowFiles.Count   -eq 0) {"No files are using AutoGrow by percentage."}  else {"AutoGrow by percentage: " + ($percentageAutoGrowFiles   -join ', ') + ","}
                    $largeIncrementAutoGrowMessage= if ($largeIncrementAutoGrowFiles.Count-eq 0) {"No files have large increments for AutoGrow."} else {"Large Increment AutoGrow: " + ($largeIncrementAutoGrowFiles -join ', ') + ","}
                    $autoGrowthStatus = if (($unlimitedAutoGrowFiles.Count + $percentageAutoGrowFiles.Count + $largeIncrementAutoGrowFiles.Count) -gt 0) {"REVIEW"} else {"OK"}

                    # Query Store
                    if ($QueryStoreResults -and $QueryStoreResults.Count -gt 0) {
                        if ($QueryStoreResults[0].PSObject.Properties.Name -contains "message" -and $QueryStoreResults[0].message -like "*no query store*") {
                            $QueryStoreStatus="OK"; $QueryStoreDatabases="There is no query store feature available on this SQL version"
                        } else {
                            $QueryStoreEnabled   = $QueryStoreResults | Where-Object {$_.query_store_status -ne $null}
                            $QueryStoreDatabases = ($QueryStoreEnabled | ForEach-Object {"$($_.database_name): $($_.query_store_status) / $($_.query_capture_mode_desc)"}) -join ", "
                            $QueryStoreStatus    = if ($QueryStoreEnabled | Where-Object {$_.query_store_status -eq "OFF"}) {"REVIEW"} else {"OK"}
                        }
                    } else {$QueryStoreStatus="OK"; $QueryStoreDatabases="This SQL version doesn't have Query Store available"}

                    # ADR
                    $ADRRows = @($ADRQueryResults)
                    if ($ADRRows.Count -eq 0 -or $null -eq $ADRRows[0]) {
                        $ADRStatus="REVIEW"; $ADRDetails="ADR query returned no rows."
                    } elseif ($ADRRows[0].PSObject.Properties.Match('message').Count -gt 0) {
                        $ADRStatus="OK"; $ADRDetails=$ADRRows[0].message
                    } else {
                        $ADRDetails = ($ADRRows | ForEach-Object {"$($_.database_name): $($_.status)"}) -join ", "
                        $ADRStatus  = if (@($ADRRows | Where-Object {$_.status -in @('OFF','UNKNOWN')}).Count -gt 0) {"REVIEW"} else {"OK"}
                    }

                    # TempDB
                    $totalProcessors     = [int]$result.'Total Visible Processors'
                    $tempDBDataFiles     = @($tempDBFileSizeResult | Where-Object {$_.FileType -eq "ROWS"})
                    $totalTempDBDataFiles= $tempDBDataFiles.Count
                    $tempDBFileSizes     = $tempDBDataFiles | Select-Object -ExpandProperty SizeMB | Sort-Object -Unique
                    $allFilesSameSize    = ($tempDBFileSizes.Count -eq 1)
                    $tempDBMultipleOf4   = ($totalTempDBDataFiles % 4 -eq 0)
                    $recMin = if ($totalProcessors -ge 8) {4} elseif ($totalProcessors -ge 4) {2} else {1}
                    $recMax = if ($totalProcessors -ge 4) {8} else {1}
                    $singleFileOkay = ($majorVersionInt -ge 16 -and $totalTempDBDataFiles -eq 1)
                    $tempDBStatus   = if ((-not $singleFileOkay -and ($totalTempDBDataFiles -lt $recMin -or $totalTempDBDataFiles -gt $recMax)) -or -not $allFilesSameSize -or (-not $tempDBMultipleOf4 -and $totalTempDBDataFiles -gt 1)) {"REVIEW"} else {"OK"}
                    $tempDBUniformSize = if ($totalTempDBDataFiles -eq 1) {"Nothing to compare because TempDB has only one data file"} elseif (-not $allFilesSameSize) {($tempDBDataFiles | ForEach-Object {"$($_.FileName): $($_.SizeMB) MB"}) -join ", "} else {"All data files have the same size"}

                    # SQL service accounts
                    $SQLServiceAccountsStatus="OK"; $SQLServiceAccountsDetails=""
                    foreach ($row in $sqlaccountsresult) {
                        if ($row.service_account -in @('LocalSystem','LocalService','NT AUTHORITY\NETWORK SERVICE') -or $row.service_account -like 'NT Service*') {$SQLServiceAccountsStatus='REVIEW'}
                        $SQLServiceAccountsDetails += "$($row.servicename) Account: $($row.service_account) | "
                    }

                    # Always On
                    $alwaysOnDetails = if ($alwaysonresult -and $alwaysonresult.AlwaysOnDetails) {$alwaysonresult.AlwaysOnDetails} else {"NO"}
                    $alwaysOnStatus  = if ($alwaysOnDetails -eq "NO") {"REVIEW"} else {"OK"}

                    $resultObject | Add-Member -MemberType NoteProperty -Name "Memory Status"                                -Value $memoryStatus
                    $resultObject | Add-Member -MemberType NoteProperty -Name "Config Status"                                -Value $configStatus
                    $resultObject | Add-Member -MemberType NoteProperty -Name "MaxDop Status"                                -Value $maxDopStatus
                    $resultObject | Add-Member -MemberType NoteProperty -Name "Databases out of native compatibility"        -Value $compatLevelsMessage
                    $resultObject | Add-Member -MemberType NoteProperty -Name "Unlimited AutoGrow"                           -Value $unlimitedAutoGrowMessage
                    $resultObject | Add-Member -MemberType NoteProperty -Name "AutoGrow by Percentage"                       -Value $percentageAutoGrowMessage
                    $resultObject | Add-Member -MemberType NoteProperty -Name "Large Increment AutoGrow"                     -Value $largeIncrementAutoGrowMessage
                    $resultObject | Add-Member -MemberType NoteProperty -Name "Auto Growth Status"                           -Value $autoGrowthStatus
                    $resultObject | Add-Member -MemberType NoteProperty -Name "Compatibility Level Status"                   -Value $compatibilityLevelStatus
                    $resultObject | Add-Member -MemberType NoteProperty -Name "Database Options Status"                      -Value $databaseOptionsStatus
                    $resultObject | Add-Member -MemberType NoteProperty -Name "Database Options Divergence"                  -Value $divergentDatabasesMessage
                    $resultObject | Add-Member -MemberType NoteProperty -Name "Trace Flag List"                              -Value $traceFlagListString
                    $resultObject | Add-Member -MemberType NoteProperty -Name "Trace Flag Status"                            -Value $traceFlagStatus
                    $resultObject | Add-Member -MemberType NoteProperty -Name "DBs with too many VLFs"                       -Value $vlfMessage
                    $resultObject | Add-Member -MemberType NoteProperty -Name "VLF Status"                                   -Value $vlfStatus
                    $resultObject | Add-Member -MemberType NoteProperty -Name "Full Backup Status"                           -Value $fullBackupStatus
                    $resultObject | Add-Member -MemberType NoteProperty -Name "Log Backup Status"                            -Value $logBackupStatus
                    $resultObject | Add-Member -MemberType NoteProperty -Name "DBs missing Full Backup in the last 7 days"   -Value $fullBackupStatusMessage
                    $resultObject | Add-Member -MemberType NoteProperty -Name "DBs missing Log Backup (With full rec model)" -Value $logBackupStatusMessage
                    $resultObject | Add-Member -MemberType NoteProperty -Name "Recommended Min Server Memory (MB)"           -Value 1024
                    $resultObject | Add-Member -MemberType NoteProperty -Name "Recommended Max Server Memory (MB)"           -Value $recommendedMaxMemory
                    $resultObject | Add-Member -MemberType NoteProperty -Name "Current Max Dop"                              -Value $currentMaxDop
                    $resultObject | Add-Member -MemberType NoteProperty -Name "Recommended Max Dop"                          -Value $recommendedMaxDop
                    $resultObject | Add-Member -MemberType NoteProperty -Name "Query Store Status"                           -Value $QueryStoreStatus
                    $resultObject | Add-Member -MemberType NoteProperty -Name "Query Store Details"                          -Value $QueryStoreDatabases
                    $resultObject | Add-Member -MemberType NoteProperty -Name "ADR Status"                                   -Value $ADRStatus
                    $resultObject | Add-Member -MemberType NoteProperty -Name "ADR Details"                                  -Value $ADRDetails
                    $resultObject | Add-Member -MemberType NoteProperty -Name "TempDB Status"                                -Value $tempDBStatus
                    $resultObject | Add-Member -MemberType NoteProperty -Name "TempDB Data Files Count"                      -Value $totalTempDBDataFiles
                    $resultObject | Add-Member -MemberType NoteProperty -Name "TempDB Data Files Size"                       -Value $tempDBUniformSize
                    $resultObject | Add-Member -MemberType NoteProperty -Name "SQL Service Accounts Status"                  -Value $SQLServiceAccountsStatus
                    $resultObject | Add-Member -MemberType NoteProperty -Name "SQL Service Accounts Details"                 -Value $SQLServiceAccountsDetails.TrimEnd(" | ")
                    $resultObject | Add-Member -MemberType NoteProperty -Name "Always On Status"                             -Value $alwaysOnStatus
                    $resultObject | Add-Member -MemberType NoteProperty -Name "Always On Details"                            -Value $alwaysOnDetails

                    $jobResults += $resultObject
                }
            }
        }
        return $jobResults
    }
    Catch {
        Write-Host "Error during SQL query execution: $_"
        return @()
    }
}

# ── Server loop ───────────────────────────────────────────────────────────────

$Jobs = @(); $Results = @()
$totalServers = $serverEntries.Count; $processedServers = 0

foreach ($entry in $serverEntries) {
    $processedServers++
    $progressPercent = [math]::Round(($processedServers / $totalServers) * 100)
    Write-Progress -Activity "Processing SQL Servers" -Status "Processing $processedServers of $totalServers ($progressPercent%)" -PercentComplete $progressPercent

    $BaseServerName = ($entry -split ',')[0]
    $VMs    = if ($UseFQDN -eq "yes" -and $BaseServerName -ne "localhost") {"$BaseServerName.$DomainName"} else {$BaseServerName}
    $VMName = ($BaseServerName -split '\.')[0]

    if (IsLocalMachine -Name $BaseServerName) {
        Write-Host "Local machine detected — running directly (no PSRemoting required)."
        Try {
            $localResult = & $CoreScriptBlock $VMName $BaseServerName
            if ($localResult) {$Results += $localResult}
        }
        Catch {
            Write-Host "Error running locally: $_"
        }
    }
    else {
        Try {
            $Session = New-PSSession -ComputerName $VMs -ErrorAction Stop
            Write-Host "Successfully connected to $VMs"
            $Job = Invoke-Command -Session $Session -ScriptBlock $CoreScriptBlock -ArgumentList $VMName, $BaseServerName -AsJob
            $Jobs += $Job
        }
        Catch {
            Write-Host "Failed to connect to $($VMs): $_"
        }
    }
}

Write-Progress -Activity "Processing SQL Servers" -Status "Completed" -PercentComplete 100 -Completed

if ($Jobs.Count -gt 0) {
    Write-Host "Waiting for remote jobs to finish..."
    Wait-Job -Job $Jobs
    $Jobs | ForEach-Object {
        $JobResult = Receive-Job -Job $_
        if ($JobResult) {$Results += $JobResult}
    }
    $Jobs | Remove-Job
}

if ($Results.Count -eq 0) {
    Write-Host "No results collected. Exiting..." -ForegroundColor Red
    Stop-Transcript; Exit 1
}

# ── Output ────────────────────────────────────────────────────────────────────

$columnOrder = @(
    "Server Name","SQL Instance Name","SQL Server Version","SQL Build Number","SQL Edition",
    "Always On Status","Always On Details","Memory Status","Config Status","MaxDop Status",
    "Query Store Status","ADR Status","TempDB Status","Auto Growth Status","Database Options Status",
    "Compatibility Level Status","Trace Flag Status","CHECKDB Status","VLF Status",
    "Full Backup Status","Log Backup Status","Login or User Test Status","SQL Service Accounts Status",
    "Total Server Memory (MB)","Current Min Server Memory (MB)","Recommended Min Server Memory (MB)",
    "Current Max Server Memory (MB)","Recommended Max Server Memory (MB)","Total Visible Processors",
    "Current Max Dop","Recommended Max Dop","Optimize for Ad Hoc Workloads","Backup Compression Default",
    "Remote Admin Connections","Remote Access","xp_cmdshell","Databases out of native compatibility",
    "Database Options Divergence","Unlimited AutoGrow","AutoGrow by Percentage","Large Increment AutoGrow",
    "Trace Flag List","DBs with missing CHECKDB in the last 7 days","DBs with too many VLFs",
    "DBs missing Full Backup in the last 7 days","DBs missing Log Backup (With full rec model)",
    "Login and Users validation","Query Store Details","ADR Details","SQL Service Accounts Details",
    "TempDB Data Files Count","TempDB Data Files Size"
)

$filteredResults = $Results | Select-Object -Property $columnOrder -ExcludeProperty PSComputerName,RunspaceId,PSShowComputerName
$uniqueResults   = $filteredResults | Group-Object -Property "Server Name","SQL Instance Name" | ForEach-Object {$_.Group | Select-Object -First 1}

$executionDate = Get-Date -Format "yyyyMMdd_HHmmss"
$filePath      = "C:\temp\SQL_Server_Best_Practices_Assessment_Results_$executionDate.csv"
$uniqueResults | Export-Csv -Path $filePath -NoTypeInformation

# ── Summary ───────────────────────────────────────────────────────────────────

$statusColumns = @(
    "Always On Status","Memory Status","Config Status","MaxDop Status","Query Store Status",
    "Auto Growth Status","Database Options Status","Compatibility Level Status","Trace Flag Status",
    "CHECKDB Status","VLF Status","Full Backup Status","Log Backup Status","ADR Status",
    "SQL Service Accounts Status","Login or User Test Status","TempDB Status"
)

$statusSummary = @{}
foreach ($col in $statusColumns) {$statusSummary[$col] = @{OK=0;REVIEW=0}}
foreach ($row in $uniqueResults) {
    foreach ($col in $statusColumns) {
        if     ($row.$col -eq "OK")     {$statusSummary[$col]["OK"]++}
        elseif ($row.$col -eq "REVIEW") {$statusSummary[$col]["REVIEW"]++}
    }
}

$maxLen = ($statusColumns | ForEach-Object {$_.Length} | Measure-Object -Maximum).Maximum
Write-Host "----------------------------------------------------------" -ForegroundColor Cyan
Write-Host " ===== SQL Server Best Practices Assessment Summary ===== " -ForegroundColor Cyan
Write-Host "----------------------------------------------------------" -ForegroundColor Cyan
Write-Host ("{0,-$maxLen} | {1,5} | {2,5}" -f "STATUS","OK","REVIEW") -ForegroundColor Yellow
Write-Host ("-" * ($maxLen + 20))
foreach ($col in $statusColumns) {
    Write-Host ("{0,-$maxLen} | {1,5} | {2,5}" -f $col, $statusSummary[$col]["OK"], $statusSummary[$col]["REVIEW"])
}
Write-Host ("=========================================================") -ForegroundColor Cyan
Write-Host "CSV results located at: $filePath" -ForegroundColor Green

Stop-Transcript