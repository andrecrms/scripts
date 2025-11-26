# SQL Server Scripts

Hello! I will use this repo to insert some SQL and Powershell scripts. 

**11/26/2025:**

The first PowerShell script I uploaded is called **SQL Server Best Practices Assessment v21.ps1** (https://github.com/andrecrms/scripts/blob/PowerShell-Scripts/SQL%20Server%20Best%20Practices%20Assessment%20v21.ps1). It can be used to perform a quick check against some best practices in your current SQL Server environment, some of the checks include:

1. Instance configs (Optimize For Ad Hoc workloads, Remote Dedicated Admin Connections and Backup Compression, for this last one, in case of SQL EXPRESS EDITION, it will be disregarded).
2. Min and Max Server Memory. (Status will be marked as REVIEW in case Min server memory is <> 1024 MB and if Max Server Memory is not at least 75% of the total server memory or not configured).
3. Max Degree of Parallelism. (BPCheck script logic, status will be marked as REVIEW if maxdop number is not the recommended one considering NUMA, affinity and visible processors for SQL Server)
4. Database Options (Auto Create/Auto Update/Auto Shrink/PAGE VERIFY, if one of the first two or both are turned off status column will be marked as REVIEW, if Auto Shrink is ON or PAGE VERIFY is not CHECKSUM, column also will be marked as REVIEW)
5. VLFs checks. (Any t-log file with more than 1000 VLFs, status will be marked as REVIEW)
6. Autogrow settings (For large Autogrows, unlimited Autogrow or Percent Autogrows, status will be marked as REVIEW)
7. Backup full execution in the last 7 days.
8. Backup log execution in the last 7 days.
9. CheckDB execution in the last 7 days.
10. Databases Compatibility Level. (Status column will be marked as REVIEW for any database not configured with the native compatibility level).
11. Query Store checks, the script will list all databases with Query Store enabled/disabled, the configured capture mode also will be listed, when Query Store is disabled for any user db or model, status column will be marked as REVIEW.
12. Accelerated Database Recovery check, the script will list all databases with ADR enabled/disabled, when ADR is disabled for any user db, status column will be marked as REVIEW. 
13. TempDB Files checks:
    * Autogrow is the same for all data files? If no, status column will be marked as REVIEW.
    * Size is the same for all data files? If no, status column will be marked as REVIEW.
    * Number of files until SQL Server 2019: if 4 processors, TempDB should have at least 2 files. If 8 processors, TempDB should have  at least 4 files. If 8 processors or more, TempDB should have at least 4 files but no more than 8. If one of the rules does not fit these definitions, the status column will be marked as REVIEW.
    * Number of files in case of SQL Server 2022 or higher: in case of 1 file, column will be marked as OK.
14. Trace Flags, check will vary according to SQL version
    * For SQL Server 2012 and 2014, if trace flags 1118 and 4199 are off, status column will be marked as REVIEW.
    * For SQL Server 2016, if trace flags 4199 and 7745 are off, status column will be marked as REVIEW.
    * For SQL Server 2017 and 2019, if trace flags 4199, 7745, 12310 are off, status column will be marked as REVIEW.
    * For SQL Server 2022 and 2025, if trace flags 4199, 7745, 12656 and 12618 are off, status column will be marked as REVIEW.
    * If none trace flags are enabled or if any from the list above are missing according with SQL Server version, status column also will be marked as REVIEW.
15. Logins and Users with word "test" checks, the script will check all logins and users with the word "test", if any exists the status column will be marked as REVIEW and all of them (logins/users) will be listed.
16. If SQL Server Service account or SQL Agent account are running using Local Service, Local System or NT SERVICE, if yes, then the status column will be marked as REVIEW.

**Tested on: SQL Server 2019 to 2025.**
    
**Please test it before running it in your production environment and feel free to download and modify this script to suit your needs. If you use it publicly, please give credit to the author =D. Thanks!**

**Troubleshooting**

In case of errors to connect/run script remotelly (access denied and others), here are some commands that you can try to use to make it work (credits to Ahmad Hassan from Microsoft - Egypt):

PS C:\Windows\system32> Enable-PSRemoting -Force<br>
PS C:\Windows\system32> Enable-PSRemoting -SkipNetworkProfileCheck -Force<br>
PS C:\Windows\system32> Set-NetFirewallRule -Name 'WINRM-HTTP-In-TCP' -RemoteAddress Any<br>
PS C:\temp> Set-ExecutionPolicy Unrestricted<br>
PS C:\temp> reg add HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System /v LocalAccountTokenFilterPolicy /t REG_DWORD /d 1 /f_<br>
