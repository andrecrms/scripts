# SQL Server Scripts

Hello! I will use this repo to insert some SQL and Powershell scripts. 

**03/21/2025:**

The first PowerShell script I uploaded is called **SQL Server Best Practices Assessment v20.ps1** (https://github.com/andrecrms/scripts/blob/PowerShell-Scripts/SQL%20Server%20Best%20Practices%20Assessment%20v20.ps1). It can be used to perform a quick check against some best practices in your current SQL Server environment, some of the checks include:

1. Instance configs (Optimize For Ad Hoc workloads, Remote Dedicated Admin Connections and Backup Compression, for this last one, in case of SQL EXPRESS EDITION, it will be disregarded).
2. Min and Max Server Memory. (Script mark as REVIEW in case Min server memory is <> 1024 MB and if Max Server Memory is not at least 75% of the total server memory)
3. Max Degree of Parallelism. (BPCheck script logic)
4. Database Options (Auto Create/Auto Update/PAGE VERIFY)
5. VLFs checks.
6. Autogrow settings (Large Autogrows, Unlimited Autogrow and Percent Autogrows)
7. Backup full execution in the last 7 days.
8. Backup log execution in the last 7 days.
9. CheckDB execution in the last 7 days.
10. Databases Compatibility Level.
11. TempDB Files checks (autogrow, size, number of files, for this last one, in case of SQL Server 2022 or higher, in case of 1 file, column will be marked as OK)

**Please test it before running it in your production environment and feel free to download and modify this script to suit your needs. If you use it publicly, please give credit to the author =D. Thanks!**