# SQL Server Scripts

Hello! I will use this repo to insert some SQL and Powershell scripts. 

**03/21/2025:**

The first PowerShell script I uploaded is called S**QL Server Best Practices Assessment v20.ps1.** It can be used to perform a quick check against some best practices in your current SQL Server environment, some of the checks include:

1. Instance configs (Optimize For Ad Hoc workloads, Remote Dedicated Admin Connections and Backup Compression).
2. Min and Max Server Memory.
3. Max Degree of Parallelism.
4. Database Options (Auto Create/Auto Update/PAGE VERIFY)
5. VLFs checks.
6. Autogrow settings (Large Autogrows, Uniliminted Autogrow and Percent Autogrows)
7. Backup full execution in the last 7 days.
8. Backup log execution in the last 7 days.
9. CheckDB execution in the last 7 days.
10. Databases Compatibility Level.
11. TempDB Files checks (autogrow, size, number of files)

Feel free to download and modify this script according to your needs, if you use it publicly please reference the author =D.
