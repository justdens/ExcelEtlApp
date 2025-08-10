# ExcelETLApp (TRX & NPP)

## Prerequisites
- .NET 8.0 SDK
- Visual Studio 2022 / VS Code
- PostgreSQL

## Setup
1. Copy project files to a folder.
2. Open project in Visual Studio.
3. Edit `appsettings.json` connection string.
4. Install NuGet packages if needed (EPPlus 8.x, Npgsql.EntityFrameworkCore.PostgreSQL, Microsoft.EntityFrameworkCore.Design).

## Migrate database
Use Package Manager Console:

```
Add-Migration InitTRXandNPP
Update-Database
```

Or run `SQL/init.sql` directly.

## Run
- Start the app (F5) → open `/Upload` to upload Excel (put sample in `wwwroot/data.xlsx` or use upload UI).
- After ETL, open `/` to see dashboard.
