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

Important note:
After running the EF Core migrations, you must manually execute the SQL queries that create database views found in the SQL/init.sql file on your PostgreSQL database.
EF Core migrations do not automatically create views, so this step is necessary to ensure all views used by the application are available.

## Run
- Start the app (F5) → open `/Upload` to upload Excel (put sample in `wwwroot/data.xlsx` or use upload UI).
- After ETL, open `/` to see dashboard.
