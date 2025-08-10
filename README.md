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
EF Core migrations do not automatically create database views.
You must manually execute the SQL queries that create the required views on your PostgreSQL database.
These view definitions, including the CREATE OR REPLACE VIEW ticket_size statement, are provided in the `SQL/init.sql` file.

## Run
- Start the app (F5) → open `/Upload` to upload Excel (put sample in `wwwroot/data.xlsx` or use upload UI).
- After ETL, open `/` to see dashboard.
