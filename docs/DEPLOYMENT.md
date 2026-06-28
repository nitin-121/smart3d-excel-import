# Deployment Guide — Smart3D Excel Property Import

## Prerequisites

- Smart3D V14 (or Forte 3D) installed
- .NET Framework 4.8 Developer Pack
- Visual Studio 2022 or Build Tools (MSBuild)
- NuGet package restore access

## Step 1: Configure Project Path

Edit `deploy.bat` and set your Smart3D installation path:

```batch
set SMART3D_HOME=C:\Program Files\Intergraph\SmartPlant 3D V14
```

Common paths:
- `C:\Program Files\Intergraph\SmartPlant 3D V14`
- `C:\Program Files\Hexagon\Smart 3D V14`
- `C:\Smart3D\V14`

## Step 2: Build the Project

```bash
cd src
dotnet restore
dotnet build -c Release
```

Or using MSBuild:
```bash
msbuild Smart3D.ExcelImport.csproj /p:Configuration=Release /p:Platform=x64
```

Expected output: `bin\Release\Smart3D.ExcelImport.dll`

## Step 3: Deploy to Smart3D

### Option A: Automated (Recommended)
```bash
deploy.bat
```

### Option B: Manual

1. **Copy DLLs to Smart3D bin folder:**
   ```
   copy Smart3D.ExcelImport.dll → %SMART3D_HOME%\bin\
   copy EPPlus.dll → %SMART3D_HOME%\bin\
   copy Serilog.dll → %SMART3D_HOME%\bin\
   copy Serilog.Sinks.File.dll → %SMART3D_HOME%\bin\
   ```

2. **Register COM component:**
   ```
   cd %SMART3D_HOME%\bin\
   RegAsm.exe Smart3D.ExcelImport.dll /codebase
   ```

3. **Add command to Smart3D:**
   Add this line to `%SMART3D_HOME%\CommandFiles\CustomCommands.txt`:
   ```
   ExcelImportCommand,Smart3D.ExcelImport.ExcelImportCommand,Smart3D.ExcelImport,Import Properties from Excel
   ```

## Step 4: Verify Installation

1. Start Smart3D V14
2. Open any model
3. Go to **Tools → Custom Commands**
4. Look for **"Import Properties from Excel"**
5. Click it — the file picker should appear

## Step 5: Test with Sample Data

1. Open `samples/SampleImport.csv` in Excel
2. Save as `.xlsx` format
3. Run the command in Smart3D
4. Select the `.xlsx` file
5. Check the summary report

## Troubleshooting

| Issue | Solution |
|-------|----------|
| "Command not found" | Verify CustomCommands.txt syntax and restart Smart3D |
| "DLL not found" | Check DLLs are in Smart3D bin folder |
| "COM registration failed" | Run RegAsm as Administrator |
| "Filtering service null" | Ensure a model is open before running command |
| "Type coercion failed" | Check Excel values match expected property types |
| Build fails on API refs | Set `SMART3D_HOME` env var to your install path |

## Uninstall

1. Delete DLLs from `%SMART3D_HOME%\bin\`
2. Remove command line from `CustomCommands.txt`
3. Unregister COM: `RegAsm.exe /unregister Smart3D.ExcelImport.dll`
4. Delete logs: `%APPDATA%\Smart3D\ExcelImport\`
