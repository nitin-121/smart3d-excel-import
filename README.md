# Smart3D V14 — Excel Property Import Custom Command

## Overview

Bulk-update Smart3D object properties (PipeRun, Pipeline, Equipment, Valve, etc.) from an Excel spreadsheet. Mirrors the "Import Data From Excel" feature found in Aveva E3D.

## Architecture

```
Excel (.xlsx) → EPPlus Parser → ImportRecord[] → ImportEngine → Smart3D Model
                                      ↓
                              Summary Report (WinForms)
                                      ↓
                              CSV Export / Clipboard
```

## Project Structure

```
smart3d-excel-import/
├── src/
│   ├── Smart3D.ExcelImport.csproj    # Project file (.NET 4.8)
│   ├── ExcelImportCommand.cs          # Main command (ICommand)
│   ├── SummaryReportForm.cs           # Results dialog (WinForms)
│   ├── Properties/
│   │   └── AssemblyInfo.cs            # COM registration
│   └── Core/
│       ├── Models.cs                  # Data models (ImportRecord, etc.)
│       ├── ExcelParser.cs             # EPPlus Excel reader
│       ├── ImportEngine.cs            # Smart3D property updater
│       ├── Smart3DApplication.cs      # Service access helper
│       ├── TypeCoercer.cs             # String→typed value coercion
│       └── SummaryReportGenerator.cs  # Report generation
├── samples/
│   └── SampleImport.csv               # Sample data
├── docs/
│   ├── DEPLOYMENT.md                  # Build & install guide
│   ├── API_REFERENCES.md              # Smart3D API reference
│   └── USER_GUIDE.md                  # How to use
├── deploy.bat                         # One-click deploy script
└── README.md                          # This file
```

## Excel Format

| Column | Required | Description | Example |
|--------|----------|-------------|---------|
| ObjectName | ✅ | Smart3D object name/tag | `100A-P-101` |
| ObjectType | ❌ | Class type (auto-detected if omitted) | `PipeRun` |
| AttributeName | ✅ | Property to update | `Description` |
| AttributeValue | ✅ | New value | `Process Line` |

### Supported Object Types
- PipeRun, Pipeline, PipeLine
- Equipment, EquipmentComponent
- PipeNozzle, Valve, PipeFitting
- Instrument, StructMember
- HangerSupport, PipeSupport
- CableTray, Ducting

### Automatic Type Coercion
- **Numeric** properties (temp, pressure, weight, diameter) → `double`/`int`
- **Boolean** properties (is*, has*, enabled) → `true`/`false`
- **Date** properties (*date*, *time*) → `DateTime`
- **Everything else** → `string`

## Quick Start

### 1. Build
```bash
cd src
dotnet restore
dotnet build -c Release
```

### 2. Deploy
```bash
deploy.bat
```
(Edit `SMART3D_HOME` path in deploy.bat first)

### 3. Use in Smart3D
1. Restart Smart3D V14
2. Go to **Tools → Custom Commands**
3. Click **"Import Properties from Excel"**
4. Select your `.xlsx` file
5. Review the summary report

## API References Required

| DLL | Purpose |
|-----|---------|
| `Ingr.SP3D.Common.Middle` | Core types, vectors, service manager |
| `Ingr.SP3D.Content` | Model objects, relations, property values |
| `Ingr.SP3D.Content.DataAccess` | Database queries, object filtering |
| `Ingr.SP3D.Content.ServiceManager` | Content service access |
| `Ingr.SP3D.UI` | Command framework (ICommand) |
| `Ingr.SP3D.SystemsAndSpecifications` | Spec-driven design |

## Error Handling

- **Invalid object names** → Logged, skipped, reported
- **Missing attributes** → Logged, skipped, reported
- **Type mismatches** → Attempted coercion, fallback to string
- **COM errors** → Transaction rolled back, error reported
- **All errors** → Written to log file + shown in summary report

## Logging

Logs saved to: `%APPDATA%\Smart3D\ExcelImport\logs\import_YYYYMMDD_HHMMSS.log`

## License

MIT — Use freely for your Smart3D projects.
