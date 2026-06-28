# User Guide — Smart3D Excel Property Import

## What This Tool Does

Bulk-update Smart3D object properties from an Excel spreadsheet. Instead of manually editing each object one-by-one, you can update hundreds of properties in a single operation.

**Example use cases:**
- Update descriptions for 500 pipe runs from a project register
- Set insulation types based on a process engineer's spreadsheet
- Bulk-update design temperatures and pressures from a line list
- Import equipment data from vendor Excel sheets
- Apply attribute values from a P&ID review

## Preparing Your Excel File

### Required Columns

| Column | Description | Example |
|--------|-------------|---------|
| **ObjectName** | The Smart3D object name/tag | `100A-P-101` |
| **AttributeName** | The property to update | `Description` |
| **AttributeValue** | The new value | `Crude Oil Feed Line` |

### Optional Columns

| Column | Description | Example |
|--------|-------------|---------|
| **ObjectType** | The class type (auto-detected if omitted) | `PipeRun` |

### Column Header Flexibility

The parser accepts multiple header names:

| Field | Accepted Headers |
|-------|-----------------|
| ObjectName | `ObjectName`, `Object Name`, `Tag`, `TagNo`, `Name` |
| ObjectType | `ObjectType`, `Object Type`, `Type`, `Class` |
| AttributeName | `AttributeName`, `Attribute Name`, `Attribute`, `Property` |
| AttributeValue | `AttributeValue`, `Attribute Value`, `Value`, `DataValue` |

### Sample Excel Content

| ObjectName | ObjectType | AttributeName | AttributeValue |
|------------|------------|---------------|----------------|
| 100A-P-101 | PipeRun | Description | Process Line - Crude Oil |
| 100A-P-101 | PipeRun | InsulationType | HPS |
| 100A-P-101 | PipeRun | DesignTemp | 150 |
| 200B-E-201 | Equipment | Description | Heat Exchanger |
| 200B-E-201 | Equipment | DesignPressure | 8.2 |
| 300C-VL-301 | Valve | IsFailOpen | true |

## Step-by-Step Usage

### 1. Prepare Data
- Create your Excel file with the required columns
- Save as `.xlsx` format (not `.xls`)
- Ensure object names match exactly what's in Smart3D

### 2. Open Smart3D Model
- Start Smart3D V14
- Open the model containing your objects
- Ensure you have write access

### 3. Run the Command
- Go to **Tools → Custom Commands**
- Click **"Import Properties from Excel"**
- Select your Excel file in the file picker
- Click **Open**

### 4. Confirm Import
- A dialog shows how many records were found
- Click **Yes** to proceed or **No** to cancel

### 5. Review Results
- The summary report shows:
  - Total records processed
  - Success count (green)
  - Failed count (red)
  - Duration
- Each row shows: Object, Attribute, Old Value → New Value, Status

### 6. Export Report (Optional)
- Click **Export CSV** to save results to a file
- Click **Copy Report** to copy text summary to clipboard

## Understanding Results

### Success (Green)
Property was found and updated successfully.

### Failed (Red) — Common Reasons

| Error | Cause | Fix |
|-------|-------|-----|
| "Object not found" | Name doesn't match Smart3D | Check spelling, case sensitivity |
| "Attribute not found" | Property doesn't exist on object | Verify property name in catalog |
| "Type coercion failed" | Value can't be converted | Check numeric/boolean formats |
| "Permission denied" | Object is locked or read-only | Check out object first |
| "Transaction failed" | Database error | Check Smart3D logs |

## Best Practices

1. **Backup first** — Always backup your model before bulk imports
2. **Test with 5-10 rows** — Validate with a small batch first
3. **Use ObjectType** — Speeds up lookup, avoids ambiguity
4. **Check log files** — `%APPDATA%\Smart3D\ExcelImport\logs\`
5. **One attribute per row** — Don't combine multiple attributes in one row
6. **Avoid special characters** — In object names, use only alphanumeric + hyphens

## Type Coercion Rules

The tool automatically detects the correct data type:

| Property Contains | Detected Type | Examples |
|------------------|---------------|----------|
| temp, pressure, weight, diameter, size... | Numeric (double) | `150.5`, `1000` |
| is*, has*, enabled, active... | Boolean | `true/false`, `yes/no`, `1/0` |
| date, time, created, modified... | DateTime | `2026-06-27`, `06/27/2026` |
| Everything else | String | `HPS`, `Process Line` |

## Limitations

- Maximum rows: Limited by available memory (tested up to 10,000)
- File format: `.xlsx` only (not legacy `.xls`)
- COM objects: Must be released properly (handled automatically)
- Transactions: All-or-nothing — if one record fails, the batch rolls back

## Support

- Log files: `%APPDATA%\Smart3D\ExcelImport\logs\`
- Sample data: `samples/SampleImport.csv`
- API docs: `docs/API_REFERENCES.md`
