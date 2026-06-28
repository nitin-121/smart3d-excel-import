// ============================================================================
// Smart3D V14 - Excel Property Import
// File: Core/ExcelParser.cs
// Description: Parses Excel (.xlsx) files using EPPlus library.
//              Expects columns: ObjectName, ObjectType, AttributeName, AttributeValue
// ============================================================================

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using Serilog;

namespace Smart3D.ExcelImport.Core
{
    /// <summary>
    /// Parses Excel files into ImportRecord objects for Smart3D property updates
    /// </summary>
    public class ExcelParser
    {
        private readonly ILogger _logger;

        // Expected column headers (case-insensitive matching)
        private static readonly string[] NameColumns = { "ObjectName", "Object Name", "Tag", "TagNo", "Name" };
        private static readonly string[] TypeColumns = { "ObjectType", "Object Type", "Type", "Class" };
        private static readonly string[] AttrColumns = { "AttributeName", "Attribute Name", "Attribute", "Property", "Property Name" };
        private static readonly string[] ValueColumns = { "AttributeValue", "Attribute Value", "Value", "DataValue" };

        public ExcelParser(ILogger logger)
        {
            _logger = logger;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        /// <summary>
        /// Parse an Excel file and return validated import records
        /// </summary>
        /// <param name="filePath">Full path to .xlsx file</param>
        /// <returns>List of validated ImportRecord objects</returns>
        public List<ImportRecord> Parse(string filePath)
        {
            var records = new List<ImportRecord>();

            _logger.Information("Opening Excel file: {Path}", filePath);

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                // Get the first worksheet
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null)
                {
                    throw new InvalidDataException("Excel file contains no worksheets.");
                }

                _logger.Information("Reading worksheet: {Name} ({Rows} rows x {Cols} cols)",
                    worksheet.Name, worksheet.Dimension?.Rows ?? 0, worksheet.Dimension?.Columns ?? 0);

                if (worksheet.Dimension == null || worksheet.Dimension.Rows < 2)
                {
                    throw new InvalidDataException("Excel file has no data rows (header only or empty).");
                }

                // Step 1: Map column headers to indices
                var columnMap = MapColumns(worksheet);
                ValidateColumnMap(columnMap);

                int rowCount = worksheet.Dimension.Rows;
                int skippedRows = 0;

                // Step 2: Parse data rows (starting from row 2, row 1 = headers)
                for (int row = 2; row <= rowCount; row++)
                {
                    var record = ParseRow(worksheet, row, columnMap);

                    // Skip completely empty rows
                    if (string.IsNullOrWhiteSpace(record.ObjectName) &&
                        string.IsNullOrWhiteSpace(record.AttributeName))
                    {
                        skippedRows++;
                        continue;
                    }

                    // Validate the record
                    ValidateRecord(record);
                    records.Add(record);
                }

                _logger.Information("Parsed {Count} records, skipped {Skipped} empty rows",
                    records.Count, skippedRows);
            }

            return records;
        }

        /// <summary>
        /// Maps Excel column headers to expected field names
        /// </summary>
        private Dictionary<string, int> MapColumns(ExcelWorksheet worksheet)
        {
            var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            int colCount = worksheet.Dimension.Columns;

            for (int col = 1; col <= colCount; col++)
            {
                string header = worksheet.Cells[1, col].Text?.Trim() ?? string.Empty;

                if (NameColumns.Any(h => h.Equals(header, StringComparison.OrdinalIgnoreCase)))
                    map["ObjectName"] = col;
                else if (TypeColumns.Any(h => h.Equals(header, StringComparison.OrdinalIgnoreCase)))
                    map["ObjectType"] = col;
                else if (AttrColumns.Any(h => h.Equals(header, StringComparison.OrdinalIgnoreCase)))
                    map["AttributeName"] = col;
                else if (ValueColumns.Any(h => h.Equals(header, StringComparison.OrdinalIgnoreCase)))
                    map["AttributeValue"] = col;
            }

            _logger.Debug("Column mapping: {Mapping}",
                string.Join(", ", map.Select(kvp => $"{kvp.Key}→Col{kvp.Value}")));

            return map;
        }

        /// <summary>
        /// Validates that all required columns were found
        /// </summary>
        private void ValidateColumnMap(Dictionary<string, int> map)
        {
            var missing = new List<string>();

            if (!map.ContainsKey("ObjectName"))
                missing.Add("ObjectName");
            if (!map.ContainsKey("AttributeName"))
                missing.Add("AttributeName");
            if (!map.ContainsKey("AttributeValue"))
                missing.Add("AttributeValue");

            if (missing.Any())
            {
                throw new InvalidDataException(
                    $"Missing required columns: {string.Join(", ", missing)}. " +
                    $"Expected headers: ObjectName, ObjectType, AttributeName, AttributeValue");
            }

            // ObjectType is optional — will be auto-detected if missing
            if (!map.ContainsKey("ObjectType"))
            {
                _logger.Warning("ObjectType column not found — will attempt auto-detection from model.");
            }
        }

        /// <summary>
        /// Parse a single row into an ImportRecord
        /// </summary>
        private ImportRecord ParseRow(ExcelWorksheet worksheet, int row, Dictionary<string, int> columnMap)
        {
            string GetCellValue(string field)
            {
                if (!columnMap.TryGetValue(field, out int col)) return string.Empty;
                return worksheet.Cells[row, col].Text?.Trim() ?? string.Empty;
            }

            return new ImportRecord
            {
                RowNumber = row,
                ObjectName = GetCellValue("ObjectName"),
                ObjectType = GetCellValue("ObjectType"),
                AttributeName = GetCellValue("AttributeName"),
                AttributeValue = GetCellValue("AttributeValue")
            };
        }

        /// <summary>
        /// Validate a single import record
        /// </summary>
        private void ValidateRecord(ImportRecord record)
        {
            var errors = new List<string>();

            if (string.IsNullOrWhiteSpace(record.ObjectName))
                errors.Add("ObjectName is required");

            if (string.IsNullOrWhiteSpace(record.AttributeName))
                errors.Add("AttributeName is required");

            if (string.IsNullOrWhiteSpace(record.AttributeValue))
                errors.Add("AttributeValue is required");

            // Validate object type if provided
            if (!string.IsNullOrWhiteSpace(record.ObjectType) &&
                !SupportedObjectTypes.IsSupported(record.ObjectType))
            {
                errors.Add($"ObjectType '{record.ObjectType}' is not supported. " +
                          $"Supported: {string.Join(", ", SupportedObjectTypes.All)}");
            }

            if (errors.Any())
            {
                record.IsValid = false;
                record.ValidationError = string.Join("; ", errors);
                _logger.Warning("Row {Row} validation failed: {Errors}",
                    record.RowNumber, record.ValidationError);
            }
        }
    }
}
