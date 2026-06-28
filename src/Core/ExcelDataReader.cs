using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

using Smart3D.ExcelImport.Services;
using Smart3D.ExcelImport.Core;

namespace Smart3D.ExcelImport.Core
{
    /// Reads and validates Excel spreadsheet data for bulk property import.
    /// Supports .xlsx format with EPPlus library.
    /// Expected columns: ObjectName, ObjectType, AttributeName, AttributeValue
    /// Uses column header mapping for flexible column ordering.
    /// </summary>
    public sealed class ExcelDataReader : IDisposable
    {
        private readonly string _filePath;
        private readonly LoggingService _logger;
        private ExcelPackage _package;
        private bool _disposed;

        // Expected column headers (case-insensitive matching)
        private static readonly string[] NameColumns = { "ObjectName", "Object Name", "Tag", "TagNo", "Name" };
        private static readonly string[] TypeColumns = { "ObjectType", "Object Type", "Type", "Class" };
        private static readonly string[] AttrColumns = { "AttributeName", "Attribute Name", "Attribute", "Property", "Property Name" };
        private static readonly string[] ValueColumns = { "AttributeValue", "Attribute Value", "Value", "DataValue" };

        /// <summary>
        /// Initializes the Excel data reader.
        /// </summary>
        /// <param name="filePath">Full path to the Excel file.</param>
        /// <param name="logger">Logger for diagnostic output.</param>
        public ExcelDataReader(string filePath, LoggingService logger)
        {
            _filePath = filePath ?? throw new ArgumentNullException(nameof(filePath));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));

            if (!File.Exists(filePath))
                throw new FileNotFoundException("Excel file not found.", filePath);

            if (!Path.GetExtension(filePath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Only .xlsx files are supported.", nameof(filePath));

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        /// <summary>
        /// Reads all import records from the Excel file.
        /// </summary>
        /// <returns>List of validated import records.</returns>
        public List<ImportRecord> ReadAllRecords()
        {
            _logger.LogInfo($"Opening Excel file: {_filePath}");

            _package = new ExcelPackage(new FileInfo(_filePath));
            var worksheet = _package.Workbook.Worksheets.FirstOrDefault()
                ?? throw new InvalidDataException("Excel file contains no worksheets.");

            var records = new List<ImportRecord>();
            var rowCount = worksheet.Dimension?.Rows ?? 0;
            var colCount = worksheet.Dimension?.Columns ?? 0;

            _logger.LogInfo($"Worksheet '{worksheet.Name}' has {rowCount} rows and {colCount} columns.");

            if (colCount < 3)
                throw new InvalidDataException(
                    "Excel file must have at least 3 columns: ObjectName, AttributeName, AttributeValue");

            // Map column headers to indices
            var columnMap = MapColumns(worksheet);

            // Read data rows (skip header row 1)
            for (int row = 2; row <= rowCount; row++)
            {
                var record = ParseRow(worksheet, row, columnMap);
                if (record != null)
                {
                    records.Add(record);
                }
            }

            _logger.LogInfo($"Successfully read {records.Count} valid records from Excel.");
            return records;
        }

        /// <summary>
        /// Reads records with validation and returns a detailed result set.
        /// </summary>
        public ImportResult ReadRecordsWithValidation()
        {
            var result = new ImportResult();

            try
            {
                _package = new ExcelPackage(new FileInfo(_filePath));
                var worksheet = _package.Workbook.Worksheets.FirstOrDefault()
                    ?? throw new InvalidDataException("Excel file contains no worksheets.");

                var columnMap = MapColumns(worksheet);
                var rowCount = worksheet.Dimension?.Rows ?? 0;

                for (int row = 2; row <= rowCount; row++)
                {
                    try
                    {
                        var record = ParseRow(worksheet, row, columnMap);
                        if (record != null)
                        {
                            record.RowNumber = row;
                            result.Records.Add(record);
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning($"Row {row}: {ex.Message}");
                        result.FailureCount++;
                    }
                }

                result.Success = true;
                result.TotalRecordsFound = result.Records.Count + result.FailureCount;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                _logger.LogError("Failed to read Excel file", ex);
            }

            return result;
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

            // Validate required columns
            if (!map.ContainsKey("ObjectName"))
                throw new InvalidDataException("Missing required column: ObjectName");
            if (!map.ContainsKey("AttributeName"))
                throw new InvalidDataException("Missing required column: AttributeName");
            if (!map.ContainsKey("AttributeValue"))
                throw new InvalidDataException("Missing required column: AttributeValue");

            return map;
        }

        private ImportRecord ParseRow(ExcelWorksheet worksheet, int row, Dictionary<string, int> columnMap)
        {
            string GetCellValue(string field)
            {
                if (!columnMap.TryGetValue(field, out int col)) return string.Empty;
                return worksheet.Cells[row, col].Text?.Trim() ?? string.Empty;
            }

            var objectName = GetCellValue("ObjectName");
            var objectType = GetCellValue("ObjectType");
            var attributeName = GetCellValue("AttributeName");
            var attributeValue = GetCellValue("AttributeValue");

            // Skip empty rows
            if (string.IsNullOrWhiteSpace(objectName) &&
                string.IsNullOrWhiteSpace(attributeName))
            {
                return null;
            }

            // Validate required fields
            if (string.IsNullOrWhiteSpace(objectName))
                throw new FormatException("ObjectName is required.");
            if (string.IsNullOrWhiteSpace(attributeName))
                throw new FormatException("AttributeName is required.");

            // ObjectType is optional — if provided, validate against supported types
            if (!string.IsNullOrWhiteSpace(objectType) && !SupportedObjectTypes.IsSupported(objectType))
            {
                _logger.LogWarning($"Row {row}: ObjectType '{objectType}' is not in the supported types list.");
            }

            return new ImportRecord
            {
                ObjectName = objectName,
                ObjectType = objectType,
                AttributeName = attributeName,
                AttributeValue = attributeValue,
                RowNumber = row
            };
        }

        #region IDisposable
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    _package?.Dispose();
                }
                _disposed = true;
            }
        }
        #endregion
    }
}
