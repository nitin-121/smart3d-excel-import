// ============================================================================
// Smart3D V14 - Excel Property Import
// File: Core/Models.cs
// Description: Unified data models — combines types from both delegations
// ============================================================================

using System;
using System.Collections.Generic;

namespace Smart3D.ExcelImport.Core
{
    /// <summary>
    /// Represents a single row from the Excel import file
    /// </summary>
    public class ImportRecord
    {
        /// <summary>Row number in Excel file (for error reporting)</summary>
        public int RowNumber { get; set; }

        /// <summary>Name/Tag of the Smart3D object to update</summary>
        public string ObjectName { get; set; } = string.Empty;

        /// <summary>Object class type (PipeRun, Pipeline, Equipment, etc.)</summary>
        public string ObjectType { get; set; } = string.Empty;

        /// <summary>Original object type string from Excel (before mapping)</summary>
        public string OriginalType { get; set; } = string.Empty;

        /// <summary>Property/attribute name to set</summary>
        public string AttributeName { get; set; } = string.Empty;

        /// <summary>Value to assign (will be coerced to correct type)</summary>
        public string AttributeValue { get; set; } = string.Empty;

        /// <summary>Validation status after pre-check</summary>
        public bool IsValid { get; set; } = true;

        /// <summary>Validation error message if invalid</summary>
        public string ValidationError { get; set; } = string.Empty;

        public override string ToString()
        {
            return $"Row {RowNumber}: {ObjectName} ({ObjectType}) - {AttributeName} = {AttributeValue}";
        }
    }

    /// <summary>
    /// Result of processing a single ImportRecord
    /// </summary>
    public class RecordResult
    {
        public int RowNumber { get; set; }
        public string ObjectName { get; set; } = string.Empty;
        public string AttributeName { get; set; } = string.Empty;
        public bool Success { get; set; }
        public string ErrorMessage { get; set; } = string.Empty;
        public string OldValue { get; set; } = string.Empty;
        public string NewValue { get; set; } = string.Empty;
    }

    /// <summary>
    /// Aggregate result of the entire import operation
    /// </summary>
    public class ImportResult
    {
        public bool Success { get; set; }
        public string ErrorMessage { get; set; } = string.Empty;
        public int TotalCount { get; set; }
        public int TotalRecordsFound => Records.Count + FailureCount;
        public int SuccessCount { get; set; }
        public int FailureCount { get; set; }
        public int SkipCount { get; set; }
        public int SkippedCount => SkipCount;
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public DateTime Timestamp { get; set; } = DateTime.Now;
        public TimeSpan Duration => EndTime - StartTime;
        public List<RecordResult> Results { get; set; } = new List<RecordResult>();
        public List<ImportResultDetail> Details { get; set; } = new List<ImportResultDetail>();
        public string ExcelFilePath { get; set; } = string.Empty;
        public string LogFilePath { get; set; } = string.Empty;
        public string ReportPath { get; set; } = string.Empty;
    }

    /// <summary>
    /// Enhanced import result with detailed per-record status (used by BulkPropertyImportCommand)
    /// </summary>
    public class ImportResultDetail
    {
        public int RowNumber { get; set; }
        public string ObjectName { get; set; } = string.Empty;
        public string AttributeName { get; set; } = string.Empty;
        public string AttributeValue { get; set; } = string.Empty;
        public ImportStatus Status { get; set; }
        public string ErrorMessage { get; set; } = string.Empty;
    }

    /// <summary>
    /// Extended import result class for the enhanced pipeline
    /// </summary>
    public class ImportResultEx
    {
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
        public int TotalRecordsFound { get; set; }
        public int SuccessCount { get; set; }
        public int FailureCount { get; set; }
        public int SkippedCount { get; set; }
        public List<ImportRecord> Records { get; set; } = new List<ImportRecord>();
        public List<ImportResultDetail> Details { get; set; } = new List<ImportResultDetail>();
        public string ReportPath { get; set; }
        public DateTime Timestamp { get; set; } = DateTime.Now;

        public override string ToString()
        {
            return $"Import Result: Success={Success}, Total={TotalRecordsFound}, " +
                   $"Updated={SuccessCount}, Failed={FailureCount}, Skipped={SkippedCount}";
        }
    }

    /// <summary>
    /// Configuration options for the import operation
    /// </summary>
    public class ImportConfiguration
    {
        public string ExcelFilePath { get; set; } = string.Empty;
        public bool StrictTypeMatching { get; set; } = false;
        public bool SkipDuplicates { get; set; } = true;
        public bool CreateBackup { get; set; } = true;
        public int MaxErrors { get; set; } = 100;
        public bool ValidateProperties { get; set; } = true;
        public string LogFilePath { get; set; } = string.Empty;
        public bool GenerateReport { get; set; } = true;
        public string ReportOutputPath { get; set; } = string.Empty;
    }

    /// <summary>
    /// Status of a single record import operation
    /// </summary>
    public enum ImportStatus
    {
        Processing,
        Success,
        Failed,
        ObjectNotFound,
        TypeMismatch,
        PropertyNotFound,
        Skipped
    }

    /// <summary>
    /// Supported Smart3D object types for property import
    /// </summary>
    public static class SupportedObjectTypes
    {
        public const string PipeRun = "PipeRun";
        public const string Pipeline = "Pipeline";
        public const string PipeLine = "PipeLine";
        public const string Equipment = "Equipment";
        public const string PipeNozzle = "PipeNozzle";
        public const string Instrument = "Instrument";
        public const string StructMember = "StructMember";
        public const string HangerSupport = "HangerSupport";
        public const string Valve = "Valve";
        public const string PipeFitting = "PipeFitting";
        public const string EquipmentComponent = "EquipmentComponent";
        public const string PipeSupport = "PipeSupport";
        public const string CableTray = "CableTray";
        public const string Ducting = "Ducting";

        public static readonly HashSet<string> All = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            PipeRun, Pipeline, PipeLine, Equipment, PipeNozzle,
            Instrument, StructMember, HangerSupport, Valve,
            PipeFitting, EquipmentComponent, PipeSupport, CableTray, Ducting
        };

        public static bool IsSupported(string objectType) =>
            !string.IsNullOrWhiteSpace(objectType) && All.Contains(objectType.Trim());
    }

    /// <summary>
    /// Data type coercion targets for property values
    /// </summary>
    public enum PropertyDataType
    {
        String,
        Integer,
        Double,
        Boolean,
        Date
    }

}
