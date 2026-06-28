using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.ServiceManager;
using Ingr.SP3D.Content;
using Smart3D.ExcelImport.Models;
using Smart3D.ExcelImport.Services;
using ImportRecord = Smart3D.ExcelImport.Models.ImportRecord;
using ImportResult = Smart3D.ExcelImport.Models.ImportResult;

namespace Smart3D.ExcelImport.Core
{
    /// <summary>
    /// Engine responsible for updating Smart3D object properties from import records.
    /// Handles property setting via IPropertyValues interface and direct property access.
    /// </summary>
    public sealed class PropertyUpdateEngine
    {
        private readonly Application _application;
        private readonly LoggingService _logger;
        private readonly ObjectFinder _objectFinder;
        private readonly TransactionManager _transactionManager;

        // Statistics tracking
        private int _successCount;
        private int _failureCount;
        private int _skipCount;
        private readonly List<ImportResultDetail> _results = new List<ImportResultDetail>();

        public int SuccessCount => _successCount;
        public int FailureCount => _failureCount;
        public int SkipCount => _skipCount;
        public IReadOnlyList<ImportResultDetail> Results => _results;

        public PropertyUpdateEngine(
            Application application,
            LoggingService logger,
            ObjectFinder objectFinder,
            TransactionManager transactionManager)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _objectFinder = objectFinder ?? throw new ArgumentNullException(nameof(objectFinder));
            _transactionManager = transactionManager ?? throw new ArgumentNullException(nameof(transactionManager));
        }

        /// <summary>
        /// Processes all import records and updates the model.
        /// </summary>
        public ImportResult ProcessRecords(List<ImportRecord> records, ImportConfiguration config)
        {
            var result = new ImportResult
            {
                TotalRecordsFound = records.Count,
                Records = records
            };

            _logger.LogInfo($"Starting property update for {records.Count} records...");

            // Pre-load all required objects for efficiency
            var objectNames = records.Select(r => r.ObjectName).Distinct().ToList();
            var objectCache = _objectFinder.FindMultipleObjects(objectNames);

            // Begin transaction for batch update
            var transaction = _transactionManager.BeginTransaction("Bulk Property Import");
            try
            {
                foreach (var record in records)
                {
                    ProcessSingleRecord(record, record.ObjectType, config, objectCache);
                }

                // Commit all changes
                transaction.CommitTransaction();
                result.Success = true;
            }
            catch (Exception ex)
            {
                _logger.LogError("Transaction failed, rolling back.", ex);
                try { transaction.RollbackTransaction(); } catch { }
                result.Success = false;
                result.ErrorMessage = ex.Message;
            }

            result.SuccessCount = _successCount;
            result.FailureCount = _failureCount;
            result.SkippedCount = _skipCount;
            result.Details = _results;

            _logger.LogInfo($"Import complete. Success: {_successCount}, Failed: {_failureCount}, Skipped: {_skipCount}");
            return result;
        }

        private void ProcessSingleRecord(
            ImportRecord record, 
            string objectType,
            ImportConfiguration config,
            Dictionary<string, IObject> objectCache)
        {
            var detail = new ImportResultDetail
            {
                RowNumber = record.RowNumber,
                ObjectName = record.ObjectName,
                AttributeName = record.AttributeName,
                AttributeValue = record.AttributeValue,
                Status = ImportStatus.Processing
            };

            try
            {
                // Find the target object
                if (!objectCache.TryGetValue(record.ObjectName, out var targetObject))
                {
                    detail.Status = ImportStatus.ObjectNotFound;
                    detail.ErrorMessage = $"Object '{record.ObjectName}' not found in model.";
                    _skipCount++;
                    _results.Add(detail);
                    _logger.LogWarning(detail.ErrorMessage);
                    return;
                }

                // Validate object type matches
                if (!string.IsNullOrEmpty(objectType) &&
                    !targetObject.Type.Equals(objectType, StringComparison.OrdinalIgnoreCase))
                {
                    if (config.StrictTypeMatching)
                    {
                        detail.Status = ImportStatus.TypeMismatch;
                        detail.ErrorMessage = $"Type mismatch for '{record.ObjectName}'. " +
                            $"Expected: {objectType}, Actual: {targetObject.Type}";
                        _skipCount++;
                        _results.Add(detail);
                        _logger.LogWarning(detail.ErrorMessage);
                        return;
                    }
                    else
                    {
                        _logger.LogWarning(
                            $"Type mismatch for '{record.ObjectName}' but continuing (strict mode off).");
                    }
                }

                // Update the property
                bool updated = SetPropertyValue(targetObject, record.AttributeName, record.AttributeValue);

                if (updated)
                {
                    detail.Status = ImportStatus.Success;
                    _successCount++;
                    _logger.LogInfo($"Updated {record.ObjectName}.{record.AttributeName} = {record.AttributeValue}");
                }
                else
                {
                    detail.Status = ImportStatus.PropertyNotFound;
                    detail.ErrorMessage = $"Property '{record.AttributeName}' not found on object '{record.ObjectName}'.";
                    _skipCount++;
                    _logger.LogWarning(detail.ErrorMessage);
                }
            }
            catch (COMException ex)
            {
                detail.Status = ImportStatus.Failed;
                detail.ErrorMessage = $"COM Error: {ex.Message}";
                _failureCount++;
                _logger.LogError($"COM error updating {record.ObjectName}.{record.AttributeName}", ex);
            }
            catch (Exception ex)
            {
                detail.Status = ImportStatus.Failed;
                detail.ErrorMessage = ex.Message;
                _failureCount++;
                _logger.LogError($"Error updating {record.ObjectName}.{record.AttributeName}", ex);
            }

            _results.Add(detail);
        }

        /// <summary>
        /// Sets a property value on a Smart3D object using multiple strategies.
        /// </summary>
        private bool SetPropertyValue(IObject targetObject, string propertyName, string value)
        {
            // Strategy 1: Try IPropertyValues interface
            var propertyValues = targetObject as IPropertyValues;
            if (propertyValues != null)
            {
                try
                {
                    if (propertyValues.HasProperty(propertyName))
                    {
                        propertyValues.SetPropertyValue(propertyName, value);
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning($"IPropertyValues.SetPropertyValue failed: {ex.Message}");
                }
            }

            // Strategy 2: Try direct property via reflection/COM
            try
            {
                var prop = targetObject.GetType().GetProperty(propertyName);
                if (prop != null)
                {
                    var convertedValue = ConvertPropertyValue(value, prop.PropertyType);
                    prop.SetValue(targetObject, convertedValue);
                    return true;
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Direct property set failed: {ex.Message}");
            }

            // Strategy 3: Try SetPropertyValue via Middle API
            try
            {
                PropertyHelper.SetPropertyValue(targetObject, propertyName, value);
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"PropertyHelper.SetPropertyValue failed: {ex.Message}");
            }

            // Strategy 4: For PipeRun/Pipeline/Equipment, try COM cast to specific interfaces
            var comObj = targetObject as object;
            if (comObj != null)
            {
                // Try IPipeRun
                var pipeRun = comObj as Ingr.SP3D.Content.IPipeRun;
                if (pipeRun != null)
                    return SetPipeRunProperty(pipeRun, propertyName, value);

                // Try IPipeline
                var pipeline = comObj as Ingr.SP3D.Content.IPipeline;
                if (pipeline != null)
                    return SetPipelineProperty(pipeline, propertyName, value);

                // Try IEquipment
                var equipment = comObj as Ingr.SP3D.Content.IEquipment;
                if (equipment != null)
                    return SetEquipmentProperty(equipment, propertyName, value);
            }

            return false;
        }

        private bool SetPipeRunProperty(Ingr.SP3D.Content.IPipeRun pipeRun, string propertyName, string value)
        {
            try
            {
                // Common PipeRun properties
                switch (propertyName.ToLowerInvariant())
                {
                    case "pipingspec":
                    case "spec":
                        pipeRun.PipingSpec = value;
                        return true;
                    case "pipingspecname":
                        pipeRun.PipingSpecName = value;
                        return true;
                    case "service":
                        pipeRun.Service = value;
                        return true;
                    case "fluid":
                        pipeRun.Fluid = value;
                        return true;
                    case "operatingtemp":
                    case "operatingtemperature":
                        pipeRun.OperatingTemperature = ParseDoubleValue(value);
                        return true;
                    case "designpressure":
                        pipeRun.DesignPressure = ParseDoubleValue(value);
                        return true;
                    default:
                        // Try RunInfo properties
                        pipeRun.RunInfo?.SetPropertyValue(propertyName, value);
                        return true;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"Failed to set PipeRun property '{propertyName}'", ex);
                return false;
            }
        }

        private bool SetPipelineProperty(Ingr.SP3D.Content.IPipeline pipeline, string propertyName, string value)
        {
            try
            {
                switch (propertyName.ToLowerInvariant())
                {
                    case "name":
                    case "pipelinename":
                        pipeline.Name = value;
                        return true;
                    case "description":
                        pipeline.Description = value;
                        return true;
                    case "service":
                        pipeline.Service = value;
                        return true;
                    default:
                        pipeline.SetPropertyValue(propertyName, value);
                        return true;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"Failed to set Pipeline property '{propertyName}'", ex);
                return false;
            }
        }

        private bool SetEquipmentProperty(Ingr.SP3D.Content.IEquipment equipment, string propertyName, string value)
        {
            try
            {
                switch (propertyName.ToLowerInvariant())
                {
                    case "name":
                    case "equipmentname":
                        equipment.Name = value;
                        return true;
                    case "description":
                        equipment.Description = value;
                        return true;
                    case "tag":
                        equipment.Tag = value;
                        return true;
                    case "service":
                        equipment.Service = value;
                        return true;
                    case "operatingweight":
                        equipment.OperatingWeight = ParseDoubleValue(value);
                        return true;
                    case "designweight":
                        equipment.DesignWeight = ParseDoubleValue(value);
                        return true;
                    default:
                        equipment.SetPropertyValue(propertyName, value);
                        return true;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"Failed to set Equipment property '{propertyName}'", ex);
                return false;
            }
        }

        private object ConvertPropertyValue(string value, Type targetType)
        {
            if (targetType == typeof(string))
                return value;
            if (targetType == typeof(double) || targetType == typeof(double?))
                return ParseDoubleValue(value);
            if (targetType == typeof(int) || targetType == typeof(int?))
                return int.Parse(value);
            if (targetType == typeof(bool) || targetType == typeof(bool?))
                return bool.Parse(value);
            if (targetType == typeof(DateTime) || targetType == typeof(DateTime?))
                return DateTime.Parse(value);

            return Convert.ChangeType(value, targetType);
        }

        private double ParseDoubleValue(string value)
        {
            // Handle various number formats
            value = value?.Replace(",", "").Trim() ?? "0";
            if (double.TryParse(value, 
                System.Globalization.NumberStyles.Any, 
                System.Globalization.CultureInfo.InvariantCulture, 
                out var result))
            {
                return result;
            }
            throw new FormatException($"Cannot parse '{value}' as a number.");
        }
    }
}
