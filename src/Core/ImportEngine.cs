// ============================================================================
// Smart3D V14 - Excel Property Import
// File: Core/ImportEngine.cs
// Description: Core engine that processes ImportRecord objects and updates
//              Smart3D model properties using the V14 API
// ============================================================================

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Content;
using Ingr.SP3D.Content.DataAccess;
using Serilog;

namespace Smart3D.ExcelImport.Core
{
    /// <summary>
    /// Executes bulk property import from parsed Excel records into Smart3D model
    /// </summary>
    public class ImportEngine
    {
        private readonly IModel _model;
        private readonly ILogger _logger;
        private ImportResult _result;

        public ImportEngine(IModel model, ILogger logger)
        {
            _model = model ?? throw new ArgumentNullException(nameof(model));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <summary>
        /// Execute the full import operation
        /// </summary>
        public ImportResult ExecuteImport(List<ImportRecord> records)
        {
            _result = new ImportResult
            {
                StartTime = DateTime.Now,
                TotalCount = records.Count
            };

            _logger.Information("Starting import of {Count} records into model: {Model}",
                records.Count, _model.Name);

            // Begin Smart3D transaction for the entire import
            var txn = Smart3DApplicationHelper.TransactionService;
            try
            {
                txn?.StartTransaction();

                foreach (var record in records)
                {
                    var recordResult = ProcessRecord(record);
                    _result.Results.Add(recordResult);

                    if (recordResult.Success)
                        _result.SuccessCount++;
                    else
                        _result.FailureCount++;
                }

                txn?.CommitTransaction();
                _logger.Information("Transaction committed successfully");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Import transaction failed — rolling back");
                try { txn?.RollbackTransaction(); } catch { }
                throw;
            }
            finally
            {
                _result.EndTime = DateTime.Now;
            }

            return _result;
        }

        /// <summary>
        /// Process a single import record — find object and update property
        /// </summary>
        private RecordResult ProcessRecord(ImportRecord record)
        {
            var result = new RecordResult
            {
                RowNumber = record.RowNumber,
                ObjectName = record.ObjectName,
                AttributeName = record.AttributeName
            };

            try
            {
                // Skip invalid records
                if (!record.IsValid)
                {
                    result.Success = false;
                    result.ErrorMessage = record.ValidationError;
                    _logger.Warning("Row {Row}: Skipped - {Error}", record.RowNumber, record.ValidationError);
                    return result;
                }

                // Step 1: Find the target object in Smart3D model
                IObject targetObject = FindObject(record.ObjectName, record.ObjectType);
                if (targetObject == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Object '{record.ObjectName}' not found in model";
                    _logger.Warning("Row {Row}: Object '{Name}' not found", record.RowNumber, record.ObjectName);
                    return result;
                }

                // Step 2: Get current value for reporting
                result.OldValue = GetPropertyValue(targetObject, record.AttributeName);

                // Step 3: Coerce value to correct type
                if (!TypeCoercer.TryCoerce(record.AttributeValue, record.AttributeName,
                    out object coercedValue, out PropertyDataType dataType))
                {
                    result.Success = false;
                    result.ErrorMessage = $"Cannot coerce value '{record.AttributeValue}' for property '{record.AttributeName}'";
                    _logger.Warning("Row {Row}: Type coercion failed for '{Value}'", record.RowNumber, record.AttributeValue);
                    return result;
                }

                // Step 4: Set the property value
                SetPropertyValue(targetObject, record.AttributeName, coercedValue, dataType);

                result.NewValue = coercedValue.ToString();
                result.Success = true;

                _logger.Information("Row {Row}: Updated {Object}.{Attr} = {NewVal} (was {OldVal})",
                    record.RowNumber, record.ObjectName, record.AttributeName, result.NewValue, result.OldValue);

                // Release COM object
                Smart3DApplicationHelper.ReleaseComObject(targetObject);
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                _logger.Error(ex, "Row {Row}: Error processing {Object}.{Attr}",
                    record.RowNumber, record.ObjectName, record.AttributeName);
            }

            return result;
        }

        /// <summary>
        /// Find a Smart3D object by name and optional type using the Filtering service
        /// </summary>
        private IObject FindObject(string objectName, string objectType)
        {
            try
            {
                var filtering = Smart3DApplicationHelper.FilteringService;
                if (filtering == null)
                {
                    _logger.Error("Filtering service not available");
                    return null;
                }

                // Build filter criteria
                string filterCriteria = $"Name = '{objectName}'";

                // If type specified, add type filter
                if (!string.IsNullOrWhiteSpace(objectType))
                {
                    filterCriteria += $" AND Type = '{objectType}'";
                }

                // Execute filter on active model
                var objects = filtering.GetObjects(_model, filterCriteria);

                if (objects == null || objects.Count == 0)
                {
                    // Fallback: search by relation in model hierarchy
                    return FindObjectByRelation(objectName, objectType);
                }

                return objects[0];
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error finding object '{Name}'", objectName);
                return null;
            }
        }

        /// <summary>
        /// Fallback: Find object by traversing the model hierarchy
        /// </summary>
        private IObject FindObjectByRelation(string objectName, string objectType)
        {
            try
            {
                // Search through systems hierarchy
                var systems = _model.Systems;
                if (systems == null) return null;

                foreach (ISystem system in systems)
                {
                    var obj = SearchSystemHierarchy(system, objectName, objectType);
                    if (obj != null) return obj;
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Hierarchy search failed for '{Name}'", objectName);
            }

            return null;
        }

        /// <summary>
        /// Recursively search a system hierarchy for an object by name
        /// </summary>
        private IObject SearchSystemHierarchy(ISystem system, string objectName, string objectType)
        {
            try
            {
                // Check system members
                var members = system.Members;
                if (members != null)
                {
                    foreach (IObject member in members)
                    {
                        if (member.Name.Equals(objectName, StringComparison.OrdinalIgnoreCase))
                        {
                            if (string.IsNullOrWhiteSpace(objectType) ||
                                member.Type.ToString().Equals(objectType, StringComparison.OrdinalIgnoreCase))
                            {
                                return member;
                            }
                        }
                    }
                }

                // Recurse into child systems
                var childSystems = system.ChildSystems;
                if (childSystems != null)
                {
                    foreach (ISystem child in childSystems)
                    {
                        var found = SearchSystemHierarchy(child, objectName, objectType);
                        if (found != null) return found;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Debug(ex, "Hierarchy traversal error in system '{Name}'", system.Name);
            }

            return null;
        }

        /// <summary>
        /// Get a property value from a Smart3D object
        /// </summary>
        private string GetPropertyValue(IObject obj, string propertyName)
        {
            try
            {
                var propValues = obj.PropertyValues;
                if (propValues == null) return "(none)";

                var propValue = propValues[propertyName];
                if (propValue == null) return "(empty)";

                return propValue.ToString() ?? "(null)";
            }
            catch
            {
                return "(not found)";
            }
        }

        /// <summary>
        /// Set a property value on a Smart3D object
        /// </summary>
        private void SetPropertyValue(IObject obj, string propertyName, object value, PropertyDataType dataType)
        {
            var propValues = obj.PropertyValues;
            if (propValues == null)
                throw new InvalidOperationException($"Object '{obj.Name}' does not support property values");

            // Set value based on detected type
            switch (dataType)
            {
                case PropertyDataType.Integer:
                    propValues[propertyName] = (int)value;
                    break;
                case PropertyDataType.Double:
                    propValues[propertyName] = (double)value;
                    break;
                case PropertyDataType.Boolean:
                    propValues[propertyName] = (bool)value;
                    break;
                case PropertyDataType.Date:
                    propValues[propertyName] = (DateTime)value;
                    break;
                default:
                    propValues[propertyName] = value.ToString();
                    break;
            }

            // Persist changes
            obj.Update();
        }
    }
}
