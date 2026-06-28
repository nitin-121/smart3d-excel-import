// ============================================================================
// Smart3D V14 - Excel Property Import
// File: Core/TypeCoercer.cs
// Description: Coerces string values from Excel to appropriate .NET types
//              matching Smart3D property data types
// ============================================================================

using System;
using System.Globalization;

namespace Smart3D.ExcelImport.Core
{
    /// <summary>
    /// Coerces string values from Excel into typed objects for Smart3D properties
    /// </summary>
    public static class TypeCoercer
    {
        /// <summary>
        /// Attempts to coerce a string value to the appropriate type based on the property name
        /// and common Smart3D property naming conventions
        /// </summary>
        /// <param name="value">String value from Excel</param>
        /// <param name="propertyName">Name of the property (for type inference)</param>
        /// <param name="result">Coerced value as object</param>
        /// <param name="dataType">Detected data type</param>
        /// <returns>True if coercion succeeded</returns>
        public static bool TryCoerce(string value, string propertyName, out object result, out PropertyDataType dataType)
        {
            result = null;
            dataType = PropertyDataType.String;

            if (string.IsNullOrWhiteSpace(value))
            {
                result = string.Empty;
                dataType = PropertyDataType.String;
                return true;
            }

            string trimmed = value.Trim();

            // Infer type from property name patterns
            if (IsNumericProperty(propertyName))
            {
                if (TryCoerceNumeric(trimmed, out object numericResult, out PropertyDataType numericType))
                {
                    result = numericResult;
                    dataType = numericType;
                    return true;
                }
            }

            if (IsBooleanProperty(propertyName))
            {
                if (TryCoerceBoolean(trimmed, out object boolResult))
                {
                    result = boolResult;
                    dataType = PropertyDataType.Boolean;
                    return true;
                }
            }

            if (IsDateProperty(propertyName))
            {
                if (TryCoerceDate(trimmed, out object dateResult))
                {
                    result = dateResult;
                    dataType = PropertyDataType.Date;
                    return true;
                }
            }

            // Default: string
            result = trimmed;
            dataType = PropertyDataType.String;
            return true;
        }

        /// <summary>
        /// Coerce to numeric type (int or double)
        /// </summary>
        private static bool TryCoerceNumeric(string value, out object result, out PropertyDataType dataType)
        {
            result = null;

            // Try integer first
            if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int intVal))
            {
                result = intVal;
                dataType = PropertyDataType.Integer;
                return true;
            }

            // Try double (handles decimals, scientific notation)
            if (double.TryParse(value, NumberStyles.Float | NumberStyles.AllowThousands,
                CultureInfo.InvariantCulture, out double dblVal))
            {
                result = dblVal;
                dataType = PropertyDataType.Double;
                return true;
            }

            return false;
        }

        /// <summary>
        /// Coerce to boolean
        /// </summary>
        private static bool TryCoerceBoolean(string value, out object result)
        {
            result = null;

            string lower = value.ToLowerInvariant();

            if (lower == "true" || lower == "yes" || lower == "1" || lower == "y" || lower == "on")
            {
                result = true;
                return true;
            }

            if (lower == "false" || lower == "no" || lower == "0" || lower == "n" || lower == "off")
            {
                result = false;
                return true;
            }

            return false;
        }

        /// <summary>
        /// Coerce to DateTime
        /// </summary>
        private static bool TryCoerceDate(string value, out object result)
        {
            result = null;

            // Try common date formats
            string[] formats = {
                "yyyy-MM-dd", "MM/dd/yyyy", "dd/MM/yyyy",
                "yyyy/MM/dd", "dd-MMM-yyyy", "MMM dd, yyyy",
                "yyyy-MM-dd HH:mm:ss", "MM/dd/yyyy HH:mm:ss"
            };

            if (DateTime.TryParseExact(value, formats, CultureInfo.InvariantCulture,
                DateTimeStyles.None, out DateTime dateVal))
            {
                result = dateVal;
                return true;
            }

            // Fallback to general parsing
            if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out dateVal))
            {
                result = dateVal;
                return true;
            }

            return false;
        }

        /// <summary>
        /// Infer if property is numeric based on naming conventions
        /// </summary>
        private static bool IsNumericProperty(string propertyName)
        {
            if (string.IsNullOrWhiteSpace(propertyName)) return false;

            string[] numericPatterns = {
                "temp", "pressure", "weight", "length", "diameter", "thickness",
                "rating", "flow", "volume", "elevation", "angle", "distance",
                "size", "count", "number", "qty", "quantity", "capacity",
                "power", "voltage", "current", "frequency", "speed", "rpm",
                "load", "force", "stress", "strain", "modulus", "density",
                "min", "max", "min.", "max.", "nominal", "design"
            };

            string lower = propertyName.ToLowerInvariant();
            return Array.Exists(numericPatterns, p => lower.Contains(p));
        }

        /// <summary>
        /// Infer if property is boolean based on naming conventions
        /// </summary>
        private static bool IsBooleanProperty(string propertyName)
        {
            if (string.IsNullOrWhiteSpace(propertyName)) return false;

            string[] boolPatterns = {
                "is", "has", "can", "enabled", "active", "visible", "required",
                "insulated", "traced", "fireproof", "seismic", "critical",
                "tagged", "locked", "checked", "approved", "verified"
            };

            string lower = propertyName.ToLowerInvariant();
            return Array.Exists(boolPatterns, p =>
                lower.StartsWith(p) || lower.EndsWith(p) || lower.Contains(" " + p + " "));
        }

        /// <summary>
        /// Infer if property is date-based
        /// </summary>
        private static bool IsDateProperty(string propertyName)
        {
            if (string.IsNullOrWhiteSpace(propertyName)) return false;

            string[] datePatterns = {
                "date", "time", "created", "modified", "issued", "approved",
                "reviewed", "commissioned", "installed", "tested", "inspected"
            };

            string lower = propertyName.ToLowerInvariant();
            return Array.Exists(datePatterns, p => lower.Contains(p));
        }
    }
}
