// ============================================================================
// Smart3D V14 - Excel Property Import
// File: Core/ObjectTypeMapper.cs
// Description: Maps user-friendly object type names to Smart3D API types
// ============================================================================

using System;
using System.Collections.Generic;

namespace Smart3D.ExcelImport.Core
{
    /// <summary>
    /// Maps user-friendly object type names from Excel to Smart3D API type identifiers.
    /// </summary>
    public static class ObjectTypeMapper
    {
        private static readonly Dictionary<string, string> TypeMappings =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            // PipeRun variations
            { "PipeRun", "PipeRun" },
            { "Pipe Run", "PipeRun" },
            { "Piping", "PipeRun" },
            { "Pipe", "PipeRun" },
            { "RUN", "PipeRun" },

            // Pipeline variations
            { "Pipeline", "Pipeline" },
            { "Pipe Line", "Pipeline" },
            { "Line", "Pipeline" },
            { "PipingLine", "Pipeline" },

            // Equipment variations
            { "Equipment", "Equipment" },
            { "Equip", "Equipment" },
            { "Vessel", "Equipment" },
            { "Tank", "Equipment" },
            { "Pump", "Equipment" },
            { "HeatExchanger", "Equipment" },
            { "Heat Exchanger", "Equipment" },
            { "Column", "Equipment" },
            { "Tower", "Equipment" },

            // Support variations
            { "Support", "Support" },
            { "Hanger", "Support" },
            { "Restraint", "Support" },

            // Generic
            { "Component", "Component" },
            { "Part", "Component" },
            { "Object", "Object" },
        };

        /// <summary>
        /// Maps a user-friendly type name to the Smart3D API type identifier.
        /// </summary>
        public static string MapToSmart3DType(string userType)
        {
            if (string.IsNullOrWhiteSpace(userType))
                return string.Empty;

            if (TypeMappings.TryGetValue(userType.Trim(), out var mappedType))
                return mappedType;

            // If no mapping found, return the original (might be a valid Smart3D type already)
            return userType.Trim();
        }

        /// <summary>
        /// Returns all supported type names for UI display.
        /// </summary>
        public static IEnumerable<string> GetSupportedTypes()
        {
            return TypeMappings.Keys;
        }

        /// <summary>
        /// Checks if a type name is supported.
        /// </summary>
        public static bool IsSupportedType(string typeName)
        {
            return !string.IsNullOrWhiteSpace(typeName) && TypeMappings.ContainsKey(typeName);
        }

        /// <summary>
        /// Gets the Smart3D numeric object type value for filtering.
        /// </summary>
        public static int GetSmart3DObjectTypeValue(string typeName)
        {
            var mapped = MapToSmart3DType(typeName);

            switch (mapped.ToLowerInvariant())
            {
                case "piperun":
                    return 1; // SP3D ObjectType constant
                case "pipeline":
                    return 2;
                case "equipment":
                    return 3;
                case "support":
                    return 4;
                case "component":
                    return 5;
                default:
                    return 0; // Generic object
            }
        }
    }
}
