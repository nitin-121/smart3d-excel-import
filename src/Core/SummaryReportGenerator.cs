// ============================================================================
// Smart3D V14 - Excel Property Import
// File: Core/SummaryReportGenerator.cs
// Description: Generates formatted summary reports from import results
// ============================================================================

using System;
using System.IO;
using System.Linq;
using System.Text;

namespace Smart3D.ExcelImport.Core
{
    /// <summary>
    /// Generates formatted summary reports from import results
    /// </summary>
    public static class SummaryReportGenerator
    {
        /// <summary>
        /// Generate a plain-text summary report
        /// </summary>
        public static string GenerateTextReport(ImportResult result)
        {
            var sb = new StringBuilder();

            sb.AppendLine("╔══════════════════════════════════════════════════════════════╗");
            sb.AppendLine("║          Smart3D Excel Property Import — Summary            ║");
            sb.AppendLine("╠══════════════════════════════════════════════════════════════╣");
            sb.AppendLine($"║  File:      {Truncate(result.ExcelFilePath, 45),-45} ║");
            sb.AppendLine($"║  Duration:  {result.Duration.TotalSeconds:F1} seconds{"",-37} ║");
            sb.AppendLine("╠══════════════════════════════════════════════════════════════╣");
            sb.AppendLine($"║  Total Records:    {result.TotalCount,-37} ║");
            sb.AppendLine($"║  ✅ Succeeded:     {result.SuccessCount,-37} ║");
            sb.AppendLine($"║  ❌ Failed:        {result.FailureCount,-37} ║");
            sb.AppendLine($"║  ⏭️  Skipped:       {result.SkipCount,-37} ║");
            sb.AppendLine("╠══════════════════════════════════════════════════════════════╣");

            // Success rate
            double rate = result.TotalCount > 0
                ? (double)result.SuccessCount / result.TotalCount * 100
                : 0;
            sb.AppendLine($"║  Success Rate:     {rate:F1}%{"",-36} ║");
            sb.AppendLine("╚══════════════════════════════════════════════════════════════╝");
            sb.AppendLine();

            // Failure details
            var failures = result.Results.Where(r => !r.Success).ToList();
            if (failures.Any())
            {
                sb.AppendLine("─── FAILURE DETAILS ──────────────────────────────────────────");
                sb.AppendLine();
                foreach (var f in failures)
                {
                    sb.AppendLine($"  Row {f.RowNumber}: {f.ObjectName}.{f.AttributeName}");
                    sb.AppendLine($"    Error: {f.ErrorMessage}");
                    sb.AppendLine();
                }
            }

            // Successful updates
            var successes = result.Results.Where(r => r.Success).ToList();
            if (successes.Any())
            {
                sb.AppendLine("─── SUCCESSFUL UPDATES ──────────────────────────────────────");
                sb.AppendLine();
                foreach (var s in successes)
                {
                    sb.AppendLine($"  Row {s.RowNumber}: {s.ObjectName}.{s.AttributeName}");
                    sb.AppendLine($"    {s.OldValue} → {s.NewValue}");
                }
            }

            return sb.ToString();
        }

        /// <summary>
        /// Generate CSV export of all results
        /// </summary>
        public static string GenerateCsvReport(ImportResult result)
        {
            var sb = new StringBuilder();
            sb.AppendLine("RowNumber,ObjectName,AttributeName,OldValue,NewValue,Status,ErrorMessage");

            foreach (var r in result.Results)
            {
                sb.AppendLine($"{r.RowNumber},{Escape(r.ObjectName)},{Escape(r.AttributeName)}," +
                            $"{Escape(r.OldValue)},{Escape(r.NewValue)}," +
                            $"{(r.Success ? "SUCCESS" : "FAILED")},{Escape(r.ErrorMessage)}");
            }

            return sb.ToString();
        }

        /// <summary>
        /// Save CSV report to file
        /// </summary>
        public static string SaveCsvReport(ImportResult result, string outputPath)
        {
            string csv = GenerateCsvReport(result);
            File.WriteAllText(outputPath, csv);
            return outputPath;
        }

        private static string Escape(string value)
        {
            if (string.IsNullOrEmpty(value)) return "\"\"";
            if (value.Contains(",") || value.Contains("\"") || value.Contains("\n"))
                return "\"" + value.Replace("\"", "\"\"") + "\"";
            return "\"" + value + "\"";
        }

        private static string Truncate(string value, int maxLength)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;
            return value.Length <= maxLength ? value : "..." + value.Substring(value.Length - maxLength + 3);
        }
    }
}
