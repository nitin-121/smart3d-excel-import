using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Smart3D.ExcelImport.Core;

namespace Smart3D.ExcelImport.Services
{
    /// <summary>
    /// Generates summary reports for the import operation.
    /// Supports both text and HTML output formats.
    /// </summary>
    public sealed class ReportGenerator
    {
        private readonly LoggingService _logger;

        public ReportGenerator(LoggingService logger)
        {
            _logger = logger;
        }

        /// <summary>
        /// Generates a text summary report.
        /// </summary>
        public string GenerateTextReport(ImportResult result, ImportConfiguration config)
        {
            var sb = new StringBuilder();
            sb.AppendLine("=".PadRight(80, '='));
            sb.AppendLine("  SMART3D V14 - BULK PROPERTY IMPORT SUMMARY REPORT");
            sb.AppendLine("=".PadRight(80, '='));
            sb.AppendLine();
            sb.AppendLine($"  Generated: {result.Timestamp:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"  Excel File: {config.ExcelFilePath}");
            sb.AppendLine();
            sb.AppendLine("-".PadRight(80, '-'));
            sb.AppendLine("  SUMMARY");
            sb.AppendLine("-".PadRight(80, '-'));
            sb.AppendLine($"  Total Records:    {result.TotalRecordsFound}");
            sb.AppendLine($"  Updated:          {result.SuccessCount}");
            sb.AppendLine($"  Failed:           {result.FailureCount}");
            sb.AppendLine($"  Skipped:          {result.SkippedCount}");
            sb.AppendLine($"  Success Rate:     {(result.TotalRecordsFound > 0 ? (double)result.SuccessCount / result.TotalRecordsFound * 100 : 0):F1}%");
            sb.AppendLine();

            if (result.Details != null && result.Details.Count > 0)
            {
                sb.AppendLine("-".PadRight(80, '-'));
                sb.AppendLine("  DETAILS");
                sb.AppendLine("-".PadRight(80, '-'));

                // Group by status
                var grouped = result.Details.GroupBy(d => d.Status).OrderBy(g => g.Key);
                foreach (var group in grouped)
                {
                    sb.AppendLine();
                    sb.AppendLine($"  [{group.Key}] ({group.Count()} items)");
                    foreach (var detail in group)
                    {
                        sb.AppendLine($"    Row {detail.RowNumber}: {detail.ObjectName}.{detail.AttributeName} = {detail.AttributeValue}");
                        if (!string.IsNullOrEmpty(detail.ErrorMessage))
                            sb.AppendLine($"      -> {detail.ErrorMessage}");
                    }
                }
            }

            sb.AppendLine();
            sb.AppendLine("=".PadRight(80, '='));
            sb.AppendLine("  END OF REPORT");
            sb.AppendLine("=".PadRight(80, '='));

            return sb.ToString();
        }

        /// <summary>
        /// Generates an HTML summary report.
        /// </summary>
        public string GenerateHtmlReport(ImportResult result, ImportConfiguration config)
        {
            var successRate = result.TotalRecordsFound > 0 
                ? (double)result.SuccessCount / result.TotalRecordsFound * 100 
                : 0;

            var html = new StringBuilder();
            html.AppendLine("<!DOCTYPE html>");
            html.AppendLine("<html><head><title>Smart3D Import Report</title>");
            html.AppendLine("<style>");
            html.AppendLine(@"
                body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; background: #f5f5f5; }
                .container { max-width: 1200px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
                h1 { color: #2c3e50; border-bottom: 3px solid #3498db; padding-bottom: 10px; }
                h2 { color: #34495e; margin-top: 30px; }
                .summary-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin: 20px 0; }
                .summary-card { padding: 20px; border-radius: 8px; text-align: center; color: white; }
                .card-total { background: #3498db; }
                .card-success { background: #27ae60; }
                .card-failed { background: #e74c3c; }
                .card-skipped { background: #f39c12; }
                .card-number { font-size: 36px; font-weight: bold; }
                .card-label { font-size: 14px; opacity: 0.9; }
                table { width: 100%; border-collapse: collapse; margin-top: 15px; }
                th { background: #34495e; color: white; padding: 12px; text-align: left; }
                td { padding: 10px 12px; border-bottom: 1px solid #ecf0f1; }
                tr:hover { background: #f8f9fa; }
                .status-success { color: #27ae60; font-weight: bold; }
                .status-failed { color: #e74c3c; font-weight: bold; }
                .status-skipped { color: #f39c12; font-weight: bold; }
                .meta-info { color: #7f8c8d; font-size: 13px; }
            ");
            html.AppendLine("</style></head><body>");
            html.AppendLine("<div class='container'>");
            html.AppendLine("<h1>Smart3D V14 - Bulk Property Import Report</h1>");
            html.AppendLine($"<p class='meta-info'>Generated: {result.Timestamp:yyyy-MM-dd HH:mm:ss} | File: {config.ExcelFilePath}</p>");

            // Summary cards
            html.AppendLine("<div class='summary-grid'>");
            html.AppendLine($"<div class='summary-card card-total'><div class='card-number'>{result.TotalRecordsFound}</div><div class='card-label'>Total Records</div></div>");
            html.AppendLine($"<div class='summary-card card-success'><div class='card-number'>{result.SuccessCount}</div><div class='card-label'>Updated</div></div>");
            html.AppendLine($"<div class='summary-card card-failed'><div class='card-number'>{result.FailureCount}</div><div class='card-label'>Failed</div></div>");
            html.AppendLine($"<div class='summary-card card-skipped'><div class='card-number'>{result.SkippedCount}</div><div class='card-label'>Skipped</div></div>");
            html.AppendLine("</div>");

            // Details table
            if (result.Details != null && result.Details.Count > 0)
            {
                html.AppendLine("<h2>Import Details</h2>");
                html.AppendLine("<table>");
                html.AppendLine("<tr><th>Row</th><th>Object</th><th>Attribute</th><th>Value</th><th>Status</th><th>Notes</th></tr>");
                foreach (var detail in result.Details.OrderBy(d => d.RowNumber))
                {
                    var statusClass = detail.Status == ImportStatus.Success ? "status-success" :
                                     detail.Status == ImportStatus.Failed ? "status-failed" : "status-skipped";
                    html.AppendLine($"<tr>");
                    html.AppendLine($"<td>{detail.RowNumber}</td>");
                    html.AppendLine($"<td>{detail.ObjectName}</td>");
                    html.AppendLine($"<td>{detail.AttributeName}</td>");
                    html.AppendLine($"<td>{detail.AttributeValue}</td>");
                    html.AppendLine($"<td class='{statusClass}'>{detail.Status}</td>");
                    html.AppendLine($"<td>{detail.ErrorMessage ?? "-"}</td>");
                    html.AppendLine("</tr>");
                }
                html.AppendLine("</table>");
            }

            html.AppendLine("</div></body></html>");
            return html.ToString();
        }

        /// <summary>
        /// Saves the report to a file.
        /// </summary>
        public string SaveReport(ImportResult result, ImportConfiguration config, string outputPath = null)
        {
            outputPath ??= Path.Combine(
                Path.GetDirectoryName(config.ExcelFilePath) ?? Environment.CurrentDirectory,
                $"ImportReport_{DateTime.Now:yyyyMMdd_HHmmss}.html");

            var htmlReport = GenerateHtmlReport(result, config);
            File.WriteAllText(outputPath, htmlReport);

            _logger.LogInfo($"Report saved to: {outputPath}");
            return outputPath;
        }
    }
}
