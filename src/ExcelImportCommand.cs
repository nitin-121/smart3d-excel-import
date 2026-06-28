// ============================================================================
// Smart3D V14 - Excel Property Import Custom Command
// File: ExcelImportCommand.cs
// Description: Main command class registered with Smart3D command framework.
//              Reads Excel data and bulk-updates object properties.
// ============================================================================

using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.ServiceManager;
using Ingr.SP3D.Content.ServiceManager;
using Ingr.SP3D.UI;
using Serilog;
using Smart3D.ExcelImport.Core;

namespace Smart3D.ExcelImport
{
    /// <summary>
    /// Smart3D V14 Custom Command: Import Properties from Excel
    /// Mirrors Aveva E3D "Import Data From Excel" functionality
    /// </summary>
    [ComVisible(true)]
    [Guid("B2C3D4E5-F6A7-8901-BCDE-F12345678901")]
    [ClassInterface(ClassInterfaceType.None)]
    public class ExcelImportCommand : ICommand
    {
        private ILogger _logger;
        private ImportResult _lastResult;

        /// <summary>
        /// Command execution entry point called by Smart3D
        /// </summary>
        public void Execute()
        {
            InitializeLogger();

            _logger.Information("=== Smart3D Excel Property Import Started ===");

            try
            {
                // Step 1: Show file picker to select Excel file
                string excelPath = ShowFileDialog();
                if (string.IsNullOrEmpty(excelPath))
                {
                    _logger.Information("User cancelled file selection.");
                    return;
                }

                // Step 2: Validate file exists
                if (!File.Exists(excelPath))
                {
                    MessageBox.Show($"File not found: {excelPath}", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Step 3: Parse Excel data
                _logger.Information("Reading Excel file: {Path}", excelPath);
                var parser = new ExcelParser(_logger);
                var importRecords = parser.Parse(excelPath);

                if (importRecords.Count == 0)
                {
                    MessageBox.Show("No valid records found in Excel file.", "Warning",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Step 4: Confirm with user
                var confirmResult = MessageBox.Show(
                    $"Found {importRecords.Count} records to import.\n\nProceed with property updates?",
                    "Confirm Import",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (confirmResult != DialogResult.Yes)
                {
                    _logger.Information("User cancelled import.");
                    return;
                }

                // Step 5: Initialize Smart3D services
                var model = Smart3DApplication.ActiveModel;
                if (model == null)
                {
                    MessageBox.Show("No active model found. Please open a model first.",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Step 6: Execute import within transaction
                var engine = new ImportEngine(model, _logger);
                _lastResult = engine.ExecuteImport(importRecords);

                // Step 7: Show summary report
                ShowSummaryReport(_lastResult);

                _logger.Information("=== Import Complete: {Success}/{Total} succeeded ===",
                    _lastResult.SuccessCount, _lastResult.TotalCount);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Unhandled exception during import");
                MessageBox.Show(
                    $"Import failed with error:\n\n{ex.Message}\n\nSee log for details.",
                    "Import Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Log.CloseAndFlush();
            }
        }

        /// <summary>
        /// Required by ICommand interface — command name shown in Smart3D UI
        /// </summary>
        public string Name => "Import Properties from Excel";

        /// <summary>
        /// Required by ICommand interface — command description
        /// </summary>
        public string Description => "Bulk update object properties from Excel spreadsheet";

        /// <summary>
        /// Required by ICommand interface — command category
        /// </summary>
        public string Category => "Custom Commands";

        /// <summary>
        /// Required by ICommand interface — whether command is enabled
        /// </summary>
        public bool Enabled => Smart3DApplication.ActiveModel != null;

        /// <summary>
        /// Required by ICommand interface — command icon (optional)
        /// </summary>
        public System.Drawing.Image Icon => null;

        /// <summary>
        /// Shows the OpenFileDialog for Excel file selection
        /// </summary>
        private string ShowFileDialog()
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = "Select Excel File for Property Import";
                dialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
                dialog.FilterIndex = 1;
                dialog.RestoreDirectory = true;
                dialog.CheckFileExists = true;

                if (dialog.ShowDialog() == DialogResult.OK)
                    return dialog.FileName;
            }
            return null;
        }

        /// <summary>
        /// Displays the import summary report to the user
        /// </summary>
        private void ShowSummaryReport(ImportResult result)
        {
            var report = SummaryReportGenerator.Generate(result);

            using (var form = new SummaryReportForm(report))
            {
                form.ShowDialog();
            }
        }

        /// <summary>
        /// Initializes Serilog logger for import operations
        /// </summary>
        private void InitializeLogger()
        {
            string logPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Smart3D", "ExcelImport", "logs");

            Directory.CreateDirectory(logPath);

            string logFile = Path.Combine(logPath,
                $"import_{DateTime.Now:yyyyMMdd_HHmmss}.log");

            _logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.File(logFile,
                    outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff} [{Level:u3}] {Message:lj}{NewLine}{Exception}")
                .CreateLogger();
        }
    }
}
