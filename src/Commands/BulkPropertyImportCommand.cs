using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Smart3D.ExcelImport.Core;
using Smart3D.ExcelImport.Models;
using Smart3D.ExcelImport.Services;
using Smart3D.ExcelImport.UI;

namespace Smart3D.ExcelImport.Commands
{
    /// <summary>
    /// Main command for bulk property import from Excel to Smart3D V14.
    /// Orchestrates the entire import workflow:
    /// 1. Show file selection dialog
    /// 2. Read and validate Excel data
    /// 3. Find objects in model
    /// 4. Update properties within a transaction
    /// 5. Generate summary report
    /// </summary>
    public sealed class BulkPropertyImportCommand : CommandBase
    {
        private readonly LoggingService _logger;
        private ImportConfiguration _config;

        public BulkPropertyImportCommand()
        {
            var logPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Smart3D.ExcelImport",
                $"ImportLog_{DateTime.Now:yyyyMMdd_HHmmss}.log");

            Directory.CreateDirectory(Path.GetDirectoryName(logPath));
            _logger = new LoggingService(logPath);
        }

        public override void Execute()
        {
            _logger.LogInfo("=== Smart3D V14 Bulk Property Import Command Started ===");

            try
            {
                // Step 1: Show configuration dialog
                _config = ShowImportDialog();
                if (_config == null)
                {
                    _logger.LogInfo("User cancelled the import operation.");
                    return;
                }

                _logger.LogInfo($"Import configuration: Excel={_config.ExcelFilePath}, " +
                    $"StrictType={_config.StrictTypeMatching}, SkipDuplicates={_config.SkipDuplicates}");

                // Step 2: Read Excel data
                List<ImportRecord> records;
                using (var reader = new ExcelDataReader(_config.ExcelFilePath, _logger))
                {
                    var readResult = reader.ReadRecordsWithValidation();
                    if (!readResult.Success)
                    {
                        MessageBox.Show(
                            $"Failed to read Excel file:\n{readResult.ErrorMessage}",
                            "Import Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                        return;
                    }
                    records = readResult.Records;
                }

                if (records.Count == 0)
                {
                    MessageBox.Show(
                        "No valid records found in the Excel file.",
                        "Import Warning",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // Step 3: Confirm with user
                var confirmResult = MessageBox.Show(
                    $"Found {records.Count} records to import.\n\n" +
                    $"File: {_config.ExcelFilePath}\n\n" +
                    "Do you want to proceed with the import?",
                    "Confirm Import",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (confirmResult != DialogResult.Yes)
                {
                    _logger.LogInfo("User cancelled import at confirmation.");
                    return;
                }

                // Step 4: Process import
                var importResult = ProcessImport(records);

                // Step 5: Show results
                ShowResults(importResult);
            }
            catch (Exception ex)
            {
                _logger.LogError("Unhandled exception in import command", ex);
                MessageBox.Show(
                    $"An error occurred during import:\n\n{ex.Message}",
                    "Import Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                _logger.LogInfo("=== Import Command Completed ===");
            }
        }

        private ImportConfiguration ShowImportDialog()
        {
            using (var dialog = new ImportDialog(_logger))
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    return dialog.Configuration;
                }
            }
            return null;
        }

        private ImportResult ProcessImport(List<ImportRecord> records)
        {
            var result = new ImportResult { Records = records };

            // Initialize components
            var transactionManager = new TransactionManager(Application, _logger);
            var objectFinder = new ObjectFinder(Application, _logger);
            var updateEngine = new PropertyUpdateEngine(Application, _logger, objectFinder, transactionManager);

            // Show progress dialog
            using (var progressDialog = new ImportProgressDialog())
            {
                progressDialog.Show();

                // Process in background
                var bgWorker = new System.ComponentModel.BackgroundWorker();
                bgWorker.WorkerReportsProgress = true;
                bgWorker.WorkerSupportsCancellation = true;

                bgWorker.DoWork += (sender, e) =>
                {
                    result = updateEngine.ProcessRecords(records, _config);
                };

                bgWorker.RunWorkerCompleted += (sender, e) =>
                {
                    progressDialog.Close();
                };

                bgWorker.RunWorkerAsync();

                // Keep UI responsive
                while (bgWorker.IsBusy)
                {
                    Application.DoEvents();
                    System.Threading.Thread.Sleep(50);
                }
            }

            // Generate report if configured
            if (_config.GenerateReport)
            {
                var reportGenerator = new ReportGenerator(_logger);
                var reportPath = reportGenerator.SaveReport(result, _config, _config.ReportOutputPath);
                result.ReportPath = reportPath;
            }

            return result;
        }

        private void ShowResults(ImportResult result)
        {
            using (var summaryDialog = new ResultSummaryDialog(result, _logger))
            {
                summaryDialog.ShowDialog();
            }
        }
    }
}
