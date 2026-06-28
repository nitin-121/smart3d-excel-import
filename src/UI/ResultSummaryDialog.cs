using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Smart3D.ExcelImport.Core;
using Smart3D.ExcelImport.Core;
using Smart3D.ExcelImport.Services;

namespace Smart3D.ExcelImport.UI
{
    /// <summary>
    /// Dialog showing the import results summary.
    /// Displays statistics, detailed results, and provides options to save reports.
    /// </summary>
    public partial class ResultSummaryDialog : Form
    {
        private readonly ImportResult _result;
        private readonly LoggingService _logger;
        private DataGridView dgvResults;
        private Button btnClose;
        private Button btnSaveReport;
        private Button btnViewLog;
        private Label lblSummary;

        public ResultSummaryDialog(ImportResult result, LoggingService logger)
        {
            _result = result;
            _logger = logger;
            InitializeComponent();
            PopulateResults();
        }

        private void InitializeComponent()
        {
            this.Text = "Smart3D V14 - Import Results";
            this.Size = new Size(900, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimizeBox = true;
            this.MaximizeBox = true;

            // Summary panel
            var panelTop = new Panel
            {
                Dock = DockStyle.Top,
                Height = 100,
                BackColor = Color.White,
                Padding = new Padding(20)
            };

            var successRate = _result.TotalRecordsFound > 0
                ? (double)_result.SuccessCount / _result.TotalRecordsFound * 100
                : 0;

            lblSummary = new Label
            {
                Text = $"Import Complete: {_result.SuccessCount} updated, " +
                       $"{_result.FailedCount} failed, {_result.SkippedCount} skipped " +
                       $"({successRate:F1}% success rate)",
                Font = new Font("Segoe UI", 14, FontStyle.Bold),
                ForeColor = _result.Success ? Color.DarkGreen : Color.DarkRed,
                Location = new Point(20, 35),
                AutoSize = true
            };

            panelTop.Controls.Add(lblSummary);

            // Data grid
            dgvResults = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None
            };

            dgvResults.Columns.Add("Row", "Row");
            dgvResults.Columns.Add("ObjectName", "Object Name");
            dgvResults.Columns.Add("Attribute", "Attribute");
            dgvResults.Columns.Add("Value", "Value");
            dgvResults.Columns.Add("Status", "Status");
            dgvResults.Columns.Add("Notes", "Notes");

            // Style the grid
            dgvResults.Columns["Row"].Width = 50;
            dgvResults.Columns["ObjectName"].Width = 150;
            dgvResults.Columns["Attribute"].Width = 120;
            dgvResults.Columns["Value"].Width = 120;
            dgvResults.Columns["Status"].Width = 100;
            dgvResults.Columns["Notes"].Width = 250;

            // Button panel
            var panelBottom = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 50,
                Padding = new Padding(10, 10, 10, 10)
            };

            btnSaveReport = new Button
            {
                Text = "Save Report",
                Location = new Point(10, 10),
                Size = new Size(120, 30)
            };
            btnSaveReport.Click += BtnSaveReport_Click;

            btnViewLog = new Button
            {
                Text = "View Log",
                Location = new Point(140, 10),
                Size = new Size(100, 30)
            };
            btnViewLog.Click += BtnViewLog_Click;

            btnClose = new Button
            {
                Text = "Close",
                DialogResult = DialogResult.OK,
                Location = new Point(760, 10),
                Size = new Size(100, 30)
            };

            panelBottom.Controls.AddRange(new Control[] { btnSaveReport, btnViewLog, btnClose });

            this.AcceptButton = btnClose;

            this.Controls.AddRange(new Control[] { dgvResults, panelBottom, panelTop });
        }

        private void PopulateResults()
        {
            if (_result.Details == null) return;

            foreach (var detail in _result.Details)
            {
                var rowIndex = dgvResults.Rows.Add(
                    detail.RowNumber,
                    detail.ObjectName,
                    detail.AttributeName,
                    detail.AttributeValue,
                    detail.Status.ToString(),
                    detail.ErrorMessage ?? ""
                );

                // Color-code by status
                var row = dgvResults.Rows[rowIndex];
                switch (detail.Status)
                {
                    case ImportStatus.Success:
                        row.DefaultCellStyle.ForeColor = Color.DarkGreen;
                        break;
                    case ImportStatus.Failed:
                        row.DefaultCellStyle.ForeColor = Color.DarkRed;
                        row.DefaultCellStyle.BackColor = Color.MistyRose;
                        break;
                    case ImportStatus.ObjectNotFound:
                    case ImportStatus.PropertyNotFound:
                        row.DefaultCellStyle.ForeColor = Color.DarkOrange;
                        break;
                    case ImportStatus.TypeMismatch:
                    case ImportStatus.Skipped:
                        row.DefaultCellStyle.ForeColor = Color.Gray;
                        break;
                }
            }
        }

        private void BtnSaveReport_Click(object sender, EventArgs e)
        {
            using (var saveDialog = new SaveFileDialog())
            {
                saveDialog.Filter = "HTML Report (*.html)|*.html|Text Report (*.txt)|*.txt";
                saveDialog.FilterIndex = 1;
                saveDialog.FileName = $"ImportReport_{DateTime.Now:yyyyMMdd_HHmmss}";

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var reportGenerator = new ReportGenerator(_logger);
                        string reportContent;

                        if (saveDialog.FileName.EndsWith(".txt"))
                        {
                            reportContent = reportGenerator.GenerateTextReport(_result, 
                                new ImportConfiguration { ExcelFilePath = "N/A" });
                        }
                        else
                        {
                            reportContent = reportGenerator.GenerateHtmlReport(_result, 
                                new ImportConfiguration { ExcelFilePath = "N/A" });
                        }

                        File.WriteAllText(saveDialog.FileName, reportContent);
                        MessageBox.Show("Report saved successfully.", "Save Report",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Failed to save report: {ex.Message}", "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void BtnViewLog_Click(object sender, EventArgs e)
        {
            var logContent = _logger.GetFullLog();
            using (var logForm = new Form())
            {
                logForm.Text = "Import Log";
                logForm.Size = new Size(800, 600);
                logForm.StartPosition = FormStartPosition.CenterParent;

                var textBox = new TextBox
                {
                    Dock = DockStyle.Fill,
                    Multiline = true,
                    ReadOnly = true,
                    ScrollBars = ScrollBars.Both,
                    Font = new Font("Consolas", 9),
                    Text = logContent
                };

                logForm.Controls.Add(textBox);
                logForm.ShowDialog();
            }
        }
    }
}
