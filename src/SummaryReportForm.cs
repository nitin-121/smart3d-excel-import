// ============================================================================
// Smart3D V14 - Excel Property Import
// File: SummaryReportForm.cs
// Description: Windows Forms dialog showing import results with export options
// ============================================================================

using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Smart3D.ExcelImport.Core;

namespace Smart3D.ExcelImport
{
    /// <summary>
    /// Windows Form displaying import summary report with export capabilities
    /// </summary>
    public class SummaryReportForm : Form
    {
        private readonly ImportResult _result;
        private DataGridView dgvResults;
        private Label lblSummary;
        private Button btnExportCsv;
        private Button btnCopyClipboard;
        private Button btnOk;

        public SummaryReportForm(ImportResult result)
        {
            _result = result ?? throw new ArgumentNullException(nameof(result));
            InitializeComponents();
            LoadResults();
        }

        private void InitializeComponents()
        {
            // Form settings
            this.Text = "Smart3D — Excel Import Summary";
            this.Size = new Size(900, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MinimumSize = new Size(700, 400);

            // Summary label
            lblSummary = new Label
            {
                Location = new Point(12, 12),
                Size = new Size(860, 60),
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                Text = GetSummaryText()
            };

            // DataGridView
            dgvResults = new DataGridView
            {
                Location = new Point(12, 80),
                Size = new Size(860, 420),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                RowHeadersVisible = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.Fixed3D
            };

            // Define columns
            dgvResults.Columns.Add("Row", "Row");
            dgvResults.Columns.Add("Object", "Object Name");
            dgvResults.Columns.Add("Attribute", "Attribute");
            dgvResults.Columns.Add("OldValue", "Old Value");
            dgvResults.Columns.Add("NewValue", "New Value");
            dgvResults.Columns.Add("Status", "Status");
            dgvResults.Columns.Add("Error", "Error Message");

            // Color coding
            dgvResults.CellFormatting += (s, e) =>
            {
                if e.ColumnIndex == 5 // Status column
                {
                    var row = dgvResults.Rows[e.RowIndex];
                    string status = row.Cells[5].Value?.ToString();
                    if (status == "SUCCESS")
                        e.CellStyle.ForeColor = Color.Green;
                    else
                        e.CellStyle.ForeColor = Color.Red;
                }
            };

            // Export CSV button
            btnExportCsv = new Button
            {
                Text = "📄 Export CSV",
                Location = new Point(12, 510),
                Size = new Size(130, 35),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            btnExportCsv.Click += BtnExportCsv_Click;

            // Copy to clipboard button
            btnCopyClipboard = new Button
            {
                Text = "📋 Copy Report",
                Location = new Point(155, 510),
                Size = new Size(130, 35),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            btnCopyClipboard.Click += BtnCopyClipboard_Click;

            // OK button
            btnOk = new Button
            {
                Text = "OK",
                Location = new Point(797, 510),
                Size = new Size(75, 35),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                DialogResult = DialogResult.OK
            };

            // Add controls
            this.Controls.Add(lblSummary);
            this.Controls.Add(dgvResults);
            this.Controls.Add(btnExportCsv);
            this.Controls.Add(btnCopyClipboard);
            this.Controls.Add(btnOk);

            this.AcceptButton = btnOk;
        }

        private void LoadResults()
        {
            dgvResults.Rows.Clear();

            foreach (var r in _result.Results)
            {
                dgvResults.Rows.Add(
                    r.RowNumber,
                    r.ObjectName,
                    r.AttributeName,
                    r.OldValue,
                    r.NewValue,
                    r.Success ? "SUCCESS" : "FAILED",
                    r.ErrorMessage
                );
            }

            // Auto-size
            dgvResults.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
        }

        private string GetSummaryText()
        {
            double rate = _result.TotalCount > 0
                ? (double)_result.SuccessCount / _result.TotalCount * 100
                : 0;

            return $"Import Complete: {_result.SuccessCount}/{_result.TotalCount} succeeded " +
                   $"({rate:F1}%) | Failed: {_result.FailureCount} | " +
                   $"Duration: {_result.Duration.TotalSeconds:F1}s";
        }

        private void BtnExportCsv_Click(object sender, EventArgs e)
        {
            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*";
                dialog.FileName = $"ImportReport_{DateTime.Now:yyyyMMdd_HHmmss}.csv";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        SummaryReportGenerator.SaveCsvReport(_result, dialog.FileName);
                        MessageBox.Show($"Report saved to:\n{dialog.FileName}",
                            "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Export failed:\n{ex.Message}",
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void BtnCopyClipboard_Click(object sender, EventArgs e)
        {
            try
            {
                string report = SummaryReportGenerator.GenerateTextReport(_result);
                Clipboard.SetText(report);
                MessageBox.Show("Report copied to clipboard!", "Copied",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Copy failed:\n{ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
