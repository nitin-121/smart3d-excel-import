using System;
using System.IO;
using System.Windows.Forms;
using Smart3D.ExcelImport.Core;
using Smart3D.ExcelImport.Models;
using Smart3D.ExcelImport.Services;

namespace Smart3D.ExcelImport.UI
{
    /// <summary>
    /// Dialog for configuring the import operation.
    /// Allows the user to select an Excel file and set import options.
    /// </summary>
    public partial class ImportDialog : Form
    {
        private readonly LoggingService _logger;
        private TextBox txtFilePath;
        private Button btnBrowse;
        private Button btnOK;
        private Button btnCancel;
        private CheckBox chkStrictTypeMatching;
        private CheckBox chkSkipDuplicates;
        private CheckBox chkGenerateReport;
        private CheckBox chkValidateProperties;
        private NumericUpDown nudMaxErrors;
        private Label lblStatus;
        private Label lblPreview;

        public ImportConfiguration Configuration { get; private set; }

        public ImportDialog(LoggingService logger)
        {
            _logger = logger;
            InitializeComponent();
            UpdateStatusReady();
        }

        private void InitializeComponent()
        {
            this.Text = "Smart3D V14 - Bulk Property Import from Excel";
            this.Size = new System.Drawing.Size(650, 520);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            var lblTitle = new Label
            {
                Text = "Bulk Property Import from Excel",
                Font = new System.Drawing.Font("Segoe UI", 14, System.Drawing.FontStyle.Bold),
                Location = new System.Drawing.Point(20, 15),
                AutoSize = true
            };

            var lblDescription = new Label
            {
                Text = "Import object properties (PipeRun, Pipeline, Equipment) from an Excel spreadsheet.\n" +
                       "Expected columns: ObjectName, ObjectType, AttributeName, AttributeValue",
                Location = new System.Drawing.Point(20, 50),
                Size = new System.Drawing.Size(600, 40),
                ForeColor = System.Drawing.Color.Gray
            };

            var lblFile = new Label
            {
                Text = "Excel File:",
                Location = new System.Drawing.Point(20, 100),
                AutoSize = true
            };

            txtFilePath = new TextBox
            {
                Location = new System.Drawing.Point(20, 120),
                Size = new System.Drawing.Size(480, 25),
                ReadOnly = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };

            btnBrowse = new Button
            {
                Text = "Browse...",
                Location = new System.Drawing.Point(510, 118),
                Size = new System.Drawing.Size(100, 28)
            };
            btnBrowse.Click += BtnBrowse_Click;

            // Options Group
            var grpOptions = new GroupBox
            {
                Text = "Import Options",
                Location = new System.Drawing.Point(20, 160),
                Size = new System.Drawing.Size(600, 180)
            };

            chkStrictTypeMatching = new CheckBox
            {
                Text = "Require exact type matching (PipeRun, Pipeline, Equipment)",
                Location = new System.Drawing.Point(15, 25),
                AutoSize = true,
                Checked = false
            };

            chkSkipDuplicates = new CheckBox
            {
                Text = "Skip duplicate entries (same object + attribute)",
                Location = new System.Drawing.Point(15, 55),
                AutoSize = true,
                Checked = true
            };

            chkValidateProperties = new CheckBox
            {
                Text = "Validate property names before setting",
                Location = new System.Drawing.Point(15, 85),
                AutoSize = true,
                Checked = true
            };

            chkGenerateReport = new CheckBox
            {
                Text = "Generate summary report after import",
                Location = new System.Drawing.Point(15, 115),
                AutoSize = true,
                Checked = true
            };

            var lblMaxErrors = new Label
            {
                Text = "Max errors before abort:",
                Location = new System.Drawing.Point(15, 148),
                AutoSize = true
            };

            nudMaxErrors = new NumericUpDown
            {
                Location = new System.Drawing.Point(200, 145),
                Size = new System.Drawing.Size(80, 25),
                Minimum = 1,
                Maximum = 10000,
                Value = 100
            };

            grpOptions.Controls.AddRange(new Control[] {
                chkStrictTypeMatching, chkSkipDuplicates, chkValidateProperties,
                chkGenerateReport, lblMaxErrors, nudMaxErrors
            });

            // Preview label
            lblPreview = new Label
            {
                Text = "",
                Location = new System.Drawing.Point(20, 350),
                Size = new System.Drawing.Size(600, 40),
                ForeColor = System.Drawing.Color.DarkBlue
            };

            // Status label
            lblStatus = new Label
            {
                Text = "Ready",
                Location = new System.Drawing.Point(20, 395),
                Size = new System.Drawing.Size(600, 20),
                ForeColor = System.Drawing.Color.Gray
            };

            // Buttons
            btnOK = new Button
            {
                Text = "Import",
                Location = new System.Drawing.Point(420, 430),
                Size = new System.Drawing.Size(90, 30),
                Enabled = false
            };
            btnOK.Click += BtnOK_Click;

            btnCancel = new Button
            {
                Text = "Cancel",
                Location = new System.Drawing.Point(520, 430),
                Size = new System.Drawing.Size(90, 30),
                DialogResult = DialogResult.Cancel
            };

            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;

            this.Controls.AddRange(new Control[] {
                lblTitle, lblDescription, lblFile, txtFilePath, btnBrowse,
                grpOptions, lblPreview, lblStatus, btnOK, btnCancel
            });
        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
                openFileDialog.FilterIndex = 1;
                openFileDialog.Title = "Select Excel Import File";
                openFileDialog.CheckFileExists = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtFilePath.Text = openFileDialog.FileName;
                    PreviewFile(openFileDialog.FileName);
                }
            }
        }

        private void PreviewFile(string filePath)
        {
            try
            {
                UpdateStatus("Reading file...");

                using (var reader = new Core.ExcelDataReader(filePath, _logger))
                {
                    var result = reader.ReadRecordsWithValidation();
                    if (result.Success)
                    {
                        lblPreview.Text = $"Found {result.Records.Count} valid records to import.";
                        lblStatus.Text = "File loaded successfully. Click Import to proceed.";
                        btnOK.Enabled = true;
                    }
                    else
                    {
                        lblPreview.Text = $"Error: {result.ErrorMessage}";
                        lblStatus.Text = "Failed to read file.";
                        btnOK.Enabled = false;
                    }
                }
            }
            catch (Exception ex)
            {
                lblPreview.Text = $"Error: {ex.Message}";
                lblStatus.Text = "Failed to read file.";
                btnOK.Enabled = false;
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            Configuration = new ImportConfiguration
            {
                ExcelFilePath = txtFilePath.Text,
                StrictTypeMatching = chkStrictTypeMatching.Checked,
                SkipDuplicates = chkSkipDuplicates.Checked,
                ValidateProperties = chkValidateProperties.Checked,
                GenerateReport = chkGenerateReport.Checked,
                MaxErrors = (int)nudMaxErrors.Checked,
                ReportOutputPath = Path.ChangeExtension(txtFilePath.Text, ".report.html")
            };

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void UpdateStatus(string message)
        {
            lblStatus.Text = message;
            this.Refresh();
        }

        private void UpdateStatusReady()
        {
            lblStatus.Text = "Ready - Select an Excel file to begin.";
        }
    }
}
