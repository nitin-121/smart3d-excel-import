using System;
using System.Drawing;
using System.Windows.Forms;

namespace Smart3D.ExcelImport.UI
{
    /// <summary>
    /// Progress dialog shown during the import operation.
    /// Provides visual feedback on the import progress.
    /// </summary>
    public partial class ImportProgressDialog : Form
    {
        private ProgressBar progressBar;
        private Label lblStatus;
        private Label lblProgress;
        private Button btnCancel;

        public ImportProgressDialog()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "Smart3D V14 - Importing Properties...";
            this.Size = new Size(500, 180);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ControlBox = false;

            lblStatus = new Label
            {
                Text = "Preparing import...",
                Location = new Point(20, 20),
                Size = new Size(450, 20),
                Font = new Font("Segoe UI", 10)
            };

            progressBar = new ProgressBar
            {
                Location = new Point(20, 50),
                Size = new Size(450, 25),
                Style = ProgressBarStyle.Marquee,
                MarqueeAnimationSpeed = 30
            };

            lblProgress = new Label
            {
                Text = "",
                Location = new Point(20, 85),
                Size = new Size(450, 20),
                ForeColor = Color.Gray
            };

            btnCancel = new Button
            {
                Text = "Cancel",
                Location = new Point(380, 110),
                Size = new Size(90, 28),
                DialogResult = DialogResult.Cancel
            };

            this.CancelButton = btnCancel;

            this.Controls.AddRange(new Control[] {
                lblStatus, progressBar, lblProgress, btnCancel
            });
        }

        /// <summary>
        /// Updates the progress display.
        /// </summary>
        public void UpdateProgress(int current, int total, string objectName)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new Action(() => UpdateProgress(current, total, objectName)));
                return;
            }

            var percentage = (int)((double)current / total * 100);
            lblStatus.Text = $"Processing record {current} of {total} ({percentage}%)";
            lblProgress.Text = $"Current: {objectName}";
            this.Refresh();
        }

        /// <summary>
        /// Updates the status message.
        /// </summary>
        public void UpdateStatus(string message)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new Action(() => UpdateStatus(message)));
                return;
            }

            lblStatus.Text = message;
            this.Refresh();
        }
    }
}
