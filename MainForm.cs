using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Drawing;

namespace AttendanceApp
{
    public partial class MainForm : Form
    {
        private List<string> uploadedFiles = new List<string>();
        private const int BUTTON_PADDING = 10;

        public MainForm()
        {
            try
            {
                InitializeComponent();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error initializing form: {ex.Message}\n\n{ex.StackTrace}", "Initialization Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }

        private void InitializeComponent()
        {
            this.Text = "Attendance Summarizer";
            this.Size = new System.Drawing.Size(1000, 900);
            this.MinimumSize = new System.Drawing.Size(800, 700);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = new System.Drawing.Font("Segoe UI", 10);
            this.Padding = new Padding(10);

            // Main panel with automatic scroll
            var mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10),
                AutoScroll = true
            };
            this.Controls.Add(mainPanel);
            this.mainPanel = mainPanel;

            // Results section (add first so it appears last)
            var resultsGroupBox = new GroupBox
            {
                Text = "Preview",
                Dock = DockStyle.Top,
                Height = 280,
                Padding = new Padding(10),
                Margin = new Padding(0, 0, 0, 10)
            };
            mainPanel.Controls.Add(resultsGroupBox);

            // Settings section
            var settingsGroupBox = new GroupBox
            {
                Text = "Settings",
                Dock = DockStyle.Top,
                Height = 120,
                Padding = new Padding(10),
                Margin = new Padding(0, 0, 0, 10)
            };
            mainPanel.Controls.Add(settingsGroupBox);

            // Upload section
            var uploadGroupBox = new GroupBox
            {
                Text = "Upload Files",
                Dock = DockStyle.Top,
                Height = 300,
                Padding = new Padding(10),
                Margin = new Padding(0, 0, 0, 10)
            };
            mainPanel.Controls.Add(uploadGroupBox);

            // Title (add last so it appears first)
            var titleLabel = new Label
            {
                Text = "Attendance Summarizer",
                Font = new System.Drawing.Font("Segoe UI", 16, System.Drawing.FontStyle.Bold),
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 10)
            };
            mainPanel.Controls.Add(titleLabel);

            // Drag and drop zone panel
            var dragDropPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = System.Drawing.Color.FromArgb(240, 248, 255),
                AllowDrop = true,
                AutoScroll = true,
                BorderStyle = BorderStyle.FixedSingle,
                Name = "dragDropPanel",
                Margin = new Padding(0)
            };
            dragDropPanel.Paint += (s, e) =>
            {
                // Draw dashed border
                var pen = new System.Drawing.Pen(System.Drawing.Color.FromArgb(100, 150, 220), 2)
                {
                    DashStyle = System.Drawing.Drawing2D.DashStyle.Dash
                };
                e.Graphics.DrawRectangle(pen, 1, 1, dragDropPanel.Width - 3, dragDropPanel.Height - 3);
                pen.Dispose();
            };
            dragDropPanel.DragEnter += (s, e) =>
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    e.Effect = DragDropEffects.Copy;
                    dragDropPanel.BackColor = System.Drawing.Color.FromArgb(220, 240, 255);
                    dragDropPanel.Invalidate();
                }
            };
            dragDropPanel.DragLeave += (s, e) =>
            {
                dragDropPanel.BackColor = System.Drawing.Color.FromArgb(240, 248, 255);
                dragDropPanel.Invalidate();
            };
            dragDropPanel.DragDrop += (s, e) =>
            {
                dragDropPanel.BackColor = System.Drawing.Color.FromArgb(240, 248, 255);
                dragDropPanel.Invalidate();
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                    var validFiles = files.Where(f => f.EndsWith(".csv", StringComparison.OrdinalIgnoreCase)
                        || f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)).ToList();
                    if (validFiles.Count > 0)
                    {
                        uploadedFiles = validFiles;
                        UpdateFileListBox();
                    }
                    else
                    {
                        MessageBox.Show("Please drop only CSV or XLSX files.", "Invalid Files", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            };
            uploadGroupBox.Controls.Add(dragDropPanel);
            this.dragDropPanel = dragDropPanel;

            // Header panel for buttons and label
            var headerPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 50,
                Padding = new Padding(5),
                Margin = new Padding(0),
                BackColor = System.Drawing.Color.FromArgb(240, 248, 255)
            };
            dragDropPanel.Controls.Add(headerPanel);

            // Drag and drop instruction label
            var dragDropLabel = new Label
            {
                Text = "📁 Drag and drop files here",
                Dock = DockStyle.Fill,
                Font = new System.Drawing.Font("Segoe UI", 10, System.Drawing.FontStyle.Regular),
                ForeColor = System.Drawing.Color.FromArgb(100, 100, 100),
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Margin = new Padding(5, 0, 0, 0)
            };
            headerPanel.Controls.Add(dragDropLabel);

            // Button panel for organizing buttons
            var buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                Width = 230,
                Height = 50,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                Padding = new Padding(0),
                Margin = new Padding(0),
                BackColor = System.Drawing.Color.FromArgb(240, 248, 255)
            };
            headerPanel.Controls.Add(buttonPanel);

            // Clear button inside drop zone
            var clearButton = new Button
            {
                Text = "Clear All",
                Width = 100,
                Height = 40,
                BackColor = System.Drawing.Color.FromArgb(220, 38, 38),
                ForeColor = System.Drawing.Color.White,
                Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold),
                Margin = new Padding(5),
                FlatStyle = FlatStyle.Flat
            };
            clearButton.Click += (s, e) => ClearFiles();
            buttonPanel.Controls.Add(clearButton);

            // Browse button inside drop zone
            var uploadButton = new Button
            {
                Text = "Browse",
                Width = 100,
                Height = 40,
                BackColor = System.Drawing.Color.FromArgb(59, 130, 246),
                ForeColor = System.Drawing.Color.White,
                Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold),
                Margin = new Padding(5),
                FlatStyle = FlatStyle.Flat
            };
            uploadButton.Click += (s, e) => UploadFiles();
            buttonPanel.Controls.Add(uploadButton);

            // File chips panel
            var chipsPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                BackColor = System.Drawing.Color.Transparent,
                AutoScroll = true,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                Name = "chipsPanel",
                Margin = new Padding(0),
                Padding = new Padding(5)
            };
            dragDropPanel.Controls.Add(chipsPanel);
            this.chipsPanel = chipsPanel;

            // Create a responsive flow panel for settings controls
            var settingsPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                AutoScroll = false,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };
            settingsGroupBox.Controls.Add(settingsPanel);

            // Month section
            var monthLabel = new Label
            {
                Text = "Month:",
                AutoSize = true,
                Font = new System.Drawing.Font("Segoe UI", 9),
                Margin = new Padding(0, 5, 5, 5)
            };
            settingsPanel.Controls.Add(monthLabel);

            var monthComboBox = new ComboBox
            {
                Width = 180,
                Height = 30,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Margin = new Padding(0, 0, 20, 0)
            };
            for (int i = 1; i <= 12; i++)
                monthComboBox.Items.Add(CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(i));
            monthComboBox.SelectedIndex = DateTime.Now.Month - 1;
            settingsPanel.Controls.Add(monthComboBox);
            this.monthComboBox = monthComboBox;

            // Year section
            var yearLabel = new Label
            {
                Text = "Year:",
                AutoSize = true,
                Font = new System.Drawing.Font("Segoe UI", 9),
                Margin = new Padding(0, 5, 5, 5)
            };
            settingsPanel.Controls.Add(yearLabel);

            var yearInput = new NumericUpDown
            {
                Width = 100,
                Height = 30,
                Minimum = 2000,
                Maximum = 2100,
                Value = DateTime.Now.Year,
                Margin = new Padding(0, 0, 20, 0)
            };
            settingsPanel.Controls.Add(yearInput);
            this.yearInput = yearInput;

            // Holiday section
            var holidayLabel = new Label
            {
                Text = "Holiday Count:",
                AutoSize = true,
                Font = new System.Drawing.Font("Segoe UI", 9),
                Margin = new Padding(0, 5, 5, 5)
            };
            settingsPanel.Controls.Add(holidayLabel);

            var holidayInput = new NumericUpDown
            {
                Width = 100,
                Height = 30,
                Minimum = 0,
                Maximum = 31,
                Value = 0,
                Margin = new Padding(0, 0, 0, 0)
            };
            settingsPanel.Controls.Add(holidayInput);
            this.holidayInput = holidayInput;

            // Add results controls to the resultsGroupBox
            var dataGridView = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
                BackgroundColor = System.Drawing.Color.White,
                Margin = new Padding(0),
                AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = System.Drawing.Color.FromArgb(240, 240, 240)
                }
            };
            resultsGroupBox.Controls.Add(dataGridView);
            this.dataGridView = dataGridView;

            // Loading spinner panel (overlay)
            var spinnerPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Visible = false,
                BackColor = System.Drawing.Color.FromArgb(240, 248, 255),
                Name = "spinnerPanel"
            };
            resultsGroupBox.Controls.Add(spinnerPanel);
            this.spinnerPanel = spinnerPanel;

            // Loading spinner label (hidden by default)
            var spinnerLabel = new Label
            {
                Text = "⠋",
                Font = new System.Drawing.Font("Segoe UI", 48, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(59, 130, 246),
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill,
                Visible = false,
                BackColor = System.Drawing.Color.Transparent,
                Name = "spinnerLabel"
            };
            spinnerPanel.Controls.Add(spinnerLabel);
            this.spinnerLabel = spinnerLabel;

            // Action buttons
            var actionPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 60,
                Padding = new Padding(10),
                Margin = new Padding(0)
            };
            mainPanel.Controls.Add(actionPanel);

            var analyzeButton = new Button
            {
                Text = "Analyze & Generate",
                Dock = DockStyle.Left,
                Width = 150,
                BackColor = System.Drawing.Color.FromArgb(59, 130, 246),
                ForeColor = System.Drawing.Color.White,
                Font = new System.Drawing.Font("Segoe UI", 10, System.Drawing.FontStyle.Bold),
                Margin = new Padding(5)
            };
            analyzeButton.Click += (s, e) => AnalyzeFiles();
            actionPanel.Controls.Add(analyzeButton);

            var downloadButton = new Button
            {
                Text = "Download Excel",
                Dock = DockStyle.Right,
                Width = 150,
                BackColor = System.Drawing.Color.FromArgb(34, 197, 94),
                ForeColor = System.Drawing.Color.White,
                Font = new System.Drawing.Font("Segoe UI", 10, System.Drawing.FontStyle.Bold),
                Enabled = false,
                Margin = new Padding(5)
            };
            downloadButton.Click += (s, e) => DownloadExcel();
            actionPanel.Controls.Add(downloadButton);
            this.downloadButton = downloadButton;

            var exitButton = new Button
            {
                Text = "Exit",
                Dock = DockStyle.Right,
                Width = 120,
                BackColor = System.Drawing.Color.FromArgb(229, 231, 235),
                Font = new System.Drawing.Font("Segoe UI", 10),
                Margin = new Padding(5)
            };
            exitButton.Click += (s, e) => this.Close();
            actionPanel.Controls.Add(exitButton);

            // Progress bar
            var progressBar = new ProgressBar
            {
                Dock = DockStyle.Bottom,
                Height = 3,
                Style = ProgressBarStyle.Marquee,
                Visible = false,
                Margin = new Padding(0)
            };
            mainPanel.Controls.Add(progressBar);
            this.progressBar = progressBar;

            // Status bar
            var statusBar = new StatusStrip
            {
                Dock = DockStyle.Bottom,
                SizingGrip = false,
                BackColor = System.Drawing.Color.FromArgb(240, 240, 240),
                ForeColor = System.Drawing.Color.FromArgb(64, 64, 64)
            };
            this.Controls.Add(statusBar);

            var statusLabel = new ToolStripStatusLabel
            {
                Text = "Ready",
                AutoSize = true,
                Font = new System.Drawing.Font("Segoe UI", 9),
                ForeColor = System.Drawing.Color.Green,
                Spring = true,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            };
            statusBar.Items.Add(statusLabel);
            this.statusLabel = statusLabel;

            // Info button in action panel - simple text button
            var infoButton = new Button
            {
                Size = new System.Drawing.Size(45, 40),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Dock = DockStyle.Left,
                Margin = new Padding(5),
                Text = "ⓘ",
                Font = new System.Drawing.Font("Segoe UI", 14, System.Drawing.FontStyle.Bold)
            };
            infoButton.FlatAppearance.BorderSize = 0;

            actionPanel.Controls.Add(infoButton);
            this.infoButton = infoButton;

            // Info card panel (hidden by default, wider for side-by-side layout)
            var infoCardPanel = new Panel
            {
                Width = 400,
                Height = 0,
                BackColor = System.Drawing.Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Location = new System.Drawing.Point(10, 65),
                Anchor = AnchorStyles.Top | AnchorStyles.Left,
                Visible = true,
                Tag = "infoCard",
                AutoScroll = false
            };
            infoCardPanel.Paint += (s, e) =>
            {
                e.Graphics.Clear(System.Drawing.Color.White);
                var border = new Pen(System.Drawing.Color.FromArgb(59, 130, 246), 2);
                e.Graphics.DrawRectangle(border, 0, 0, infoCardPanel.Width - 1, infoCardPanel.Height - 1);
                border.Dispose();
            };
            this.Controls.Add(infoCardPanel);
            infoCardPanel.BringToFront();
            this.infoCardPanel = infoCardPanel;

            // Register click handler for info button
            infoButton.Click += (s, e) => ToggleInfoCard();
            
            // Initialize timer for info card animation
            infoCardTimer = new Timer();
            infoCardTimer.Interval = 20;
            infoCardTimer.Tick += InfoCardTimer_Tick;

            // Initialize spinner timer
            spinnerTimer = new Timer();
            spinnerTimer.Interval = 80;
            spinnerTimer.Tick += SpinnerTimer_Tick;

            // Add key down handler to close info card with Escape
            this.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Escape && infoCardExpanded)
                {
                    infoCardExpanded = false;
                    if (infoCardTimer != null)
                    {
                        infoCardTimer.Start();
                    }
                    e.Handled = true;
                }
            };

            // Add click handlers to close info card
            AddClickHandlerRecursive(this);
        }

        private void AddClickHandlerRecursive(Control parent)
        {
            foreach (Control control in parent.Controls)
            {
                // Skip the info card and info button themselves
                if (control == infoCardPanel || control == infoButton)
                {
                    continue;
                }

                control.Click += (s, e) =>
                {
                    if (infoCardExpanded)
                    {
                        infoCardExpanded = false;
                        if (infoCardTimer != null)
                        {
                            infoCardTimer.Start();
                        }
                    }
                };

                // Recursively add handlers to child controls
                if (control.HasChildren)
                {
                    AddClickHandlerRecursive(control);
                }
            }
        }

        private Panel mainPanel;
        private Panel dragDropPanel;
        private FlowLayoutPanel chipsPanel;
        private ComboBox monthComboBox;
        private NumericUpDown yearInput;
        private NumericUpDown holidayInput;
        private ToolStripStatusLabel statusLabel;
        private ProgressBar progressBar;
        private DataGridView dataGridView;
        private Button downloadButton;
        private byte[] excelBytes;
        private Panel infoCardPanel;
        private Button infoButton;
        private Timer infoCardTimer;
        private bool infoCardExpanded = false;
        private const int INFO_CARD_MAX_HEIGHT = 350;
        private Panel spinnerPanel;
        private Label spinnerLabel;
        private Timer spinnerTimer;
        private int spinnerIndex = 0;
        private readonly string[] spinnerFrames = { "⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏" };

        private void UpdateFileListBox()
        {
            chipsPanel.Controls.Clear();
            foreach (var file in uploadedFiles)
            {
                CreateFileChip(file);
            }
        }

        private void CreateFileChip(string filePath)
        {
            var chipPanel = new Panel
            {
                Size = new System.Drawing.Size(220, 38),
                BackColor = System.Drawing.Color.FromArgb(59, 130, 246),
                Margin = new Padding(5, 5, 5, 5),
                BorderStyle = BorderStyle.None
            };

            var fileNameLabel = new Label
            {
                Text = Path.GetFileName(filePath),
                Location = new System.Drawing.Point(8, 8),
                Size = new System.Drawing.Size(175, 22),
                ForeColor = System.Drawing.Color.White,
                Font = new System.Drawing.Font("Segoe UI", 9),
                AutoEllipsis = true,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Tag = filePath
            };
            chipPanel.Controls.Add(fileNameLabel);

            var closeButton = new Button
            {
                Text = "✕",
                Location = new System.Drawing.Point(191, 7),
                Size = new System.Drawing.Size(24, 24),
                BackColor = System.Drawing.Color.FromArgb(59, 130, 246),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new System.Drawing.Font("Segoe UI", 10, System.Drawing.FontStyle.Bold),
                Cursor = Cursors.Hand,
                Tag = filePath
            };
            closeButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(59, 130, 246);
            closeButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(191, 0, 0);
            closeButton.Click += (s, e) => RemoveFile(filePath);
            chipPanel.Controls.Add(closeButton);

            chipsPanel.Controls.Add(chipPanel);
        }

        private void RemoveFile(string filePath)
        {
            uploadedFiles.Remove(filePath);
            UpdateFileListBox();
        }

        private void ClearFiles()
        {
            if (uploadedFiles.Count == 0)
                return;

            var result = MessageBox.Show(
                "Are you sure you want to clear all uploaded files?",
                "Clear All Files",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                uploadedFiles.Clear();
                UpdateFileListBox();
            }
        }

        private void UploadFiles()
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "CSV and Excel files|*.csv;*.xlsx|CSV files|*.csv|Excel files|*.xlsx";
                dialog.Multiselect = true;
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    uploadedFiles = dialog.FileNames.ToList();
                    UpdateFileListBox();
                }
            }
        }

        private void AnalyzeFiles()
        {
            if (uploadedFiles.Count == 0)
            {
                MessageBox.Show("Please upload at least one file.", "No Files", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            progressBar.Visible = true;
            statusLabel.Text = "Processing...";
            statusLabel.ForeColor = System.Drawing.Color.Orange;
            downloadButton.Enabled = false;
            
            // Show spinner in preview
            spinnerIndex = 0;
            spinnerPanel.Visible = true;
            spinnerLabel.Visible = true;
            dataGridView.Visible = false;
            spinnerPanel.BringToFront();
            if (!spinnerTimer.Enabled)
                spinnerTimer.Start();

            try
            {
                int month = monthComboBox.SelectedIndex + 1;
                int year = (int)yearInput.Value;
                int holidays = (int)holidayInput.Value;

                var summary = ExcelHelper.ProcessAttendanceFiles(uploadedFiles, year, month, holidays);

                if (summary.Count == 0)
                {
                    // Hide spinner on no data
                    spinnerTimer.Stop();
                    spinnerPanel.Visible = false;
                    spinnerLabel.Visible = false;
                    dataGridView.Visible = true;
                    
                    MessageBox.Show("No attendance data found for the selected month.", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    statusLabel.Text = "No data found";
                    statusLabel.ForeColor = System.Drawing.Color.Red;
                    return;
                }

                // Hide spinner and show preview
                spinnerTimer.Stop();
                spinnerPanel.Visible = false;
                spinnerLabel.Visible = false;
                dataGridView.Visible = true;
                
                // Show preview
                dataGridView.DataSource = summary.Take(200).ToList();

                excelBytes = ExcelHelper.BuildExcelWorkbook(summary, month, year);

                statusLabel.Text = "Analysis complete! Ready to download.";
                statusLabel.ForeColor = System.Drawing.Color.Green;
                downloadButton.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                statusLabel.Text = "Error occurred";
                statusLabel.ForeColor = System.Drawing.Color.Red;
                
                // Hide spinner on error
                spinnerTimer.Stop();
                spinnerPanel.Visible = false;
                spinnerLabel.Visible = false;
                dataGridView.Visible = true;
            }
            finally
            {
                progressBar.Visible = false;
            }
        }

        private void DownloadExcel()
        {
            if (excelBytes == null)
            {
                MessageBox.Show("No data to download. Please run analysis first.", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (var dialog = new SaveFileDialog())
            {
                var monthName = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(monthComboBox.SelectedIndex + 1);
                dialog.FileName = $"{monthName.ToUpper()}_summary.xlsx";
                dialog.Filter = "Excel files|*.xlsx";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    File.WriteAllBytes(dialog.FileName, excelBytes);
                    MessageBox.Show($"File saved to:\n{dialog.FileName}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void ToggleInfoCard()
        {
            try
            {
                infoCardExpanded = !infoCardExpanded;

                if (infoCardExpanded && infoCardPanel.Controls.Count == 0)
                {
                    PopulateInfoCard();
                }

                if (infoCardTimer != null)
                {
                    infoCardTimer.Start();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error toggling info card: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void InfoCardTimer_Tick(object sender, EventArgs e)
        {
            int targetHeight = infoCardExpanded ? INFO_CARD_MAX_HEIGHT : 0;
            int step = 10;

            if (infoCardExpanded)
            {
                if (infoCardPanel.Height < targetHeight)
                {
                    infoCardPanel.Height = Math.Min(infoCardPanel.Height + step, targetHeight);
                }
                else
                {
                    infoCardTimer.Stop();
                }
            }
            else
            {
                if (infoCardPanel.Height > targetHeight)
                {
                    infoCardPanel.Height = Math.Max(infoCardPanel.Height - step, targetHeight);
                }
                else
                {
                    infoCardTimer.Stop();
                }
            }

            infoCardPanel.Invalidate();
        }

        private void PopulateInfoCard()
        {
            infoCardPanel.Controls.Clear();

            int yPos = 10;

            // App icon at top
            if (File.Exists("icon.ico"))
            {
                var iconPictureBox = new PictureBox
                {
                    Image = new System.Drawing.Bitmap("icon.ico"),
                    Size = new System.Drawing.Size(60, 60),
                    Location = new System.Drawing.Point(10, yPos),
                    SizeMode = PictureBoxSizeMode.StretchImage
                };
                infoCardPanel.Controls.Add(iconPictureBox);
                yPos += 70;
            }

            // Developed by label
            var developerLabel = new Label
            {
                Text = "Developed by:",
                Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(59, 130, 246),
                Location = new System.Drawing.Point(10, yPos),
                AutoSize = true
            };
            infoCardPanel.Controls.Add(developerLabel);
            yPos += 22;

            // Developer name
            var developerNameTextBox = new TextBox
            {
                Text = "Chrispen Dery",
                Font = new System.Drawing.Font("Segoe UI", 10),
                Location = new System.Drawing.Point(10, yPos),
                Width = 300,
                Height = 25,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                BackColor = System.Drawing.Color.White,
                ForeColor = System.Drawing.Color.Black,
                Cursor = Cursors.Hand
            };
            infoCardPanel.Controls.Add(developerNameTextBox);
            yPos += 35;

            // Email label
            var emailLabel = new Label
            {
                Text = "Email:",
                Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(59, 130, 246),
                Location = new System.Drawing.Point(10, yPos),
                AutoSize = true
            };
            infoCardPanel.Controls.Add(emailLabel);
            yPos += 22;

            // Email value
            var emailValueTextBox = new TextBox
            {
                Text = "derychrispen72@gmail.com",
                Font = new System.Drawing.Font("Segoe UI", 9),
                Location = new System.Drawing.Point(10, yPos),
                Width = 300,
                Height = 25,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                BackColor = System.Drawing.Color.White,
                ForeColor = System.Drawing.Color.FromArgb(30, 144, 255),
                Cursor = Cursors.Hand
            };
            emailValueTextBox.Click += (s, e) =>
            {
                try
                {
                    var psi = new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = "mailto:derychrispen72@gmail.com",
                        UseShellExecute = true
                    };
                    System.Diagnostics.Process.Start(psi);
                }
                catch
                {
                    MessageBox.Show("Could not open email client. Please check your email settings.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            };
            emailValueTextBox.DoubleClick += (s, e) =>
            {
                try
                {
                    var psi = new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = "mailto:derychrispen72@gmail.com",
                        UseShellExecute = true
                    };
                    System.Diagnostics.Process.Start(psi);
                }
                catch
                {
                    MessageBox.Show("Could not open email client. Please check your email settings.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            };
            infoCardPanel.Controls.Add(emailValueTextBox);
            yPos += 35;

            // Contact label
            var contactLabel = new Label
            {
                Text = "Contact:",
                Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(59, 130, 246),
                Location = new System.Drawing.Point(10, yPos),
                AutoSize = true
            };
            infoCardPanel.Controls.Add(contactLabel);
            yPos += 22;

            // Contact value
            var contactValueTextBox = new TextBox
            {
                Text = "+233 55 0722 898",
                Font = new System.Drawing.Font("Segoe UI", 10),
                Location = new System.Drawing.Point(10, yPos),
                Width = 300,
                Height = 25,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                BackColor = System.Drawing.Color.White,
                ForeColor = System.Drawing.Color.Black,
                Cursor = Cursors.Hand
            };
            contactValueTextBox.Click += (s, e) =>
            {
                try
                {
                    System.Windows.Forms.Clipboard.SetText("+233 55 0722 898");
                    MessageBox.Show("Phone number copied to clipboard!", "Copied", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch { }
            };
            infoCardPanel.Controls.Add(contactValueTextBox);

            // Add mouse down handler to main form to detect clicks
            mainPanel.MouseDown += (s, e) =>
            {
                if (infoCardExpanded)
                {
                    // Check if click is outside the info card and button
                    Point globalLoc = this.PointToClient(System.Windows.Forms.Control.MousePosition);
                    Rectangle infoCardBounds = infoCardPanel.Bounds;
                    Rectangle infoButtonBounds = infoButton.Bounds;
                    
                    if (!infoCardBounds.Contains(globalLoc) && !infoButtonBounds.Contains(globalLoc))
                    {
                        infoCardExpanded = false;
                        if (infoCardTimer != null)
                        {
                            infoCardTimer.Start();
                        }
                    }
                }
            };
        }

        private void CloseInfoCardIfNeeded(MouseEventArgs e)
        {
            if (!infoCardExpanded)
                return;

            // Check if click is on the info button or info card panel
            Rectangle infoButtonBounds = infoButton.Bounds;
            Rectangle infoCardBounds = infoCardPanel.Bounds;

            bool clickedOnButton = infoButtonBounds.Contains(e.Location);
            bool clickedOnCard = infoCardBounds.Contains(e.Location);

            // Close the card if clicked outside of it and the button
            if (!clickedOnButton && !clickedOnCard)
            {
                infoCardExpanded = false;
                if (infoCardTimer != null)
                {
                    infoCardTimer.Start();
                }
            }
        }

        private void SpinnerTimer_Tick(object sender, EventArgs e)
        {
            if (spinnerLabel != null && spinnerLabel.Visible)
            {
                spinnerLabel.Text = spinnerFrames[spinnerIndex];
                spinnerIndex = (spinnerIndex + 1) % spinnerFrames.Length;
            }
        }
    }
}
