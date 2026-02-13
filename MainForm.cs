using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AttendanceApp
{
    public partial class MainForm : Form
    {
        // ── Data ──
        private List<string> uploadedFiles = new List<string>();
        private byte[] excelBytes;

        // ── Controls ──
        private Panel mainPanel;
        private Panel dragDropPanel;
        private FlowLayoutPanel chipsPanel;
        private ComboBox monthComboBox;
        private NumericUpDown yearInput;
        private NumericUpDown holidayInput;
        private Label statusLabel;
        private Panel progressBar;
        private DataGridView dataGridView;
        private Button downloadButton;
        private Button analyzeButton;
        private Panel spinnerPanel;
        private Label spinnerLabel;
        private Timer spinnerTimer;
        private int spinnerIndex = 0;
        private readonly string[] spinnerFrames = { "⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏" };

        // ── Theme ──
        private static readonly Color BgDark       = Color.FromArgb(18, 18, 30);
        private static readonly Color BgCard       = Color.FromArgb(26, 28, 44);
        private static readonly Color BgInput      = Color.FromArgb(36, 38, 58);
        private static readonly Color AccentBlue   = Color.FromArgb(70, 130, 255);
        private static readonly Color AccentGreen  = Color.FromArgb(40, 200, 110);
        private static readonly Color AccentRed    = Color.FromArgb(230, 70, 70);
        private static readonly Color AccentOrange = Color.FromArgb(255, 165, 50);
        private static readonly Color TextPrimary  = Color.FromArgb(225, 228, 238);
        private static readonly Color TextMuted    = Color.FromArgb(120, 125, 150);
        private static readonly Color Border       = Color.FromArgb(45, 48, 70);
        private static readonly Color ChipBg       = Color.FromArgb(50, 95, 190);
        private static readonly Color DropZoneBg   = Color.FromArgb(22, 24, 40);

        public MainForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            this.Text = "Attendance Summarizer";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = new Font("Segoe UI", 9.5f);
            this.BackColor = BgDark;
            this.ForeColor = TextPrimary;
            this.DoubleBuffered = true;
            this.Padding = new Padding(0);
            this.AutoScaleMode = AutoScaleMode.Dpi;

            // ── Set a comfortable default size, allow resize ──
            this.ClientSize = new Size(940, 700);
            this.MinimumSize = new Size(780, 560);

            BuildLayout();

            this.ResumeLayout(false);
        }

        private void BuildLayout()
        {
            // ══════════════════════════════════════════
            //  MAIN CONTENT AREA (added first — Fill)
            // ══════════════════════════════════════════
            mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(24, 16, 24, 16),
                BackColor = BgDark
            };
            this.Controls.Add(mainPanel);

            // ── Progress accent line (added after Fill) ──
            progressBar = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 2,
                BackColor = AccentBlue,
                Visible = false
            };
            this.Controls.Add(progressBar);

            // ══════════════════════════════════════════
            //  STATUS BAR (added last — docks first)
            // ══════════════════════════════════════════
            var statusBar = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 30,
                BackColor = Color.FromArgb(14, 14, 24),
                Padding = new Padding(18, 0, 18, 0)
            };
            statusLabel = new Label
            {
                Text = "● Ready",
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 8.5f),
                ForeColor = AccentGreen,
                TextAlign = ContentAlignment.MiddleLeft
            };
            statusBar.Controls.Add(statusLabel);
            this.Controls.Add(statusBar);

            // ────────────────────────────
            //  We use a TableLayoutPanel for the main body so that
            //  the top sections have fixed heights and the preview
            //  section stretches to fill remaining space.
            // ────────────────────────────
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                BackColor = Color.Transparent,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            // Row 0: Title        – fixed 50px
            // Row 1: Upload       – fixed 210px
            // Row 2: Settings     – fixed 70px
            // Row 3: Preview      – fill remaining
            // Row 4: Actions      – fixed 52px
            layout.RowCount = 5;
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 210));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 70));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 58));

            // ── Row 0: Title ──
            var titlePanel = BuildTitlePanel();
            layout.Controls.Add(titlePanel, 0, 0);

            // ── Row 1: Upload ──
            var uploadCard = BuildUploadSection();
            layout.Controls.Add(uploadCard, 0, 1);

            // ── Row 2: Settings ──
            var settingsCard = BuildSettingsSection();
            layout.Controls.Add(settingsCard, 0, 2);

            // ── Row 3: Preview (fills remaining) ──
            var previewCard = BuildPreviewSection();
            layout.Controls.Add(previewCard, 0, 3);

            // ── Row 4: Action buttons ──
            var actionPanel = BuildActionPanel();
            layout.Controls.Add(actionPanel, 0, 4);

            mainPanel.Controls.Add(layout);

            // ── Spinner timer ──
            spinnerTimer = new Timer { Interval = 80 };
            spinnerTimer.Tick += (s, e) =>
            {
                if (spinnerLabel != null && spinnerLabel.Visible)
                {
                    spinnerLabel.Text = spinnerFrames[spinnerIndex];
                    spinnerIndex = (spinnerIndex + 1) % spinnerFrames.Length;
                }
            };
        }

        // ══════════════════════════════════════════════════════════════
        //  SECTION BUILDERS
        // ══════════════════════════════════════════════════════════════

        private Panel BuildTitlePanel()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.Transparent,
                Margin = new Padding(0, 0, 0, 4)
            };

            var title = new Label
            {
                Text = "Attendance Summarizer",
                Font = new Font("Segoe UI", 18, FontStyle.Bold),
                ForeColor = TextPrimary,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(2, 0, 0, 0)
            };
            panel.Controls.Add(title);

            var ver = new Label
            {
                Text = "v1.1",
                Font = new Font("Segoe UI", 8.5f),
                ForeColor = TextMuted,
                Dock = DockStyle.Right,
                Width = 40,
                TextAlign = ContentAlignment.MiddleCenter
            };
            panel.Controls.Add(ver);

            return panel;
        }

        private Panel BuildUploadSection()
        {
            var card = MakeCard();
            card.Margin = new Padding(0, 0, 0, 8);

            // Section label drawn by card painter
            var inner = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.Transparent,
                Padding = new Padding(10, 24, 10, 6)
            };

            dragDropPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = DropZoneBg,
                AllowDrop = true
            };
            dragDropPanel.Paint += DragDropPanel_Paint;
            dragDropPanel.DragEnter += (s, e) =>
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    e.Effect = DragDropEffects.Copy;
                    dragDropPanel.BackColor = Color.FromArgb(30, 34, 54);
                    dragDropPanel.Invalidate();
                }
            };
            dragDropPanel.DragLeave += (s, e) =>
            {
                dragDropPanel.BackColor = DropZoneBg;
                dragDropPanel.Invalidate();
            };
            dragDropPanel.DragDrop += (s, e) =>
            {
                dragDropPanel.BackColor = DropZoneBg;
                dragDropPanel.Invalidate();
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                    AddFiles((string[])e.Data.GetData(DataFormats.FileDrop));
            };

            // Header row
            var header = new Panel
            {
                Dock = DockStyle.Top,
                Height = 36,
                BackColor = Color.Transparent,
                Padding = new Padding(8, 2, 4, 2)
            };

            var dragLbl = new Label
            {
                Text = "Drag & drop CSV / XLSX files here",
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 9),
                ForeColor = TextMuted,
                TextAlign = ContentAlignment.MiddleLeft
            };
            header.Controls.Add(dragLbl);

            var btnFlow = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                Width = 200,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                BackColor = Color.Transparent
            };
            var clearBtn = MakeSmallButton("Clear All", AccentRed);
            clearBtn.Margin = new Padding(6, 1, 0, 1);
            clearBtn.Click += (s, e) => ClearFiles();
            btnFlow.Controls.Add(clearBtn);
            var browseBtn = MakeSmallButton("Browse", AccentBlue);
            browseBtn.Margin = new Padding(4, 1, 0, 1);
            browseBtn.Click += (s, e) => UploadFiles();
            btnFlow.Controls.Add(browseBtn);
            header.Controls.Add(btnFlow);

            dragDropPanel.Controls.Add(header);

            chipsPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.Transparent,
                AutoScroll = true,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                Padding = new Padding(6, 2, 6, 2)
            };
            dragDropPanel.Controls.Add(chipsPanel);

            inner.Controls.Add(dragDropPanel);
            card.Tag = "Upload Files";
            card.Controls.Add(inner);
            return card;
        }

        private Panel BuildSettingsSection()
        {
            var card = MakeCard();
            card.Margin = new Padding(0, 0, 0, 8);
            card.Tag = "Settings";

            var flow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                Padding = new Padding(12, 26, 12, 6),
                BackColor = Color.Transparent
            };

            flow.Controls.Add(MakeLabel("Month"));
            monthComboBox = new ComboBox
            {
                Width = 140,
                DropDownStyle = ComboBoxStyle.DropDownList,
                FlatStyle = FlatStyle.Flat,
                BackColor = BgInput,
                ForeColor = TextPrimary,
                Font = new Font("Segoe UI", 9.5f),
                Margin = new Padding(0, 1, 20, 0)
            };
            for (int i = 1; i <= 12; i++)
                monthComboBox.Items.Add(CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(i));
            monthComboBox.SelectedIndex = DateTime.Now.Month - 1;
            flow.Controls.Add(monthComboBox);

            flow.Controls.Add(MakeLabel("Year"));
            yearInput = new NumericUpDown
            {
                Width = 80,
                Minimum = 2000,
                Maximum = 2100,
                Value = DateTime.Now.Year,
                BackColor = BgInput,
                ForeColor = TextPrimary,
                Font = new Font("Segoe UI", 9.5f),
                BorderStyle = BorderStyle.None,
                Margin = new Padding(0, 1, 20, 0)
            };
            flow.Controls.Add(yearInput);

            flow.Controls.Add(MakeLabel("Holidays"));
            holidayInput = new NumericUpDown
            {
                Width = 60,
                Minimum = 0,
                Maximum = 31,
                Value = 0,
                BackColor = BgInput,
                ForeColor = TextPrimary,
                Font = new Font("Segoe UI", 9.5f),
                BorderStyle = BorderStyle.None,
                Margin = new Padding(0, 1, 0, 0)
            };
            flow.Controls.Add(holidayInput);

            card.Controls.Add(flow);
            return card;
        }

        private Panel BuildPreviewSection()
        {
            var card = MakeCard();
            card.Margin = new Padding(0, 0, 0, 8);
            card.Tag = "Preview";

            var inner = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.Transparent,
                Padding = new Padding(10, 26, 10, 6)
            };

            dataGridView = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = BgCard,
                GridColor = Border,
                BorderStyle = BorderStyle.None,
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(32, 36, 56),
                    ForeColor = AccentBlue,
                    Font = new Font("Segoe UI Semibold", 9f),
                    SelectionBackColor = Color.FromArgb(40, 44, 66),
                    SelectionForeColor = AccentBlue,
                    Padding = new Padding(4, 3, 4, 3)
                },
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = BgCard,
                    ForeColor = TextPrimary,
                    SelectionBackColor = Color.FromArgb(44, 48, 72),
                    SelectionForeColor = TextPrimary,
                    Font = new Font("Segoe UI", 9f)
                },
                AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(22, 24, 40),
                    ForeColor = TextPrimary,
                    SelectionBackColor = Color.FromArgb(44, 48, 72),
                    SelectionForeColor = TextPrimary
                },
                RowHeadersVisible = false,
                EnableHeadersVisualStyles = false,
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single,
                ColumnHeadersHeight = 32,
                RowTemplate = { Height = 28 }
            };
            inner.Controls.Add(dataGridView);

            // Spinner overlay
            spinnerPanel = new Panel { Dock = DockStyle.Fill, Visible = false, BackColor = BgCard };
            spinnerLabel = new Label
            {
                Text = "⠋",
                Font = new Font("Segoe UI", 36, FontStyle.Bold),
                ForeColor = AccentBlue,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill,
                BackColor = Color.Transparent
            };
            spinnerPanel.Controls.Add(spinnerLabel);
            inner.Controls.Add(spinnerPanel);

            card.Controls.Add(inner);
            return card;
        }

        private Panel BuildActionPanel()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.Transparent,
                Padding = new Padding(0, 4, 0, 0)
            };

            analyzeButton = MakeButton("⚡  Analyze", AccentBlue, 150);
            analyzeButton.Dock = DockStyle.Left;
            analyzeButton.Click += async (s, e) => await AnalyzeFilesAsync();

            // Right-side button group
            var rightGroup = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.Transparent,
                Margin = new Padding(0)
            };

            var exitBtn = MakeButton("Exit", Color.FromArgb(60, 62, 82), 90);
            exitBtn.Margin = new Padding(0, 0, 6, 0);
            exitBtn.Click += (s, e) => Close();

            downloadButton = MakeButton("📥  Download Excel", AccentGreen, 175);
            downloadButton.Enabled = false;
            downloadButton.Margin = new Padding(0);
            downloadButton.Click += (s, e) => DownloadExcel();

            rightGroup.Controls.Add(exitBtn);
            rightGroup.Controls.Add(downloadButton);

            panel.Controls.Add(analyzeButton);
            panel.Controls.Add(rightGroup);

            return panel;
        }

        // ══════════════════════════════════════════════════════════════
        //  SMALL UI FACTORIES
        // ══════════════════════════════════════════════════════════════

        private Panel MakeCard()
        {
            var card = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.Transparent
            };
            card.Paint += CardPaint;
            return card;
        }

        private void CardPaint(object sender, PaintEventArgs e)
        {
            var card = (Panel)sender;
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

            var rect = new Rectangle(0, 0, card.Width - 1, card.Height - 1);
            using (var path = RoundedRect(rect, 10))
            using (var fill = new SolidBrush(BgCard))
            using (var pen = new Pen(Border, 1))
            {
                e.Graphics.FillPath(fill, path);
                e.Graphics.DrawPath(pen, path);
            }

            // Draw section title from Tag
            var title = card.Tag as string;
            if (!string.IsNullOrEmpty(title))
            {
                using (var f = new Font("Segoe UI Semibold", 9f))
                using (var b = new SolidBrush(TextMuted))
                    e.Graphics.DrawString(title, f, b, 14, 6);
            }
        }

        private Label MakeLabel(string text)
        {
            return new Label
            {
                Text = text,
                AutoSize = true,
                Font = new Font("Segoe UI", 9f),
                ForeColor = TextMuted,
                Margin = new Padding(0, 6, 5, 0)
            };
        }

        private Button MakeButton(string text, Color bg, int width)
        {
            var btn = new Button
            {
                Text = text,
                Width = width,
                Height = 38,
                FlatStyle = FlatStyle.Flat,
                BackColor = bg,
                ForeColor = Color.White,
                Font = new Font("Segoe UI Semibold", 9.5f),
                Cursor = Cursors.Hand,
                Margin = new Padding(0)
            };
            btn.FlatAppearance.BorderSize = 0;
            btn.FlatAppearance.MouseOverBackColor = ControlPaint.Light(bg, 0.12f);
            btn.FlatAppearance.MouseDownBackColor = ControlPaint.Dark(bg, 0.08f);
            btn.Paint += (s, e) =>
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                using (var path = RoundedRect(new Rectangle(0, 0, btn.Width, btn.Height), 7))
                    btn.Region = new Region(path);
            };
            return btn;
        }

        private Button MakeSmallButton(string text, Color bg)
        {
            var btn = new Button
            {
                Text = text,
                Width = 85,
                Height = 30,
                FlatStyle = FlatStyle.Flat,
                BackColor = bg,
                ForeColor = Color.White,
                Font = new Font("Segoe UI Semibold", 8.5f),
                Cursor = Cursors.Hand,
                Margin = new Padding(4, 1, 0, 1)
            };
            btn.FlatAppearance.BorderSize = 0;
            btn.FlatAppearance.MouseOverBackColor = ControlPaint.Light(bg, 0.12f);
            btn.FlatAppearance.MouseDownBackColor = ControlPaint.Dark(bg, 0.08f);
            btn.Paint += (s, e) =>
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                using (var path = RoundedRect(new Rectangle(0, 0, btn.Width, btn.Height), 6))
                    btn.Region = new Region(path);
            };
            return btn;
        }

        private static GraphicsPath RoundedRect(Rectangle bounds, int radius)
        {
            int d = radius * 2;
            var gp = new GraphicsPath();
            gp.AddArc(bounds.X, bounds.Y, d, d, 180, 90);
            gp.AddArc(bounds.Right - d, bounds.Y, d, d, 270, 90);
            gp.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0, 90);
            gp.AddArc(bounds.X, bounds.Bottom - d, d, d, 90, 90);
            gp.CloseFigure();
            return gp;
        }

        // ══════════════════════════════════════════════════════════════
        //  DRAG-DROP ZONE PAINTING
        // ══════════════════════════════════════════════════════════════

        private void DragDropPanel_Paint(object sender, PaintEventArgs e)
        {
            var p = (Panel)sender;
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using (var pen = new Pen(Color.FromArgb(60, AccentBlue), 1.5f) { DashStyle = DashStyle.Dash })
            {
                var r = new Rectangle(3, 3, p.Width - 7, p.Height - 7);
                using (var path = RoundedRect(r, 6))
                    e.Graphics.DrawPath(pen, path);
            }
        }

        // ══════════════════════════════════════════════════════════════
        //  FILE MANAGEMENT
        // ══════════════════════════════════════════════════════════════

        private void AddFiles(IEnumerable<string> files)
        {
            var valid = files
                .Where(f => f.EndsWith(".csv", StringComparison.OrdinalIgnoreCase)
                         || f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (valid.Count == 0)
            {
                MessageBox.Show("Please select only CSV or XLSX files.",
                    "Invalid Files", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int dupes = 0;
            foreach (var f in valid)
            {
                if (uploadedFiles.Contains(f, StringComparer.OrdinalIgnoreCase))
                { dupes++; continue; }
                uploadedFiles.Add(f);
            }

            if (dupes > 0)
            {
                statusLabel.Text = $"● {dupes} duplicate(s) skipped";
                statusLabel.ForeColor = AccentOrange;
            }
            RefreshChips();
        }

        private void RefreshChips()
        {
            chipsPanel.Controls.Clear();
            foreach (var f in uploadedFiles)
                AddChip(f);
        }

        private void AddChip(string filePath)
        {
            var chip = new Panel
            {
                Size = new Size(200, 28),
                BackColor = ChipBg,
                Margin = new Padding(3, 3, 3, 3),
                Cursor = Cursors.Default
            };
            chip.Paint += (s, e) =>
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                using (var path = RoundedRect(new Rectangle(0, 0, chip.Width, chip.Height), 14))
                {
                    chip.Region = new Region(path);
                    using (var b = new SolidBrush(ChipBg))
                        e.Graphics.FillPath(b, path);
                }
            };

            var ext = Path.GetExtension(filePath).ToLowerInvariant();
            var icon = ext == ".xlsx" ? "📗" : "📄";

            var lbl = new Label
            {
                Text = $"{icon} {Path.GetFileName(filePath)}",
                Location = new Point(8, 4),
                Size = new Size(162, 20),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 8.5f),
                AutoEllipsis = true,
                TextAlign = ContentAlignment.MiddleLeft,
                BackColor = Color.Transparent
            };
            chip.Controls.Add(lbl);

            var x = new Button
            {
                Text = "✕",
                Location = new Point(172, 2),
                Size = new Size(24, 24),
                BackColor = Color.Transparent,
                ForeColor = Color.FromArgb(190, 190, 190),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 8, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            x.FlatAppearance.BorderSize = 0;
            x.FlatAppearance.MouseOverBackColor = AccentRed;
            x.Click += (s, e) => { uploadedFiles.Remove(filePath); RefreshChips(); };
            chip.Controls.Add(x);

            chipsPanel.Controls.Add(chip);
        }

        private void ClearFiles()
        {
            if (uploadedFiles.Count == 0) return;
            if (MessageBox.Show("Clear all files?", "Confirm",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                uploadedFiles.Clear();
                RefreshChips();
                statusLabel.Text = "● Ready";
                statusLabel.ForeColor = AccentGreen;
            }
        }

        private void UploadFiles()
        {
            using (var d = new OpenFileDialog())
            {
                d.Filter = "Attendance Files|*.csv;*.xlsx|CSV|*.csv|Excel|*.xlsx";
                d.Multiselect = true;
                if (d.ShowDialog() == DialogResult.OK)
                    AddFiles(d.FileNames);
            }
        }

        // ══════════════════════════════════════════════════════════════
        //  ANALYSIS (async)
        // ══════════════════════════════════════════════════════════════

        private async Task AnalyzeFilesAsync()
        {
            if (uploadedFiles.Count == 0)
            {
                MessageBox.Show("Upload at least one file.", "No Files",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            analyzeButton.Enabled = false;
            downloadButton.Enabled = false;
            progressBar.Visible = true;
            statusLabel.Text = "● Processing…";
            statusLabel.ForeColor = AccentOrange;

            spinnerIndex = 0;
            spinnerPanel.Visible = true;
            spinnerLabel.Visible = true;
            dataGridView.Visible = false;
            spinnerPanel.BringToFront();
            spinnerTimer.Start();

            try
            {
                int month = monthComboBox.SelectedIndex + 1;
                int year = (int)yearInput.Value;
                int holidays = (int)holidayInput.Value;
                var copy = uploadedFiles.ToList();

                var summary = await Task.Run(() =>
                    ExcelHelper.ProcessAttendanceFiles(copy, year, month, holidays));

                if (summary.Count == 0)
                {
                    HideSpinner();
                    MessageBox.Show("No attendance data found for the selected month.",
                        "No Data", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    statusLabel.Text = "● No data found";
                    statusLabel.ForeColor = AccentRed;
                    return;
                }

                HideSpinner();
                dataGridView.DataSource = summary.Take(200).ToList();

                excelBytes = await Task.Run(() =>
                    ExcelHelper.BuildExcelWorkbook(summary, month, year));

                statusLabel.Text = $"● Done — {summary.Count} records ready";
                statusLabel.ForeColor = AccentGreen;
                downloadButton.Enabled = true;
            }
            catch (Exception ex)
            {
                HideSpinner();
                MessageBox.Show($"Error: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                statusLabel.Text = "● Error";
                statusLabel.ForeColor = AccentRed;
            }
            finally
            {
                progressBar.Visible = false;
                analyzeButton.Enabled = true;
            }
        }

        private void HideSpinner()
        {
            spinnerTimer.Stop();
            spinnerPanel.Visible = false;
            spinnerLabel.Visible = false;
            dataGridView.Visible = true;
        }

        // ══════════════════════════════════════════════════════════════
        //  DOWNLOAD
        // ══════════════════════════════════════════════════════════════

        private void DownloadExcel()
        {
            if (excelBytes == null)
            {
                MessageBox.Show("Run analysis first.", "No Data",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            using (var d = new SaveFileDialog())
            {
                var m = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(monthComboBox.SelectedIndex + 1);
                d.FileName = $"{m.ToUpper()}_summary.xlsx";
                d.Filter = "Excel|*.xlsx";
                if (d.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        File.WriteAllBytes(d.FileName, excelBytes);
                        MessageBox.Show($"Saved to:\n{d.FileName}", "Success",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Save failed: {ex.Message}", "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
    }
}
