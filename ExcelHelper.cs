using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace AttendanceApp
{
    public class SummaryRecord
    {
        public int PersonID { get; set; }
        public string Name { get; set; }
        public string Department { get; set; }
        public int ClockInCount { get; set; }
        public int ClockOutCount { get; set; }
        public int WorkingDays { get; set; }
        public int NotClockedIn { get; set; }
        public double AttendancePercent { get; set; }
    }

    public class ExcelHelper
    {
        public static List<SummaryRecord> ProcessAttendanceFiles(List<string> filePaths, int year, int month, int holidayOverride)
        {
            var allRecords = new List<AttendanceRecord>();
            var errors = new List<string>();

            foreach (var filePath in filePaths)
            {
                try
                {
                    var records = AttendanceParser.ParseFile(filePath, year, month);
                    allRecords.AddRange(records);
                }
                catch (Exception ex)
                {
                    errors.Add($"{Path.GetFileName(filePath)}: {ex.Message}");
                }
            }

            if (errors.Count > 0 && allRecords.Count == 0)
            {
                throw new InvalidOperationException(
                    "Failed to parse all files:\n" + string.Join("\n", errors));
            }

            if (allRecords.Count == 0)
                return new List<SummaryRecord>();

            // Aggregate by Person ID, Name, Department
            var grouped = allRecords
                .GroupBy(r => new { r.PersonID, r.Name, r.Department })
                .Select(g => new
                {
                    g.Key.PersonID,
                    g.Key.Name,
                    g.Key.Department,
                    ClockInCount = g.Sum(r => r.ClockInCount),
                    ClockOutCount = g.Sum(r => r.ClockOutCount)
                })
                .ToList();

            // Calculate working days
            var businessDays = AttendanceParser.GetBusinessDaysInMonth(year, month);
            var workingDays = Math.Max(0, businessDays - holidayOverride);

            // Create summary with attendance percentage
            var summary = grouped.Select(g =>
            {
                var notClockedIn = Math.Max(0, workingDays - g.ClockInCount);
                var attendPct = workingDays > 0
                    ? Math.Round((double)g.ClockInCount / workingDays * 100, 1)
                    : 0;

                return new SummaryRecord
                {
                    PersonID = g.PersonID,
                    Name = g.Name,
                    Department = g.Department,
                    ClockInCount = g.ClockInCount,
                    ClockOutCount = g.ClockOutCount,
                    WorkingDays = workingDays,
                    NotClockedIn = notClockedIn,
                    AttendancePercent = attendPct
                };
            }).ToList();

            return summary;
        }

        // ── Theme colors (Greens) ──
        private static readonly XLColor HeaderBg     = XLColor.FromArgb(20, 70, 45);
        private static readonly XLColor HeaderFg     = XLColor.White;
        private static readonly XLColor DeptBg       = XLColor.FromArgb(40, 140, 90);
        private static readonly XLColor DeptFg       = XLColor.White;
        private static readonly XLColor ZebraLight   = XLColor.FromArgb(245, 252, 248);
        private static readonly XLColor ZebraDark    = XLColor.White;
        private static readonly XLColor SummaryBg    = XLColor.FromArgb(225, 240, 230);
        private static readonly XLColor BorderClr    = XLColor.FromArgb(200, 215, 210);
        private static readonly XLColor GoodGreen    = XLColor.FromArgb(34, 139, 34);
        private static readonly XLColor WarnOrange   = XLColor.FromArgb(218, 165, 32);
        private static readonly XLColor BadRed       = XLColor.FromArgb(178, 34, 34);
        private static readonly XLColor TitleBg      = XLColor.FromArgb(25, 95, 60);
        private static readonly XLColor TitleFg      = XLColor.White;
        private static readonly XLColor SubtitleFg   = XLColor.FromArgb(200, 230, 215);

        public static byte[] BuildExcelWorkbook(List<SummaryRecord> summary, int month, int year)
        {
            if (summary == null || summary.Count == 0)
                return null;

            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Attendance Summary");
                ws.Style.Font.FontName = "Calibri";

                var monthName = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(month);

                // ────────────────────────────
                //  TITLE BANNER (rows 1-3)
                // ────────────────────────────
                var titleRange = ws.Range("A1:H3");
                titleRange.Merge();
                titleRange.Style.Fill.BackgroundColor = TitleBg;

                var titleCell = ws.Cell("A1");
                titleCell.Value = $"  Attendance Summary Report";
                titleCell.Style.Font.FontSize = 18;
                titleCell.Style.Font.Bold = true;
                titleCell.Style.Font.FontColor = TitleFg;
                titleCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                titleCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                ws.Row(1).Height = 20;
                ws.Row(2).Height = 20;
                ws.Row(3).Height = 20;

                // Subtitle row
                var subtitleRange = ws.Range("A4:H4");
                subtitleRange.Merge();
                subtitleRange.Style.Fill.BackgroundColor = TitleBg;
                var subtitleCell = ws.Cell("A4");
                subtitleCell.Value = $"  {monthName} {year}  •  Generated {DateTime.Now:dd MMM yyyy, HH:mm}";
                subtitleCell.Style.Font.FontSize = 10;
                subtitleCell.Style.Font.FontColor = SubtitleFg;
                subtitleCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                ws.Row(4).Height = 22;

                // Spacer
                ws.Row(5).Height = 12;

                var departments = summary.Select(r => r.Department).Distinct().OrderBy(d => d).ToList();
                var currentRow = 6;
                int deptIndex = 0;

                foreach (var dept in departments)
                {
                    deptIndex++;

                    // ── Department header bar ──
                    var deptRange = ws.Range(currentRow, 1, currentRow, 8);
                    deptRange.Merge();
                    deptRange.Style.Fill.BackgroundColor = DeptBg;
                    deptRange.Style.Font.Bold = true;
                    deptRange.Style.Font.FontSize = 11;
                    deptRange.Style.Font.FontColor = DeptFg;

                    var deptData = summary.Where(r => r.Department == dept).OrderBy(r => r.PersonID).ToList();
                    ws.Cell(currentRow, 1).Value = $"  {dept}  ({deptData.Count} employees)";
                    ws.Cell(currentRow, 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    ws.Row(currentRow).Height = 28;
                    currentRow++;

                    // ── Column headers ──
                    var headers = new[] { "#", "ID", "Name", "Working Days", "Clock In", "Clock Out", "Absent", "Attendance %" };
                    for (int i = 0; i < headers.Length; i++)
                    {
                        var cell = ws.Cell(currentRow, i + 1);
                        cell.Value = headers[i];
                        cell.Style.Font.Bold = true;
                        cell.Style.Font.FontSize = 10;
                        cell.Style.Font.FontColor = HeaderFg;
                        cell.Style.Fill.BackgroundColor = HeaderBg;
                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        cell.Style.Border.BottomBorderColor = BorderClr;
                    }
                    ws.Row(currentRow).Height = 24;
                    currentRow++;

                    // ── Data rows ──
                    int seq = 0;
                    foreach (var record in deptData)
                    {
                        seq++;
                        bool isEven = seq % 2 == 0;
                        var rowBg = isEven ? ZebraLight : ZebraDark;

                        ws.Cell(currentRow, 1).Value = seq;
                        ws.Cell(currentRow, 2).Value = record.PersonID;
                        ws.Cell(currentRow, 3).Value = record.Name;
                        ws.Cell(currentRow, 4).Value = record.WorkingDays;
                        ws.Cell(currentRow, 5).Value = record.ClockInCount;
                        ws.Cell(currentRow, 6).Value = record.ClockOutCount;
                        ws.Cell(currentRow, 7).FormulaA1 = $"D{currentRow}-E{currentRow}";
                        ws.Cell(currentRow, 8).Value = record.AttendancePercent / 100.0;
                        ws.Cell(currentRow, 8).Style.NumberFormat.Format = "0.0%";

                        // Conditional color on attendance %
                        var pctColor = record.AttendancePercent >= 80 ? GoodGreen
                                     : record.AttendancePercent >= 50 ? WarnOrange
                                     : BadRed;
                        ws.Cell(currentRow, 8).Style.Font.FontColor = pctColor;
                        ws.Cell(currentRow, 8).Style.Font.Bold = true;

                        // Style the row
                        for (int col = 1; col <= 8; col++)
                        {
                            var cell = ws.Cell(currentRow, col);
                            cell.Style.Fill.BackgroundColor = rowBg;
                            cell.Style.Border.BottomBorder = XLBorderStyleValues.Hair;
                            cell.Style.Border.BottomBorderColor = BorderClr;
                            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                            // Center numeric columns
                            if (col != 3)
                                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        }

                        ws.Row(currentRow).Height = 20;
                        currentRow++;
                    }

                    // ── Department summary row ──
                    ws.Cell(currentRow, 1).Value = "";
                    ws.Cell(currentRow, 2).Value = "";
                    ws.Cell(currentRow, 3).Value = "TOTAL";
                    ws.Cell(currentRow, 3).Style.Font.Bold = true;
                    ws.Cell(currentRow, 4).Value = deptData.FirstOrDefault()?.WorkingDays ?? 0;
                    ws.Cell(currentRow, 5).Value = deptData.Sum(r => r.ClockInCount);
                    ws.Cell(currentRow, 6).Value = deptData.Sum(r => r.ClockOutCount);
                    ws.Cell(currentRow, 7).Value = deptData.Sum(r => r.NotClockedIn);

                    if (deptData.Count > 0)
                    {
                        var avgPct = Math.Round(deptData.Average(r => r.AttendancePercent), 1);
                        ws.Cell(currentRow, 8).Value = avgPct / 100.0;
                        ws.Cell(currentRow, 8).Style.NumberFormat.Format = "0.0%";
                        var avgColor = avgPct >= 80 ? GoodGreen : avgPct >= 50 ? WarnOrange : BadRed;
                        ws.Cell(currentRow, 8).Style.Font.FontColor = avgColor;
                    }

                    for (int col = 1; col <= 8; col++)
                    {
                        var cell = ws.Cell(currentRow, col);
                        cell.Style.Fill.BackgroundColor = SummaryBg;
                        cell.Style.Font.Bold = true;
                        cell.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        cell.Style.Border.TopBorderColor = BorderClr;
                        cell.Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                        cell.Style.Border.BottomBorderColor = XLColor.FromArgb(100, 110, 140);
                        cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        if (col != 3)
                            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    }
                    ws.Row(currentRow).Height = 22;
                    currentRow++;

                    // Gap between departments
                    ws.Row(currentRow).Height = 16;
                    currentRow++;
                }

                // ────────────────────────────
                //  COLUMN WIDTHS
                // ────────────────────────────
                ws.Column(1).Width = 5;     // #
                ws.Column(2).Width = 10;    // ID
                ws.Column(3).Width = 28;    // Name
                ws.Column(4).Width = 14;    // Working Days
                ws.Column(5).Width = 12;    // Clock In
                ws.Column(6).Width = 12;    // Clock Out
                ws.Column(7).Width = 12;    // Absent
                ws.Column(8).Width = 14;    // Attendance %

                // ────────────────────────────
                //  PRINT SETTINGS
                // ────────────────────────────
                ws.PageSetup.PageOrientation = XLPageOrientation.Landscape;
                ws.PageSetup.PaperSize = XLPaperSize.A4Paper;
                ws.PageSetup.FitToPages(1, 0);
                ws.SheetView.FreezeRows(6); // Freeze below first dept header

                using (var ms = new MemoryStream())
                {
                    wb.SaveAs(ms);
                    return ms.ToArray();
                }
            }
        }
    }
}
