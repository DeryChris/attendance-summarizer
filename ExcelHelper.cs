using System;
using System.Collections.Generic;
using System.Data;
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
    }

    public class ExcelHelper
    {
        public static List<SummaryRecord> ProcessAttendanceFiles(List<string> filePaths, int year, int month, int holidayOverride)
        {
            var allRecords = new List<AttendanceRecord>();

            foreach (var filePath in filePaths)
            {
                var records = AttendanceParser.ParseCSV(filePath, year, month);
                allRecords.AddRange(records);
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
            var workingDays = businessDays - holidayOverride;
            if (workingDays < 0) workingDays = 0;

            // Create summary
            var summary = grouped.Select(g => new SummaryRecord
            {
                PersonID = g.PersonID,
                Name = g.Name,
                Department = g.Department,
                ClockInCount = g.ClockInCount,
                ClockOutCount = g.ClockOutCount,
                WorkingDays = workingDays,
                NotClockedIn = Math.Max(0, workingDays - g.ClockInCount)
            }).ToList();

            return summary;
        }

        public static byte[] BuildExcelWorkbook(List<SummaryRecord> summary, int month, int year)
        {
            if (summary == null || summary.Count == 0)
                return null;

            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Summary");

                var monthName = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(month);
                var title = $"Upper West Regional Health Directorate Attendance Register For The Month of {monthName} {year}";

                // Row 1: Empty
                ws.Row(1).Height = 20;

                // Row 2: Title
                ws.Cell("B2").Value = title;
                ws.Range("B2:I2").Merge();
                ws.Cell("B2").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                ws.Cell("B2").Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                ws.Row(2).Height = 30;

                // Row 3-4: Empty
                ws.Row(3).Height = 10;
                ws.Row(4).Height = 10;

                var departments = summary.Select(r => r.Department).Distinct().OrderBy(d => d).ToList();
                var currentRow = 5;

                foreach (var dept in departments)
                {
                    // Department header (purple, bold, no borders)
                    ws.Cell(currentRow, 2).Value = $"DEP: {dept}";
                    ws.Cell(currentRow, 2).Style.Font.Bold = true;
                    ws.Cell(currentRow, 2).Style.Font.FontColor = XLColor.FromArgb(112, 48, 160); // Purple
                    ws.Row(currentRow).Height = 20;
                    currentRow++;

                    // Blank row
                    ws.Row(currentRow).Height = 10;
                    currentRow++;

                    // Column headers (bold with black borders)
                    var headers = new[] { null, "ID", "Names", "No of Days", null, "Clock in", "Clock out", null, "Not Clock In" };
                    for (int i = 0; i < headers.Length; i++)
                    {
                        var cell = ws.Cell(currentRow, i + 1);
                        if (headers[i] != null)
                        {
                            cell.Value = headers[i];
                            cell.Style.Font.Bold = true;
                            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
                            cell.Style.Border.OutsideBorderColor = XLColor.Black;
                        }
                    }
                    ws.Row(currentRow).Height = 20;
                    currentRow++;

                    // Data rows
                    var deptData = summary.Where(r => r.Department == dept).OrderBy(r => r.PersonID).ToList();
                    foreach (var record in deptData)
                    {
                        ws.Cell(currentRow, 1).Value = "";
                        ws.Cell(currentRow, 2).Value = record.PersonID;
                        ws.Cell(currentRow, 3).Value = record.Name;
                        ws.Cell(currentRow, 4).Value = record.WorkingDays;
                        ws.Cell(currentRow, 5).Value = "";
                        ws.Cell(currentRow, 6).Value = record.ClockInCount;
                        ws.Cell(currentRow, 7).Value = record.ClockOutCount;
                        ws.Cell(currentRow, 8).Value = "";
                        ws.Cell(currentRow, 9).FormulaA1 = $"D{currentRow}-F{currentRow}";

                        // Apply borders to data rows
                        for (int col = 2; col <= 9; col++)
                        {
                            var cell = ws.Cell(currentRow, col);
                            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
                            cell.Style.Border.OutsideBorderColor = XLColor.Black;
                        }

                        ws.Row(currentRow).Height = 18;
                        currentRow++;
                    }

                    // Blank rows after department
                    for (int i = 0; i < 3; i++)
                    {
                        ws.Row(currentRow).Height = 10;
                        currentRow++;
                    }
                }

                // Set column widths
                ws.Column(1).Width = 3;
                ws.Column(2).Width = 10;
                ws.Column(3).Width = 25;
                ws.Column(4).Width = 12;
                ws.Column(5).Width = 3;
                ws.Column(6).Width = 12;
                ws.Column(7).Width = 12;
                ws.Column(8).Width = 3;
                ws.Column(9).Width = 15;

                using (var ms = new MemoryStream())
                {
                    wb.SaveAs(ms);
                    return ms.ToArray();
                }
            }
        }
    }
}
