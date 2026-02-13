using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace AttendanceApp
{
    public class AttendanceRecord
    {
        public int PersonID { get; set; }
        public string Name { get; set; }
        public string Department { get; set; }
        public int ClockInCount { get; set; }
        public int ClockOutCount { get; set; }
    }

    public class AttendanceParser
    {
        /// <summary>
        /// Parses a CSV or XLSX attendance file and returns attendance records
        /// for the specified year and month.
        /// </summary>
        public static List<AttendanceRecord> ParseFile(string filePath, int selectedYear, int selectedMonth)
        {
            var ext = Path.GetExtension(filePath).ToLowerInvariant();

            if (ext == ".xlsx")
                return ParseXlsx(filePath, selectedYear, selectedMonth);
            else if (ext == ".csv")
                return ParseCsv(filePath, selectedYear, selectedMonth);
            else
                throw new NotSupportedException($"Unsupported file format: {ext}. Only .csv and .xlsx are supported.");
        }

        private static List<AttendanceRecord> ParseCsv(string filePath, int selectedYear, int selectedMonth)
        {
            var lines = new List<string[]>();

            using (var reader = new StreamReader(filePath))
            {
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    if (!string.IsNullOrWhiteSpace(line))
                    {
                        lines.Add(line.Split(','));
                    }
                }
            }

            return ParseRawRows(lines, filePath, selectedYear, selectedMonth);
        }

        private static List<AttendanceRecord> ParseXlsx(string filePath, int selectedYear, int selectedMonth)
        {
            var lines = new List<string[]>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var sheet = workbook.Worksheets.First();
                var range = sheet.RangeUsed();

                if (range == null)
                    return new List<AttendanceRecord>();

                int rowCount = range.RowCount();
                int colCount = range.ColumnCount();

                for (int r = 1; r <= rowCount; r++)
                {
                    var row = new string[colCount];
                    for (int c = 1; c <= colCount; c++)
                    {
                        var cell = sheet.Cell(r, c);
                        row[c - 1] = cell.GetFormattedString() ?? "";
                    }
                    lines.Add(row);
                }
            }

            return ParseRawRows(lines, filePath, selectedYear, selectedMonth);
        }

        /// <summary>
        /// Core parsing logic shared between CSV and XLSX formats.
        /// Detects header columns, filters by date, and builds attendance records.
        /// </summary>
        private static List<AttendanceRecord> ParseRawRows(List<string[]> lines, string filePath, int selectedYear, int selectedMonth)
        {
            var records = new List<AttendanceRecord>();
            var fallbackDepartment = Path.GetFileNameWithoutExtension(filePath);

            if (lines.Count == 0)
                return records;

            // Find header row by scanning for "person id" column
            int headerIdx = -1;
            var headerMap = new Dictionary<string, int>();

            for (int i = 0; i < lines.Count; i++)
            {
                for (int j = 0; j < lines[i].Length; j++)
                {
                    var cell = lines[i][j].ToLower().Trim();
                    if (cell.Contains("person id") || cell.Contains("person_id") || cell.Contains("personid"))
                    {
                        headerIdx = i;
                        for (int k = 0; k < lines[i].Length; k++)
                        {
                            var col = lines[i][k].ToLower().Trim();
                            if (col.Contains("person id") || col.Contains("person_id") || col.Contains("personid"))
                                headerMap["id"] = k;
                            else if (col.Contains("name"))
                                headerMap["name"] = k;
                            else if (col.Contains("date"))
                                headerMap["date"] = k;
                            else if (col.Contains("check-in") || col.Contains("checkin") || col.Contains("clock in") || col.Contains("clockin"))
                                headerMap["checkin"] = k;
                            else if (col.Contains("check-out") || col.Contains("checkout") || col.Contains("clock out") || col.Contains("clockout"))
                                headerMap["checkout"] = k;
                            else if (col.Contains("department") || col.Contains("dept"))
                                headerMap["department"] = k;
                        }
                        break;
                    }
                }
                if (headerIdx != -1) break;
            }

            // Validate required columns are found
            var requiredColumns = new[] { "id", "name", "date", "checkin", "checkout" };
            var missingColumns = requiredColumns.Where(k => !headerMap.ContainsKey(k)).ToList();

            if (headerIdx == -1 || missingColumns.Count > 0)
            {
                var missing = missingColumns.Count > 0 ? string.Join(", ", missingColumns) : "header row";
                throw new InvalidDataException(
                    $"Could not detect required columns in '{Path.GetFileName(filePath)}'. " +
                    $"Missing: {missing}. " +
                    $"Expected columns: Person ID, Name, Date, Check-in, Check-out.");
            }

            // Parse data rows
            int skippedRows = 0;
            for (int i = headerIdx + 1; i < lines.Count; i++)
            {
                var row = lines[i];
                if (row.Length == 0 || row.All(c => string.IsNullOrWhiteSpace(c)))
                    continue;

                // Skip known non-data rows (summary rows in some exports)
                var firstCol = row[0].ToLower().Trim();
                if (firstCol.StartsWith("check-in time") || firstCol.StartsWith("attended:") ||
                    (firstCol.Contains(":") && firstCol.Contains("check")))
                    continue;

                try
                {
                    // Parse Person ID
                    var personIdStr = headerMap["id"] < row.Length ? row[headerMap["id"]].Trim() : "";
                    if (string.IsNullOrEmpty(personIdStr) || !int.TryParse(personIdStr, out var personId))
                    {
                        skippedRows++;
                        continue;
                    }

                    // Parse and filter by date
                    var dateStr = headerMap["date"] < row.Length ? row[headerMap["date"]].Trim() : "";
                    if (!DateTime.TryParse(dateStr, CultureInfo.InvariantCulture, DateTimeStyles.None, out var punchDate) &&
                        !DateTime.TryParse(dateStr, out punchDate))
                    {
                        skippedRows++;
                        continue;
                    }

                    if (punchDate.Year != selectedYear || punchDate.Month != selectedMonth)
                        continue;

                    // Extract fields
                    var name = headerMap["name"] < row.Length ? row[headerMap["name"]].Trim() : "";
                    var checkinVal = headerMap["checkin"] < row.Length ? row[headerMap["checkin"]].Trim() : "";
                    var checkoutVal = headerMap["checkout"] < row.Length ? row[headerMap["checkout"]].Trim() : "";

                    var dept = (headerMap.ContainsKey("department") && headerMap["department"] < row.Length)
                        ? row[headerMap["department"]].Trim()
                        : fallbackDepartment;

                    if (string.IsNullOrEmpty(dept) || dept.Contains(":"))
                        dept = fallbackDepartment;

                    // Determine clock status — a value starting with a digit indicates a valid time
                    var clockInCount = !string.IsNullOrEmpty(checkinVal) && char.IsDigit(checkinVal[0]) ? 1 : 0;
                    var clockOutCount = !string.IsNullOrEmpty(checkoutVal) && char.IsDigit(checkoutVal[0]) ? 1 : 0;

                    records.Add(new AttendanceRecord
                    {
                        PersonID = personId,
                        Name = name,
                        Department = dept,
                        ClockInCount = clockInCount,
                        ClockOutCount = clockOutCount
                    });
                }
                catch (IndexOutOfRangeException)
                {
                    skippedRows++;
                    continue;
                }
            }

            return records;
        }

        /// <summary>
        /// Calculates the number of business days (Mon-Fri) in a given month,
        /// optionally excluding specific holiday dates.
        /// </summary>
        public static int GetBusinessDaysInMonth(int year, int month, List<DateTime> holidays = null)
        {
            var daysInMonth = DateTime.DaysInMonth(year, month);
            int count = 0;

            for (int day = 1; day <= daysInMonth; day++)
            {
                var date = new DateTime(year, month, day);
                if (date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday)
                {
                    if (holidays == null || !holidays.Any(h => h.Date == date.Date))
                        count++;
                }
            }

            return count;
        }
    }
}
