using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using CsvHelper;

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
        public static List<AttendanceRecord> ParseCSV(string filePath, int selectedYear, int selectedMonth)
        {
            var records = new List<AttendanceRecord>();
            var department = Path.GetFileNameWithoutExtension(filePath);

            try
            {
                using (var reader = new StreamReader(filePath))
                using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    var lines = new List<string[]>();
                    while (reader.Peek() != -1)
                    {
                        var line = reader.ReadLine();
                        if (!string.IsNullOrWhiteSpace(line))
                        {
                            lines.Add(line.Split(','));
                        }
                    }

                    if (lines.Count == 0) return records;

                    // Find header row
                    int headerIdx = -1;
                    var headerMap = new Dictionary<string, int>();

                    for (int i = 0; i < lines.Count; i++)
                    {
                        for (int j = 0; j < lines[i].Length; j++)
                        {
                            var cell = lines[i][j].ToLower().Trim();
                            if (cell.Contains("person id") || cell.Contains("person_id"))
                            {
                                headerIdx = i;
                                for (int k = 0; k < lines[i].Length; k++)
                                {
                                    var col = lines[i][k].ToLower().Trim();
                                    if (col.Contains("person id") || col.Contains("person_id"))
                                        headerMap["id"] = k;
                                    else if (col.Contains("name"))
                                        headerMap["name"] = k;
                                    else if (col.Contains("date"))
                                        headerMap["date"] = k;
                                    else if (col.Contains("check-in") || col.Contains("checkin"))
                                        headerMap["checkin"] = k;
                                    else if (col.Contains("check-out") || col.Contains("checkout"))
                                        headerMap["checkout"] = k;
                                    else if (col.Contains("department"))
                                        headerMap["department"] = k;
                                }
                                break;
                            }
                        }
                        if (headerIdx != -1) break;
                    }

                    if (headerIdx == -1 || !new[] { "id", "name", "date", "checkin", "checkout" }
                        .All(k => headerMap.ContainsKey(k)))
                        return records;

                    // Parse data rows
                    for (int i = headerIdx + 1; i < lines.Count; i++)
                    {
                        var row = lines[i];
                        if (row.Length == 0) continue;

                        var firstCol = row[0].ToLower().Trim();
                        if (firstCol.StartsWith("check-in time") || firstCol.StartsWith("attended:") ||
                            (firstCol.Contains(":") && firstCol.Contains("check")))
                            continue;

                        try
                        {
                            var personIdStr = row[headerMap["id"]].Trim();
                            if (string.IsNullOrEmpty(personIdStr) || !int.TryParse(personIdStr, out var personId))
                                continue;

                            var dateStr = row[headerMap["date"]].Trim();
                            if (!DateTime.TryParse(dateStr, out var punchDate) ||
                                punchDate.Year != selectedYear || punchDate.Month != selectedMonth)
                                continue;

                            var name = row[headerMap["name"]].Trim();
                            var checkinVal = headerMap["checkin"] < row.Length ? row[headerMap["checkin"]].Trim() : "";
                            var checkoutVal = headerMap["checkout"] < row.Length ? row[headerMap["checkout"]].Trim() : "";

                            var dept = (headerMap.ContainsKey("department") && headerMap["department"] < row.Length)
                                ? row[headerMap["department"]].Trim()
                                : department;

                            if (string.IsNullOrEmpty(dept) || dept.Contains(":"))
                                dept = department;

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
                        catch { }
                    }
                }
            }
            catch { }

            return records;
        }

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
