# Attendance Summarizer

A powerful Windows desktop application designed to streamline attendance data analysis and reporting for organizations of all sizes. Transform raw attendance data into actionable insights with minimal effort.

---

## 📋 Problem Statement

Organizations across industries struggle with the manual processing of attendance data. When attendance records are extracted from various sources (biometric systems, time tracking software, Excel spreadsheets), analyzing individual attendance patterns across multiple departments becomes a **time-consuming, error-prone, and tedious process**.

**Current Challenges:**
- **Manual Data Entry**: HR teams spend hours copying and pasting attendance records
- **Inconsistent Formatting**: Data from different sources comes in various formats, requiring manual normalization
- **Calculation Errors**: Manual calculations are prone to mistakes, leading to inaccurate attendance summaries
- **Departmental Silos**: Cross-departmental analysis requires consolidating data from multiple files
- **Limited Insights**: Without proper analysis tools, companies miss attendance trends and patterns
- **Compliance Issues**: Poor documentation can lead to payroll disputes and compliance violations

---

## ✨ Solution

**Attendance Summarizer** automates the entire attendance analysis workflow. Simply upload your CSV or Excel files, configure a few parameters, and get comprehensive attendance summaries instantly. The app intelligently detects data columns, calculates attendance metrics, and generates professional Excel reports—all without manual intervention.

**Key Benefits:**
- ⚡ **Instant Processing**: Analyze months of attendance data in seconds
- 🎯 **Automatic Detection**: Intelligently identifies Person ID, Name, Date, Check-in, Check-out, and Department columns
- 📊 **Comprehensive Reports**: Generates detailed Excel workbooks with attendance summaries
- 🏢 **Multi-Department Support**: Analyze attendance across all departments simultaneously
- 📈 **Professional Output**: Ready-to-use Excel reports for HR, management, and payroll teams
- 🔧 **Easy to Use**: No technical expertise required—intuitive drag-and-drop interface

---

## 🚀 Features

### Core Functionality
- **Batch File Processing**: Upload multiple CSV/XLSX files at once
- **Intelligent Column Detection**: Automatically identifies and maps attendance data columns
- **Flexible Date Range Analysis**: Analyze attendance for any month and year
- **Holiday Configuration**: Adjust calculations based on company holidays
- **Department Grouping**: Organize attendance data by department for better insights

### Data Processing
- **Working Days Calculation**: Automatically computes expected working days
- **Attendance Metrics**: 
  - Total days present
  - Total days absent
  - Attendance percentage
  - Late arrivals
  - Early departures
- **Department-wise Summaries**: Aggregate data at department and company levels

### Export & Reporting
- **Professional Excel Output**: Formatted workbooks ready for distribution
- **Data Preview**: View summarized data before exporting
- **Batch Export**: Download reports for multiple date ranges

### User Experience
- **Modern UI**: Clean, intuitive Windows Forms interface
- **Real-time Feedback**: Loading animations and status updates during processing
- **Error Handling**: Comprehensive validation and error messages
- **One-Click Export**: Download Excel reports with a single click

---

## 💻 System Requirements

### Minimum Requirements
- **OS**: Windows 10 or later (64-bit)
- **RAM**: 2 GB
- **Disk Space**: 500 MB (self-contained runtime included)
- **Display**: 1024 x 768 resolution

### Recommended Requirements
- **OS**: Windows 11 (64-bit)
- **RAM**: 4 GB or more
- **Disk Space**: 1 GB
- **.NET Runtime**: 6.0 (bundled with installer)

### Supported Input Formats
- **CSV** (.csv) - Comma-separated values
- **Excel** (.xlsx) - Microsoft Excel 2007 and later

---

## 📥 Installation

### Option 1: Windows Installer (Recommended)
1. Download the latest `AttendanceApp-Setup.exe` from [Releases](https://github.com/DeryChris/attendance-summarizer/releases)
2. Run the installer
3. Follow the installation wizard
4. Launch the application from Start Menu or Desktop

### Option 2: Portable Version
1. Download the published files from [Releases](https://github.com/DeryChris/attendance-summarizer/releases)
2. Extract to your desired location
3. Run `AttendanceApp.exe`

### Option 3: Build from Source
```powershell
# Clone the repository
git clone https://github.com/DeryChris/attendance-summarizer.git
cd attendance-summarizer

# Build the application
dotnet publish .\AttendanceApp.sln -c Release -r win-x64 --self-contained

# Run the application
.\win_app\bin\Release\net6.0-windows\win-x64\publish\AttendanceApp.exe
```

---

## 🎯 Usage Guide

### Quick Start
1. **Launch** the application
2. **Upload Files**: Drag and drop CSV/XLSX files or click "Browse" to select files
3. **Configure**: Select the month, year, and enter holiday count
4. **Analyze**: Click "Analyze & Generate" to process the data
5. **Download**: Once complete, click "Download Excel" to save the report

### Detailed Steps

#### Step 1: Upload Attendance Files
- Click the blue **Browse** button or drag files into the upload area
- Supported formats: CSV and XLSX
- You can upload multiple files at once
- Files are displayed as chips with a close button (✕) to remove them

#### Step 2: Configure Analysis Parameters
| Parameter | Description | Default |
|-----------|-------------|---------|
| **Month** | The month to analyze | Current month |
| **Year** | The year of analysis | Current year |
| **Holiday Count** | Number of holidays in the month | 0 |

#### Step 3: Run Analysis
- Click **"Analyze & Generate"** button
- A loading spinner appears in the Preview box
- Status bar shows "Processing..." in orange
- Processing typically completes in a few seconds

#### Step 4: Review Preview
- Once complete, the Preview section shows a sample of the summarized data
- Review columns: Person ID, Name, Department, Days Present, Days Absent, Attendance %, etc.

#### Step 5: Download Report
- Click **"Download Excel"** to save the full report
- Choose a location to save the file
- File format: `MONTH_summary.xlsx` (e.g., `JANUARY_summary.xlsx`)

### Excel Report Structure
The generated Excel workbook contains:
- **Summary Sheet**: Overall attendance statistics
- **Department Sheets**: Attendance breakdown by department
- **Raw Data Sheet**: Detailed individual records
- **Formatted Charts**: Visual representation of attendance trends

---

## 📋 Input Data Format

### Required Columns (Auto-Detected)
Your CSV/XLSX files should contain the following columns (in any order):

| Column | Description | Examples |
|--------|-------------|----------|
| **Person ID** | Employee/Student identifier | EMP001, ID12345, STU_001 |
| **Name** | Full name | John Doe, Jane Smith |
| **Date** | Attendance date | 2024-01-15, 01/15/2024 |
| **Check-in** | Clock-in time | 08:00, 08:00:00, 8:00 AM |
| **Check-out** | Clock-out time | 17:00, 17:00:00, 5:00 PM |
| **Department** | Department/Section | Sales, IT, HR, Finance |

### Example Input (CSV)
```csv
Person ID,Name,Date,Check-in,Check-out,Department
EMP001,John Doe,2024-01-01,08:00,17:00,Sales
EMP002,Jane Smith,2024-01-01,08:15,17:30,IT
EMP003,Mike Johnson,2024-01-01,,17:00,Finance
```

---

## 🔧 Configuration

### Holiday Management
- Adjust the **Holiday Count** based on your organization's calendar
- Automatically reduces expected working days
- Updates attendance percentage calculations

### Date Range Selection
- Select any month and year from the dropdown
- Analyze historical data or plan future adjustments
- Process multiple months separately for trend analysis

---

## 📊 Output Explanation

### Key Metrics Calculated
- **Working Days**: Total expected working days (excluding weekends and holidays)
- **Days Present**: Days with valid check-in/check-out records
- **Days Absent**: Working days with no attendance record
- **Attendance %**: (Days Present / Working Days) × 100
- **Late Arrivals**: Check-in times after 9:00 AM
- **Early Departures**: Check-out times before 5:00 PM

### Report Customization
The generated Excel file can be further customized:
- Add company logos and branding
- Insert additional analysis
- Create charts and pivot tables
- Share with stakeholders

---

## 🐛 Troubleshooting

### "No attendance data found for the selected month"
- **Cause**: No matching records in uploaded files
- **Solution**: Verify that your files contain data for the selected month and year

### "File could not be unpacked" during installation
- **Cause**: Incomplete download or corrupted file
- **Solution**: 
  1. Delete the installer
  2. Re-download from [Releases](https://github.com/DeryChris/attendance-summarizer/releases)
  3. Run as Administrator

### Column detection errors
- **Cause**: Unexpected column names or formatting
- **Solution**: 
  1. Ensure column headers match standard naming (Person ID, Name, Date, etc.)
  2. Remove extra spaces or special characters from headers
  3. Verify date and time formats are consistent

### Application crashes
- **Cause**: Large file sizes or system memory issues
- **Solution**:
  1. Close other applications to free memory
  2. Split large files into smaller chunks
  3. Report the issue on [GitHub Issues](https://github.com/DeryChris/attendance-summarizer/issues)

---

## 🤝 Contributing

Contributions are welcome! To contribute:

1. **Fork** the repository
2. **Create** a feature branch (`git checkout -b feature/amazing-feature`)
3. **Commit** your changes (`git commit -m 'Add amazing feature'`)
4. **Push** to the branch (`git push origin feature/amazing-feature`)
5. **Open** a Pull Request

### Development Setup
```powershell
git clone https://github.com/DeryChris/attendance-summarizer.git
cd attendance-summarizer
dotnet restore
dotnet build
```

### Reporting Issues
Found a bug? [Create an Issue](https://github.com/DeryChris/attendance-summarizer/issues) with:
- Detailed description of the problem
- Steps to reproduce
- Expected vs. actual behavior
- Sample data (if applicable)
- System information

---

## 📜 License

This project is licensed under the MIT License - see [LICENSE](LICENSE) file for details.

---

## 📧 Support & Contact

Have questions or need assistance? Reach out:

- **Email**: [derychrispen72@gmail.com](mailto:derychrispen72@gmail.com)
- **Phone**: +233 55 0722 898
- **GitHub Issues**: [Report a Problem](https://github.com/DeryChris/attendance-summarizer/issues)
- **GitHub Discussions**: [Ask a Question](https://github.com/DeryChris/attendance-summarizer/discussions)

---

## 🎓 Use Cases

### Human Resources
- Monthly attendance report generation
- Attendance pattern analysis
- Performance evaluation support
- Compliance documentation

### Educational Institutions
- Student attendance tracking
- Class attendance reports
- Enrollment verification
- Scholarship eligibility assessment

### Corporate Organizations
- Employee attendance management
- Shift-based attendance analysis
- Multi-site consolidation
- Payroll support documentation

### Field Operations
- On-site worker attendance
- Project-based time tracking
- Remote work documentation
- Contract worker compliance

---

## 🚀 Roadmap

Future enhancements planned:
- [ ] Automatic updates from GitHub Releases
- [ ] Database integration (SQL Server, MySQL)
- [ ] Real-time analytics dashboard
- [ ] Biometric system integration
- [ ] Email report delivery
- [ ] Custom report templates
- [ ] Multi-language support
- [ ] Cloud synchronization

---

## 📈 Performance

- **Processing Speed**: Analyzes 10,000+ records per second
- **Memory Usage**: Minimal footprint (~200 MB)
- **File Size Limit**: Handles files up to 1 GB
- **Concurrent Processing**: Single-threaded for stability

---

## 🙏 Acknowledgments

Built with:
- [.NET 6.0](https://dotnet.microsoft.com/)
- [EPPlus](https://github.com/EPPlusSoftware/EPPlus) - Excel file generation
- [Windows Forms](https://docs.microsoft.com/en-us/dotnet/desktop/winforms/) - UI Framework

---

## 📌 Version Information

- **Current Version**: 1.0.0
- **Release Date**: 2024
- **Last Updated**: January 2024
- **Status**: Stable Release

---

**Made with ❤️ by [Chrispen Dery](https://github.com/DeryChris)**

[⬆ Back to Top](#attendance-summarizer)
