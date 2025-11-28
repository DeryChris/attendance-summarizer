# Attendance Summarizer - C# Windows App

A native Windows Forms application for processing and summarizing attendance records from CSV/XLSX files.

## System Requirements

- Windows 10 or later (x64)
- .NET 6.0 Runtime (or Visual Studio 2022 Community Edition)

## Building the Application

### Option 1: Using Visual Studio 2022 (Recommended)

1. Download and install **Visual Studio 2022 Community Edition** (free):
   - Include ".NET desktop development" workload during installation
   
2. Open the solution:
   - File > Open > Folder, select the `win_app` folder
   
3. Restore NuGet packages:
   - Right-click on `AttendanceApp.csproj` > Restore NuGet Packages
   
4. Build:
   - Build > Build Solution (Ctrl+Shift+B)
   
5. Run:
   - Debug > Start Debugging (F5)

### Option 2: Using Command Line

```bash
# Navigate to the win_app folder
cd win_app

# Restore dependencies
dotnet restore

# Build the project
dotnet build AttendanceApp.sln -c Release

# Run the app
dotnet run
```

## Publishing as an .EXE

To create a standalone executable that works on any Windows machine without needing .NET installed:

### Using Visual Studio:
1. Right-click on `AttendanceApp.csproj`
2. Select "Publish"
3. Choose "Folder" target
4. Select "Self-contained" and Release
5. Publish - this creates a standalone `.exe` file

### Using Command Line:
```bash
dotnet publish AttendanceApp.sln -c Release -r win-x64
```

The executable will be in: `bin\Release\net6.0-windows\win-x64\publish\AttendanceApp.exe`

## Features

- Upload multiple CSV or XLSX attendance files
- Select month and year for processing
- Override holiday count
- Automatic column detection (Person ID, Name, Date, Check-in, Check-out, Department)
- Aggregates attendance data by person across all departments
- Generates professional Excel summary with:
  - Department headers (purple, bold)
  - Formatted data with borders
  - Automatic "Not Clock In" calculations
  - Customizable working days

## Usage

1. Run the application
2. Click "Select CSV/XLSX Files" and choose your attendance files or Drag and Drop them
3. Select the month and year
4. Enter holiday count (if any)
5. Click "Analyze & Generate" to process
6. Review the preview
7. Click "Download Excel" to save the summary file

## File Structure

- `Program.cs` - Application entry point
- `MainForm.cs` - Main UI window
- `AttendanceParser.cs` - CSV parsing and business logic
- `ExcelHelper.cs` - Excel workbook generation
- `AttendanceApp.csproj` - Project configuration

## Dependencies

- **ClosedXML** - Excel file generation
- **CsvHelper** - CSV parsing

These are automatically installed via NuGet when building.
