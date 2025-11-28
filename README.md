# Attendance Summarizer – Windows C# Application

Attendance data is one of the most sensitive datasets inside a company, yet many HR and operations teams still rely on manual spreadsheets and endless VLOOKUPs to understand who was present, missing, or late. The Attendance Summarizer turns raw CSV/XLSX exports into polished, department-aware insights in a few clicks, eliminating the tedious, error-prone work of stitching together daily logs.

## Problem Statement

> When a company’s attendance data is extracted for analysis—especially when the review must happen per employee across multiple departments—the process becomes exhausting, repetitive, and slow if handled manually. Analysts lose time reconciling formats, leaders wait for answers, and data quality suffers.

The goal of this project is to give HR, operations, and compliance teams an automated, reliable way to transform raw device logs into actionable summaries they can trust.

## Solution Overview

- Automates ingestion of multiple CSV/XLSX attendance exports at once.
- Applies smart column detection so it works with differing templates from biometric devices or HR suites.
- Lets reviewers select the target month, working days, and exceptional holidays before analysis starts.
- Generates a clean Excel workbook segmented by department so leaders can consume the results immediately.
- Highlights anomalies such as missing clock-ins without additional configuration.

The result: a repeatable, auditable pipeline that takes minutes instead of hours and scales with your workforce.

## Key Capabilities

- Multi-file import with drag-and-drop or file picker.
- Month/year selectors plus manual holiday overrides to keep calendars accurate.
- Automatic mapping of core attendance columns (Person ID, Name, Date, Check-in, Check-out, Department) even when headers vary.
- Aggregation logic that consolidates an employee’s records across every department.
- Professionally formatted Excel output featuring:
  - Department banners (purple, bold) for quick scanning.
  - Bordered tabular data prepped for sharing.
  - Computed “Not Clock In” metrics and customizable working-day totals.

## Typical Workflow

1. Launch the application.
2. Click **Select CSV/XLSX Files** (or drag and drop) to load one or many exports.
3. Choose the month and year you want to summarize.
4. Enter the number of company holidays for that period.
5. Hit **Analyze & Generate** to preview the compiled results.
6. Explore the preview grid for spot checks.
7. Click **Download Excel** to save the polished report for stakeholders.

## System Requirements

- Windows 10 or newer (x64)
- .NET 6.0 Runtime (installed automatically with Visual Studio 2022 Community)

## Getting Started

### Option 1: Visual Studio 2022 (Recommended)
1. Install **Visual Studio 2022 Community Edition** with the “.NET desktop development” workload.
2. Open the repository folder (`win_app`) via *File ▸ Open ▸ Folder*.
3. Restore packages: right-click `AttendanceApp.csproj` ▸ **Restore NuGet Packages**.
4. Build: *Build ▸ Build Solution* (`Ctrl+Shift+B`).
5. Run: *Debug ▸ Start Debugging* (`F5`).

### Option 2: .NET CLI
```bash
cd win_app
dotnet restore
dotnet build AttendanceApp.sln -c Release
dotnet run
```

## Publishing a Standalone `.exe`

### Visual Studio Publish Profile
1. Right-click `AttendanceApp.csproj`.
2. Choose **Publish** ▸ Target = **Folder**.
3. Select **Self-contained**, configuration **Release**.
4. Publish to generate a distributable `.exe`.

### Command Line Publish
```bash
dotnet publish AttendanceApp.sln -c Release -r win-x64
```

Output path: `bin\Release\net6.0-windows\win-x64\publish\AttendanceApp.exe`

## Architecture at a Glance

- `Program.cs` – Application entry point.
- `MainForm.cs` – Windows Forms UI, event handling, and user workflow.
- `AttendanceParser.cs` – Parsing pipeline, validation, and aggregation logic.
- `ExcelHelper.cs` – ClosedXML-powered report builder.
- `AttendanceApp.csproj` – Project configuration and dependencies.

## Dependencies

- **ClosedXML** for Excel generation.
- **CsvHelper** for resilient CSV parsing.

Both packages restore automatically via NuGet—no manual setup required.

## Why Teams Choose This App

- **Speed:** Turn around department-ready attendance summaries in minutes.
- **Accuracy:** Enforce consistent business rules every time, reducing spreadsheet mistakes.
- **Scalability:** Supports large data sets and multiple data sources without extra setup.
- **Shareability:** Produces boardroom-friendly Excel workbooks that managers and auditors expect.

If your organization still burns hours reconciling attendance data by hand, the Attendance Summarizer is the fastest path to clarity and confidence.
