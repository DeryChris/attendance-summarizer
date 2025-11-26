; Inno Setup Script for Attendance Summarizer
; Generated for packaging as Windows installer

[Setup]
AppName=Attendance Summarizer
AppVersion=1.0.0
AppPublisher=Chrispen Dery
AppPublisherURL=https://github.com/DeryChris/attendance-summarizer
AppSupportURL=https://github.com/DeryChris/attendance-summarizer/issues
AppUpdatesURL=https://github.com/DeryChris/attendance-summarizer/releases
DefaultDirName={pf}\Attendance Summarizer
DefaultGroupName=Attendance Summarizer
AllowNoIcons=yes
LicenseFile=LICENSE
InfoBeforeFile=
InfoAfterFile=
OutputDir=.\Installer
OutputBaseFilename=AttendanceApp-Setup
Compression=lzma
SolidCompression=yes
PrivilegesRequired=admin
SetupIconFile=win_app\icon.ico
WizardStyle=modern
UninstallDisplayIcon={app}\AttendanceApp.exe

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Copy all published files from the publish directory
Source: "win_app\bin\Release\net6.0-windows\win-x64\publish\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "win_app\icon.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Attendance Summarizer"; Filename: "{app}\AttendanceApp.exe"; IconFileName: "{app}\icon.ico"
Name: "{group}\{cm:UninstallProgram,Attendance Summarizer}"; Filename: "{uninstallexe}"
Name: "{commondesktop}\Attendance Summarizer"; Filename: "{app}\AttendanceApp.exe"; IconFileName: "{app}\icon.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\AttendanceApp.exe"; Description: "{cm:LaunchProgram,Attendance Summarizer}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: dirifempty; Name: "{app}"
