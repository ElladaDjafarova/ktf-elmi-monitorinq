; KTF Elmi Monitorinq — Windows quraşdırma faylı (Inno Setup 6)
#define AppName "KTF Elmi Monitorinq"
#ifndef AppVersion
  #define AppVersion "1.0.0"
#endif

[Setup]
AppId={{6F1B6C2E-4A57-4E0B-9C1D-3F2A1B0C9D8E}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher=Kimya texnologiyaları fakültəsi
DefaultDirName={localappdata}\Programs\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=dist
OutputBaseFilename=KTF_Elmi_Monitorinq_Setup
SetupIconFile=logo.ico
UninstallDisplayIcon={app}\ElmiMonitorinq.exe
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ShowLanguageDialog=no

[Languages]
Name: "en"; MessagesFile: "compiler:Default.isl"
Name: "ru"; MessagesFile: "compiler:Languages\Russian.isl"

[Tasks]
Name: "desktopicon"; Description: "Masaüstündə ikon yarat / Create a desktop icon"; Flags: checkedonce

[Files]
Source: "dist\ElmiMonitorinq\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion

[Icons]
Name: "{group}\Elmi Monitorinq"; Filename: "{app}\ElmiMonitorinq.exe"
Name: "{group}\Müəllim siyahısı (Excel)"; Filename: "{userdocs}\{#AppName}\muellimler.xlsx"
Name: "{group}\Proqramı sil"; Filename: "{uninstallexe}"
Name: "{userdesktop}\Elmi Monitorinq"; Filename: "{app}\ElmiMonitorinq.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\ElmiMonitorinq.exe"; Description: "Elmi Monitorinq-i indi aç / Launch now"; Flags: nowait postinstall skipifsilent
