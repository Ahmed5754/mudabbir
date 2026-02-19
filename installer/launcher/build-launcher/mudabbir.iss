; Mudabbir Windows Installer — Inno Setup Script
; Reads version from MUDABBIR_VERSION env var (default "0.1.0").
;
; Usage:
;   set MUDABBIR_VERSION=0.3.0
;   "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" mudabbir.iss
;
; Or from CI:
;   iscc /DVERSION=0.3.0 mudabbir.iss

#ifndef VERSION
  #define VERSION GetEnv("MUDABBIR_VERSION")
  #if VERSION == ""
    #define VERSION "0.1.0"
  #endif
#endif

[Setup]
AppId={{B7E3F4A2-8C1D-4E5F-9A0B-6D2C7E8F1A3B}
AppName=Mudabbir
AppVersion={#VERSION}
AppVerName=Mudabbir {#VERSION}
AppPublisher=Mudabbir
AppPublisherURL=https://github.com/Ahmed5754/Mudabbir
DefaultDirName={localappdata}\Mudabbir
DefaultGroupName=Mudabbir
OutputDir=..\..\..\dist\launcher
OutputBaseFilename=Mudabbir-Setup
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=lowest
SetupIconFile=..\assets\icon.ico
UninstallDisplayIcon={app}\Mudabbir.exe
WizardStyle=modern
DisableProgramGroupPage=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "..\..\..\dist\launcher\Mudabbir\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Mudabbir"; Filename: "{app}\Mudabbir.exe"
Name: "{autodesktop}\Mudabbir"; Filename: "{app}\Mudabbir.exe"; Tasks: desktopicon
Name: "{userstartup}\Mudabbir"; Filename: "{app}\Mudabbir.exe"; Tasks: autostart

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"
Name: "autostart"; Description: "Start Mudabbir when Windows starts"; GroupDescription: "Startup:"

[Run]
Filename: "{app}\Mudabbir.exe"; Description: "Launch Mudabbir"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}"

[Code]
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  ConfigDir: String;
  Res: Integer;
begin
  if CurUninstallStep = usPostUninstall then
  begin
    ConfigDir := ExpandConstant('{userappdata}') + '\.mudabbir';
    // Use USERPROFILE for ~/.mudabbir on Windows
    ConfigDir := GetEnv('USERPROFILE') + '\.mudabbir';
    if DirExists(ConfigDir) then
    begin
      Res := MsgBox(
        'Do you want to remove Mudabbir configuration and data?' + #13#10 +
        '(Located at: ' + ConfigDir + ')' + #13#10#13#10 +
        'Click Yes to remove everything, No to keep your data.',
        mbConfirmation, MB_YESNO or MB_DEFBUTTON2
      );
      if Res = IDYES then
      begin
        DelTree(ConfigDir, True, True, True);
      end;
    end;
  end;
end;
