; Crisco - Програма для ведення замін
; Inno Setup Script для створення інсталятора

#define MyAppName "Crisco - Програма для ведення замін"
#define MyAppVersion "1.0"
#define MyAppPublisher "Crisco Development Team"
#define MyAppExeName "Crisco_Optimized.exe"

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
AppId={{B5E9F4A2-8C3D-4E7B-9A1F-2D6C8B9E4F3A}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\Crisco
DefaultGroupName=Crisco
AllowNoIcons=yes
OutputDir=installer_output
OutputBaseFilename=Crisco_Setup
SetupIconFile=
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64
UninstallDisplayIcon={app}\{#MyAppExeName}

[Languages]
Name: "ukrainian"; MessagesFile: "compiler:Languages\Ukrainian.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "dist\Crisco_Optimized.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "dataBase.mdb"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "ІНСТРУКЦІЯ_ВСТАНОВЛЕННЯ.txt"; DestDir: "{app}"; Flags: ignoreversion

[Dirs]
Name: "{app}\Zaminy"; Permissions: users-modify

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}\Zaminy"
Type: filesandordirs; Name: "{app}\__pycache__"

[Code]
function InitializeSetup(): Boolean;
var
  ResultCode: Integer;
  DriverInstalled: Boolean;
begin
  Result := True;
  
  // Перевірка наявності Microsoft Access Database Engine
  DriverInstalled := RegKeyExists(HKEY_LOCAL_MACHINE, 'SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\Access Connectivity Engine');
  
  if not DriverInstalled then
  begin
    if MsgBox('Для роботи програми потрібен Microsoft Access Database Engine.' + #13#10 + 
              'Після встановлення програми, будь ласка, завантажте його з:' + #13#10 + 
              'https://www.microsoft.com/en-us/download/details.aspx?id=54920' + #13#10#13#10 +
              'Продовжити встановлення?', mbConfirmation, MB_YESNO) = IDYES then
    begin
      Result := True;
    end
    else
    begin
      Result := False;
    end;
  end;
end;
