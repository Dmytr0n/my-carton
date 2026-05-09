[Setup]
AppId={{8B3E4567-E89B-12D3-A456-426614174000}
AppName=My-Carton
AppVersion=1.0.0
AppPublisher=Dmytro Kliuchko
DefaultDirName={autopf}\My-Carton
DefaultGroupName=My-Carton
; Шлях, куди батник покладе готовий інсталятор
OutputDir=deploy\installer
OutputBaseFilename=MyCarton_Setup_v1.0
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "ukrainian"; MessagesFile: "compiler:Languages\Ukrainian.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Беремо файли з папки, яку згенерував батник (deploy\client)
Source: "deploy\client\moy_carton.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "deploy\client\*.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "deploy\client\*.config"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\My-Carton"; Filename: "{app}\moy_carton.exe"
Name: "{autodesktop}\My-Carton"; Filename: "{app}\moy_carton.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\moy_carton.exe"; Description: "{cm:LaunchProgram,My-Carton}"; Flags: nowait postinstall skipifsilent