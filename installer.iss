[Setup]
AppId={{8B3E4567-E89B-12D3-A456-426614174000}
AppName=My-Carton
AppVersion=1.0.0
AppPublisher=Dmytro Kliuchko
DefaultDirName={autopf}\My-Carton
DefaultGroupName=My-Carton
OutputDir=deploy\installer
OutputBaseFilename=MyCarton_Setup_v1.0
Compression=lzma
SolidCompression=yes
WizardStyle=modern
; --- ПРАВА АДМІНІСТРАТОРА ---
; Це виправить помилку "Access Denied" при записі бази даних
PrivilegesRequired=admin

[Languages]
Name: "ukrainian"; MessagesFile: "compiler:Languages\Ukrainian.isl"

[Tasks]
; Прибираємо Flags: unchecked, щоб користувач бачив галочку одразу
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
; Основний файл
Source: "deploy\client\moy_carton.exe"; DestDir: "{app}"; Flags: ignoreversion
; --- ДОДАЄМО ІКОНКУ В ПАКЕТ ---
; Переконайтеся, що ваш батник копіює karton.ico у папку deploy\client
Source: "deploy\client\karton.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "deploy\client\*.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "deploy\client\*.config"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Додаємо параметр IconFilename, щоб явно вказати вашу іконку
Name: "{group}\My-Carton"; Filename: "{app}\moy_carton.exe"; IconFilename: "{app}\karton.ico"
Name: "{autodesktop}\My-Carton"; Filename: "{app}\moy_carton.exe"; IconFilename: "{app}\karton.ico"; Tasks: desktopicon

[Run]
; Запуск від імені адміністратора після встановлення
Filename: "{app}\moy_carton.exe"; Description: "{cm:LaunchProgram,My-Carton}"; Flags: nowait postinstall skipifsilent runascurrentuser