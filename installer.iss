[Setup]
AppId={{8B3E4567-E89B-12D3-A456-426614174000}
AppName=My-Carton
AppVersion=1.0.0
AppPublisher=Dmytro Kliuchko

; --- МОЖЛИВІСТЬ ВИБОРУ ШЛЯХУ ---
; Дозволяємо користувачу змінювати шлях (сторінка вибору буде видимою)
DisableDirPage=no
; Дозволяємо встановлювати в папку, яка не є порожньою (якщо треба)
AppendDefaultDirName=yes

; Пропонуємо за замовчуванням диск C, але користувач зможе вибрати D через "Огляд"
DefaultDirName={commonpf}\My-Carton
; Або, якщо ви хочете ПРЯМО запропонувати диск D за замовчуванням (якщо він є):
; DefaultDirName=D:\My-Carton

DefaultGroupName=My-Carton
OutputDir=deploy\installer
OutputBaseFilename=MyCarton_Setup_v1.1
Compression=lzma
SolidCompression=yes
WizardStyle=modern

; Залишаємо права адміна, бо якщо користувач все ж вибере C:\Program Files, 
; без них база даних (csv) не зможе зберегтися.
PrivilegesRequired=admin

[Languages]
Name: "ukrainian"; MessagesFile: "compiler:Languages\Ukrainian.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "deploy\client\moy_carton.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "deploy\client\karton.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "deploy\client\*.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "deploy\client\*.config"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\My-Carton"; Filename: "{app}\moy_carton.exe"; IconFilename: "{app}\karton.ico"
Name: "{autodesktop}\My-Carton"; Filename: "{app}\moy_carton.exe"; IconFilename: "{app}\karton.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\moy_carton.exe"; Description: "{cm:LaunchProgram,My-Carton}"; Flags: nowait postinstall skipifsilent