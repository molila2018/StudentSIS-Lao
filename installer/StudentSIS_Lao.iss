; ═══════════════════════════════════════════════════════════════════════
;  StudentSIS Lao — Windows Installer
;  Requires Inno Setup 6.x: https://jrsoftware.org/isinfo.php
;
;  Build:   double-click  installer\build.cmd
;           (runs `dotnet publish` then compiles this script)
;  Output:  installer-output\StudentSIS_Lao_Setup_v{version}.exe
;
;  Deployment model:
;    • Per-user install — defaults to %LOCALAPPDATA%\Programs\StudentSIS_Lao
;      so teachers without admin rights can install on shared school PCs.
;    • Admin can opt to install for all users at the elevation prompt.
;    • Database (sis_lao.db) is created next to the EXE on first run; the
;      installer NEVER touches it so existing data survives upgrades.
;    • Self-contained .NET 8 build — no separate runtime install needed.
; ═══════════════════════════════════════════════════════════════════════

; ── Version ──
;     Preferred: pass in from the command line so csproj is the single source
;     of truth — build.cmd extracts <Version> from StudentSIS.csproj and calls
;     ISCC.exe /DMyAppVersion=<v>. The fallback below is only used when this
;     .iss is compiled by hand (drag-and-drop onto Inno's Compil32.exe) without
;     that /D switch — update it only in that case, otherwise leave alone.
#ifndef MyAppVersion
  #define MyAppVersion        "1.0.0"
#endif

#define MyAppName             "StudentSIS Lao"
#define MyAppNameLao          "ລະບົບຄຸ້ມຄອງຂໍ້ມູນນັກຮຽນ"
#define MyAppPublisher        "Buekthong Secondary School (ມສ ບຶກທົ່ງ)"
#define MyAppURL              "https://example.com/studentsis-lao"
#define MyAppExeName          "StudentSIS_Lao.exe"
#define MyAppId               "{{B7F9D2E1-3A4C-4D5E-9F6A-1B2C3D4E5F6A}"

; Path (relative to this .iss file) where `dotnet publish` placed its output.
; build.cmd sets this up at <project root>\publish\ for us.
#define PublishDir            "..\publish"

[Setup]
AppId={#MyAppId}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
VersionInfoVersion={#MyAppVersion}
VersionInfoProductName={#MyAppName}
VersionInfoCompany={#MyAppPublisher}

; ── Install location ──
; Per-user by default → no UAC prompt on standard accounts.
; PrivilegesRequiredOverridesAllowed=dialog lets the user choose at install time
; (per-user vs all-users) via Inno's standard elevation dialog.
DefaultDirName={autopf}\StudentSIS_Lao
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
UsePreviousAppDir=yes
UsePreviousGroup=yes
UsePreviousTasks=yes
UsePreviousLanguage=yes

; ── Wizard UI ──
WizardStyle=modern
WizardResizable=no
ShowLanguageDialog=auto

; ── Compression ──
Compression=lzma2/ultra
SolidCompression=yes
LZMAUseSeparateProcess=yes

; ── Architecture ──
; The app is 64-bit (System.Data.SQLite native interop is x64). `x64` matches
; every Inno Setup 6.x; `x64compatible` (6.3+) also accepts ARM64 emulation.
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64

; ── Output ──
OutputDir=..\installer-output
OutputBaseFilename=StudentSIS_Lao_Setup_v{#MyAppVersion}

; ── Uninstall ──
UninstallDisplayName={#MyAppName} {#MyAppVersion}
UninstallDisplayIcon={app}\{#MyAppExeName}

; ── Setup wizard icon (top-left of the installer window + Add/Remove Programs) ──
; File path is relative to the .iss file.
SetupIconFile=..\Assets\app-icon.ico

; ── Misc ──
; CloseApplications=force tells Inno to close the running app via the Windows
; Restart Manager before overwriting files — needed for clean upgrades when the
; user has the app open. Detection works via the EXE file lock; no app-side
; mutex required.
CloseApplications=force
RestartApplications=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[CustomMessages]
english.DesktopIconTask=Create a &desktop shortcut

[Tasks]
Name: "desktopicon"; Description: "{cm:DesktopIconTask}"; GroupDescription: "Additional shortcuts:"

[Files]
; Everything from the publish folder ends up in {app}.
; Templates\ and Fonts\ (if present) come along automatically because
; the csproj declares them as <None> with CopyToOutputDirectory=PreserveNewest.
;
; recursesubdirs + createallsubdirs preserves the Templates folder layout.
; ignoreversion means files are always overwritten on upgrade — except the
; sis_lao.db database, which never appears in the publish output (created at
; runtime) so it's never touched.
Source: "{#PublishDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}";                          Filename: "{app}\{#MyAppExeName}";   WorkingDir: "{app}"
Name: "{group}\{#MyAppName} (ພາສາລາວ)";               Filename: "{app}\{#MyAppExeName}";   WorkingDir: "{app}";   Comment: "{#MyAppNameLao}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}";   Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}";                    Filename: "{app}\{#MyAppExeName}";   WorkingDir: "{app}";   Tasks: desktopicon

[Run]
; Launch after install — optional check-box on the final wizard page.
Filename: "{app}\{#MyAppExeName}";  \
    Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}";  \
    Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Clean up empty install dir on uninstall. We do NOT delete the database — if the
; school wants to wipe it, they can do so manually. Preserving sis_lao.db on
; uninstall is the safer default: a casual uninstall/reinstall must not lose
; years of student records.
Type: dirifempty; Name: "{app}\Templates"
Type: dirifempty; Name: "{app}"
