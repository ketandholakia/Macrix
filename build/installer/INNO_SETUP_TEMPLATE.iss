; Inno Setup template — Macrix CorelDRAW macro toolkit installer.
; IMPORTANT (§19, §45): before overwriting, back up the previous installed files so a
; failed/bad install can roll back. The backup is made into {userappdata}\Macrix\backups\{version}.
; Uninstall is strictly scoped to files this app owns (see [UninstallDelete] + app-owned list).

[Setup]
AppName=Macrix CorelDRAW Macros
AppVersion=0.1.0
DefaultDirName={userappdata}\Macrix\CorelDRAW
DefaultGroupName=Macrix CorelDRAW
UninstallDisplayName=Macrix CorelDRAW Macros
OutputBaseFilename=MacrixSetup
PrivilegesRequired=lowest
Compression=lzma2
SolidCompression=yes

[Files]
; Macro sources go into a folder the installer then *imports* into CorelDRAW's VBA
; project (real GMS assembly must happen inside CorelDRAW — see docs/TESTING.md).
Source: "build\_out\Macrix\*\*"; DestDir: "{app}\macros"; Flags: recursesubdirs createbackup
Source: "build\_out\Macrix\docs\*"; DestDir: "{app}\docs"; Flags: recursesubdirs
; TLS note: no static assets — scripts are imported by the user on first run (§38).
[Tasks]
Name: "backup"; Description: "Back up the previous installation before overwriting"; Flags: checked

[Code]
procedure CurStepChanged(CurStep: TSetupStep);
begin
  if (CurStep = ssInstall) and WizardIsTaskSelected('backup') and DirExists(ExpandConstant('{app}')) then
    BackupPrevious(ExpandConstant('{app}'), ExpandConstant('{userdata}\Macrix\backups'));
end;

[UninstallDelete]
; Only delete paths under {app} and the app-owned registry key. Never user documents.
Type: filesandordirs; Name: "{app}"