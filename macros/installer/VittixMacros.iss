; ============================================================================
;  Vittix CorelDRAW Macro Suite — Installer
; ============================================================================
;  Built with Inno Setup 6.x (+ ISPP, bundled by default).
;
;  HOW TO ADD A NEW MACRO IN THE FUTURE — 3 touch points, all marked below:
;    1. [Components]  — one line
;    2. [Files]        — one 4-line block (src + forms + both sync scripts)
;    3. [Code]
[Setup]
AppId={{B3C1E9C4-7B77-4B1E-9A6F-1D3B6D8F9C10}}
AppName=Vittix CorelDRAW Macro Suite
AppVersion=1.0.0
AppPublisher=Vittix
AppPublisherURL=https://github.com/ketandholakia/Vittix-CDR-Macro
DefaultDirName={localappdata}\Vittix\CDRMacros
DefaultGroupName=Vittix CDR Macros
DisableProgramGroupPage=yes
OutputDir=..\dist
OutputBaseFilename=VittixCDRMacros-Setup-1.0.0
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Components]
Name: "macro"; Description: "Vittix Macros"; Types: full custom; Flags: fixed
Name: "macro\selectsame"; Description: "VittixSelectSame - select shapes with matching properties"; Types: full custom
Name: "macro\bleed"; Description: "vittixBleed - bleed / crop-mark expansion"; Types: full custom
Name: "macro\dimension"; Description: "VittixDimensionTools - dimension labeling"; Types: full custom
Name: "macro\imposition"; Description: "vittixImposition - imposition/signature layout"; Types: full custom

[Types]
Name: "full"; Description: "Install all macros"
Name: "custom"; Description: "Choose which macros to install"; Flags: iscustom

[Files]
Source: "..\VittixSelectSame\src\*"; DestDir: "{app}\VittixSelectSame\src"; Components: macro\selectsame; Flags: recursesubdirs ignoreversion
Source: "..\VittixSelectSame\forms\*"; DestDir: "{app}\VittixSelectSame\forms"; Components: macro\selectsame; Flags: recursesubdirs ignoreversion
Source: "..\VittixSelectSame\scripts\Sync-To-GMS.ps1"; DestDir: "{app}\VittixSelectSame\scripts"; Components: macro\selectsame; Flags: ignoreversion
Source: "..\VittixSelectSame\scripts\Sync-From-GMS.ps1"; DestDir: "{app}\VittixSelectSame\scripts"; Components: macro\selectsame; Flags: ignoreversion

Source: "..\vittixBleed\src\*"; DestDir: "{app}\vittixBleed\src"; Components: macro\bleed; Flags: recursesubdirs ignoreversion
Source: "..\vittixBleed\forms\*"; DestDir: "{app}\vittixBleed\forms"; Components: macro\bleed; Flags: recursesubdirs ignoreversion
Source: "..\vittixBleed\scripts\Sync-To-GMS.ps1"; DestDir: "{app}\vittixBleed\scripts"; Components: macro\bleed; Flags: ignoreversion
Source: "..\vittixBleed\scripts\Sync-From-GMS.ps1"; DestDir: "{app}\vittixBleed\scripts"; Components: macro\bleed; Flags: ignoreversion

Source: "..\vittixdimension\src\*"; DestDir: "{app}\vittixdimension\src"; Components: macro\dimension; Flags: recursesubdirs ignoreversion
Source: "..\vittixdimension\forms\*"; DestDir: "{app}\vittixdimension\forms"; Components: macro\dimension; Flags: recursesubdirs ignoreversion
Source: "..\vittixdimension\scripts\Sync-From-GMS.ps1"; DestDir: "{app}\vittixdimension\scripts"; Components: macro\dimension; Flags: ignoreversion
Source: "..\vittixdimension\scripts\Sync-To-GMS.ps1"; DestDir: "{app}\vittixdimension\scripts"; Components: macro\dimension; Flags: ignoreversion

Source: "..\vittixImposition\src\*"; DestDir: "{app}\vittixImposition\src"; Components: macro\imposition; Flags: recursesubdirs ignoreversion
Source: "..\vittixImposition\forms\*"; DestDir: "{app}\vittixImposition\forms"; Components: macro\imposition; Flags: recursesubdirs ignoreversion
Source: "..\vittixImposition\scripts\Sync-To-GMS.ps1"; DestDir: "{app}\vittixImposition\scripts"; Components: macro\imposition; Flags: ignoreversion
Source: "..\vittixImposition\scripts\Sync-From-GMS.ps1"; DestDir: "{app}\vittixImposition\scripts"; Components: macro\imposition; Flags: ignoreversion

[Icons]
Name: "{group}\Vittix CDR Macros - Source Folder"; Filename: "{app}"
Name: "{group}\Uninstall Vittix CDR Macros"; Filename: "{uninstallexe}"

[Code]
// ==========================================================================
// [3] MACRO LIST - drives the post-install "import into CorelDRAW" step.
// One record per macro: Component name (matches [Components] above),
// friendly Global Macro Project name (must match exactly what you name the
// saved GMS project inside CorelDRAW), and the sync script path relative to
// {app} (e.g. 'vittixBleed\scripts\Sync-To-GMS.ps1').
//
// ADD NEW MACROS BY ADDING ONE LINE TO SetArrayLength / the assignments below.
// ==========================================================================
const
  MacroCount = 4;

var
  MacroComponentNames: array of String;
  MacroProjectNames: array of String;
  MacroSyncScripts: array of String;

procedure InitMacroList;
begin
  SetArrayLength(MacroComponentNames, MacroCount);
  SetArrayLength(MacroProjectNames, MacroCount);
  SetArrayLength(MacroSyncScripts, MacroCount);

  MacroComponentNames[0] := 'macro\selectsame';
  MacroProjectNames[0]   := 'VittixSelectSame';
  MacroSyncScripts[0]    := 'VittixSelectSame\scripts\Sync-To-GMS.ps1';

  MacroComponentNames[1] := 'macro\bleed';
  MacroProjectNames[1]   := 'vittixBleed';
  MacroSyncScripts[1]    := 'vittixBleed\scripts\Sync-To-GMS.ps1';

  MacroComponentNames[2]   := 'macro\dimension';
  MacroProjectNames[2]   := 'VittixDimensionTools';
  MacroSyncScripts[2]   := 'vittixdimension\scripts\Sync-To-GMS.ps1';

  MacroComponentNames[3]   := 'macro\imposition';
  MacroProjectNames[3]   := 'vittixImposition';
  MacroSyncScripts[3]   := 'vittixImposition\scripts\Sync-To-GMS.ps1';
end;

function MacroScriptsFolder(Index: Integer): String;
var
  Path: String;
begin
  // MacroSyncScripts holds a path relative to {app} (e.g. 'vittixBleed\scripts\Sync-To-GMS.ps1').
  // ExtractFilePath returns the folder with a trailing separator; strip it so the
  // path can be concatenated cleanly with following text.
  Path := ExpandConstant('{app}') + '\' + ExtractFilePath(MacroSyncScripts[Index]);
  if (Length(Path) > 0) and (Path[Length(Path)] = '\') then
    Delete(Path, Length(Path), 1);
  Result := Path;
end;

procedure ImportSelectedMacros;
var
  I: Integer;
  ScriptPath: String;
  ResultCode: Integer;
  Proceed: Boolean;
begin
  for I := 0 to GetArrayLength(MacroComponentNames) - 1 do
  begin
    if WizardIsComponentSelected(MacroComponentNames[I]) then
    begin
      Proceed := MsgBox(
        'Ready to import "' + MacroProjectNames[I] + '" into CorelDRAW.' + #13#10 + #13#10 +
        'Before continuing:' + #13#10 +
        '  1. CorelDRAW must be open.' + #13#10 +
        '  2. An empty Global Macro Project named exactly "' + MacroProjectNames[I] +
             '" must already exist and be saved (Tools > Macros > Manage Macros).' + #13#10 +
        '  3. Macro Security must allow VBA project access.' + #13#10 + #13#10 +
        'Continue with the import now? (Choose No to skip — you can run' + #13#10 +
        '"' + ExtractFileName(MacroSyncScripts[I]) + '" manually from ' + MacroScriptsFolder(I) + ' later.)',
        mbConfirmation, MB_YESNO) = IDYES;

      if Proceed then
      begin
        ScriptPath := ExpandConstant('{app}') + '\' + MacroSyncScripts[I];
        if not Exec('powershell.exe',
             '-NoProfile -ExecutionPolicy Bypass -File "' + ScriptPath + '"',
             '', SW_SHOW, ewWaitUntilTerminated, ResultCode) then
        begin
          MsgBox('Failed to launch PowerShell for ' + MacroProjectNames[I] + '.', mbError, MB_OK);
        end
        else if ResultCode <> 0 then
        begin
          MsgBox(MacroProjectNames[I] + ' import script exited with an error (code ' +
            IntToStr(ResultCode) + '). Check the console output above.', mbError, MB_OK);
        end;
      end;
    end;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    InitMacroList;
    if MsgBox('Files are installed. Import the selected macro(s) into CorelDRAW now?' + #13#10 +
       '(Each import overwrites the matching Global Macro Project''s existing modules/forms.)',
       mbConfirmation, MB_YESNO) = IDYES then
      ImportSelectedMacros;
  end;
end;





