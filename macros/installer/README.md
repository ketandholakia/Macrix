# Macrix CDR Macros — Installer

Inno Setup script (`MacrixMacros.iss`) that packages all macros in this repo
into one Windows installer with a checkbox per macro, and optionally drives
the existing `Sync-*-To-GMS.ps1` scripts to import the chosen macro(s)
straight into a running CorelDRAW instance.

## Where this lives in the repo

```
macrixcdrMacro/
├── installer/
│   ├── MacrixMacros.iss   ← this script
│   └── README.md           ← this file
├── MacrixSelectSame/
├── macrixBleed/
├── scripts/
└── dist/                   ← build output goes here (gitignore this)
```

`{#RepoRoot}` in the script is `..\`, i.e. relative to `installer/`. If you
move the `.iss` file, update that one line.

## Building

1. Install [Inno Setup 6](https://jrsoftware.org/isinfo.php) (free). ISPP is
   bundled — no separate install needed.
2. Open `MacrixMacros.iss` in the Inno Setup IDE (or `iscc.exe MacrixMacros.iss`
   from the command line).
3. Compile. Output lands in `dist\MacrixCDRMacros-Setup-<version>.exe`.

## Adding a new macro later

Three places to touch, all marked with `--- ADD NEW ... HERE ---` comments
in the script:

1. **`[Components]`** — one line:
   ```
   Name: "macro\dimension"; Description: "MacrixTools — dimensioning tools"; Types: full custom
   ```

2. **`[Files]`** — one 4-line block (source, forms, and both sync scripts):
   ```
   Source: "{#RepoRoot}MacrixTools\src\*"; DestDir: "{app}\MacrixTools\src"; Components: macro\dimension; Flags: recursesubdirs ignoreversion
   Source: "{#RepoRoot}MacrixTools\forms\*"; DestDir: "{app}\MacrixTools\forms"; Components: macro\dimension; Flags: recursesubdirs ignoreversion
   Source: "{#RepoRoot}scripts\Sync-Dimension-To-GMS.ps1"; DestDir: "{app}\scripts"; Components: macro\dimension; Flags: ignoreversion
   Source: "{#RepoRoot}scripts\Sync-Dimension-From-GMS.ps1"; DestDir: "{app}\scripts"; Components: macro\dimension; Flags: ignoreversion
   ```

3. **`[Code]` → `InitMacroList`** — bump `SetArrayLength(MacroList, N)` and
   add one row:
   ```pascal
   MacroList[2].ComponentName := 'macro\dimension';
   MacroList[2].ProjectName   := 'MacrixTools';  // must match the GMS project name exactly
   MacroList[2].SyncScript    := 'Sync-Dimension-To-GMS.ps1';
   ```

No other logic in the script needs to change — the pre-flight messages,
the `powershell.exe` invocation, and error handling are all driven off this
array.

## What the installer does NOT do

- **It doesn't create the Global Macro Project inside CorelDRAW.** Per the
  existing sync scripts' own requirement, an empty GMS project with the
  exact matching name must already exist and be saved in CorelDRAW before
  the import step runs. The installer's pre-flight dialog reminds the user
  of this each time — it can't create it for them (no documented CorelDRAW
  API for "create + save a new Global Macro Project" from outside the app).
- **It doesn't touch CorelDRAW's macro security settings.** If those block
  VBA project access, the sync script will fail with the same error it
  already gives when run manually — the installer just surfaces that.
- **Uninstall only removes the copied files** under `{app}` — it does not
  reach into CorelDRAW's live GMS project to remove imported modules. If
  full uninstall-from-CorelDRAW is wanted later, that would need a
  `Sync-*-Remove-From-GMS.ps1` script analogous to the existing ones, called
  from an `[UninstallRun]` section.

## Notes

- `PrivilegesRequired=lowest` is intentional — running the installer
  elevated while CorelDRAW runs as a normal user breaks the COM
  `GetActiveObject("CorelDRAW.Application")` call the sync scripts rely on
  (different session). Don't "Run as administrator" this installer.
- The installed layout mirrors the repo layout (`{app}\<MacroName>\src`,
  `{app}\<MacroName>\forms`, `{app}\<MacroName>\scripts\...`) on purpose, so
  each macro's `Sync-*-To-GMS.ps1` / `Sync-*-From-GMS.ps1` scripts work
  unmodified post-install — they resolve paths relative to
  `$PSScriptRoot\..` (i.e. the macro's own folder).
