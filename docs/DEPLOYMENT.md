# Deployment: getting source into CorelDRAW

The repository holds plain VBA source (`.bas` / `.frm`). CorelDRAW loads a **`.gms`**
(Global Macro Storage) file, which is a *binary* VBA project — you cannot build one
externally from `.bas` text. So deployment means: get the source into a CorelDRAW VBA
project, then let CorelDRAW persist its `.gms`.

Two supported paths. Use **B** if you are unsure, or if path A reports that the VBA
project object model is not accessible.

---

## Verified facts (CorelDRAW 2021, app version 23.0)

Read from the installed type library
(`Programs64\TypeLibs\CorelDRAW.tlb`), the CorelDRAW `Application` object
(interface `IVGApplication`) exposes:

| Member | Meaning |
| --- | --- |
| `VBE` | The VBIDE project/component object model → import/remove components |
| `GMSManager` | Macro/GMS management: `GMSPath`, `UserGMSPath`, `Projects`, `RunMacro` |
| `InitializeVBA` | VBA initialisation hook |

and `GMSManager` → `GMSProjects` → `GMSProject` gives:

| Member | Meaning |
| --- | --- |
| `Projects.Load(path)` | Load a `.gms` |
| `GMSProject.Unload` | Unload it |
| `GMSProject.Macros` | Enumerate (`GMSMacro.Name` / `.Run` / `.Edit` / `.Delete`) |
| `GMSManager.RunMacro` | Run a macro by name |

This is why the paths below work: CorelDRAW is itself able to load projects and
manipulate VBA components.

> **Prerequisite for path A:** CorelDRAW must permit access to the VBA project object
> model. Without it, `Application.VBE` is unavailable and the script stops with a
> clear message (exit code 3).

---

## Path A — scripted import (`build/deploy.ps1`)

```powershell
# Inspect: what is loaded, and which macros each project exposes (read-only)
powershell -File build/deploy.ps1 -List

# Dry run: validates + prints the Add/Replace plan; writes nothing
powershell -File build/deploy.ps1

# Apply: import every src module into the target project
powershell -File build/deploy.ps1 -Apply

# Target a different project / folder
powershell -File build/deploy.ps1 -Project VittixDimensionTools.gms -Apply
```

Behaviour and safety:
- **Dry run by default.** `-Apply` is required to write anything.
- **Never quits** a CorelDRAW instance it did not start; use `-KeepCorel` to keep even
  a self-started one.
- Runs `build/validate.ps1` first and aborts on errors (override with `-SkipValidate`).
- Warns if the target project has unsaved edits.
- Exit codes: `0` ok · `2` CorelDRAW not reachable over COM · `3` VBE inaccessible ·
  `4` project not loaded · `5` validation failed.

### Why you may need `-StartCorel`

CorelDRAW only registers itself in the Windows **Running Object Table** when it is
started with the `/Automation` switch (see the `LocalServer32` registration:
`CorelDRW.exe /Automation`). A CorelDRAW you launched normally is therefore *not*
attachable over COM — `deploy.ps1` reports this as exit code `2`.

- Pass **`-StartCorel`** to let the script launch its own `/Automation` instance
  (which it will quit afterwards unless `-KeepCorel` is given).
- Or start CorelDRAW yourself with the `/Automation` switch.
- Or use the **bootstrap path (B)**, which needs no external COM at all.

> If you already have CorelDRAW open with unsaved work, prefer path B — do not have
> two instances contending for the same `.gms` files.

---

## Path B — bootstrap importer (recommended, most reliable)

Runs *inside* CorelDRAW, where VBE access is native.

1. Stage the current source:
   ```powershell
   powershell -File build/deploy.ps1 -Stage
   ```
   → copies `src\*.bas` / `*.frm` into `%APPDATA%\Vittix\staging`.

2. **One time:** in the CorelDRAW VBE, import `tools\dev-import\modDevImport.bas`
   into the target project (File → Import File…).

3. Thereafter, whenever you change source: re-run `-Stage`, then in CorelDRAW run
   ```
   modDevImport.DevImport_RunAll
   ```
   It replaces every staged module in the project it lives in and reports
   imported / replaced / failed counts.

`modDevImport.bas` is intentionally **not** part of the staged set, so it never
removes or re-imports itself while running.

---

## Saving

Imported changes live in the running project. CorelDRAW persists a `.gms` when the
project is saved/unloaded or when CorelDRAW closes. Path A attempts
`VBProject.SaveAs` automatically and reports the outcome; if a project refuses the
save, save from the VBE or unload/reload the project.

> ⚠️ Do not edit a `.gms` that is currently running a macro. Keep the VBE closed
> during an import (`deploy.ps1 -Apply` will not close it for you).

---

## Legacy

`build/sync_gms_modules.bat` still works (it copies `.bas`/`.frm` into the GMS folder
for manual import). `deploy.ps1 -Stage` supersedes it and adds validation + VBE import.
