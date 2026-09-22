# Vittix Imposition Tool (`/vittixImposition`)

A VBA macro for CorelDRAW that imposes all pages of the active document onto
new output sheets, according to configurable grid/layout settings.

## Features
- **Simple N-up** — sequential page order across a `GRID_ROWS x GRID_COLS` grid.
- **Step & Repeat** — tiles a single source page across every cell/sheet (labels/stickers).
- **Signature (booklet)** — saddle-stitch 2-up page ordering, with automatic
  blank-page padding to a multiple of 4.
- Configurable gutter, margin, bleed, and automatic crop marks.
- Outputs to a brand-new document; the source document is never modified.

## Files

| File | Purpose |
|---|---|
| `src/modSettings.bas` | Public `g_...` settings variables + `InitDefaultSettings` (edit defaults here) |
| `src/modImposition.bas` | Core logic: page ordering, placement, crop marks, `RunImposition`, `ShowImpositionForm` |
| `forms/frmImpositionSettings.frm` | UserForm for entering settings interactively |

## Developer Sync Scripts

If you edit the macro inside CorelDRAW, you can sync changes back to this
repository using the PowerShell scripts in the `scripts` directory:

- **`Sync-From-GMS.ps1`** — Extracts modules and forms from the running CorelDRAW
  instance into the `src` and `forms` directories. Run this **before you commit** to Git.
- **`Sync-To-GMS.ps1`** — Pushes `.bas`, `.cls`, and `.frm` files from `src` and
  `forms` into the CorelDRAW VBA project named `vittixImposition`. The macro must
  already exist as a saved Global Macro Project of that name in CorelDRAW.

## Settings
All settings are `Public` variables in `src/modSettings.bas` (`g_LayoutMode`,
`g_SheetWidth`, `g_GridRows`, `g_GutterX`, `g_BleedSize`, `g_AddCropMarks`,
`g_StepRepeatPage`, etc. — see that file for the full list and defaults).
Both the macro and the form read/write these directly.

## Usage

**With the settings form (recommended):**
1. Open the target document in CorelDRAW.
2. Tools > Macros > Visual Basic Editor.
3. Import `modSettings.bas`, `modImposition.bas`, and `frmImpositionSettings.frm`
   (or sync via the repo's GMS scripts, once added).
4. Run macro `ShowImpositionForm`.
5. Set Layout Mode, sheet size, grid, gutters, margins, bleed, and crop
   marks, then click **Run Imposition**.

**Without the form (edit-and-run):**
1. Edit the defaults in `InitDefaultSettings` (`modSettings.bas`).
2. Run macro `RunImposition` directly — it calls `InitDefaultSettings`
   automatically the first time it's run in a session if the form was
   never opened.

## Notes
- Signature mode currently supports 2-up saddle-stitch only. 4-up/8-up
  signature schemes need a different fold/order formula and aren't
  implemented yet.
- Tested logic assumes CorelDRAW's default `ShapeRange.Copy` / `Document.Paste`
  behavior for moving content between pages; method names may differ slightly
  across CorelDRAW versions (X7 / 2018 / 2021 / 2024) — check the VBA Object
  Browser (F2) if something doesn't resolve.
- The Unit combo box maps to `cdrMillimeter` / `cdrInch` / `cdrPoint` /
  `cdrPixel`. If any of those enum names don't compile in your CorelDRAW
  version, check the exact names via the Object Browser (F2, search `cdrUnit`)
  and adjust `UnitToComboIndex` / `ComboIndexToUnit` in the form's code.
- Recommend testing Signature mode on a small, disposable 8-page test
  document before trusting it on a real job.

*Developed for the Vittix workflow.*
