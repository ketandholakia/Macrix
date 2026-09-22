Legacy: ImpositionMacro.bas

A standalone, single-file version of the imposition macro (edit the constants at the
top and paste the whole file into a module). It defines the same four procedures as
modImposition.bas -- RunImposition, BuildPageOrder, PlacePageInCell and DrawCropMarks
-- so shipping both in one VBA project is an "Ambiguous name detected" compile error.

The modular version (modImposition.bas + modSettings.bas + mdlFormBuilder.bas +
frmImpositionSettings) supersedes it and is what gets deployed. Kept here for reference.
