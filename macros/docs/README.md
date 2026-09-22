# Macrix Dimension Tools

CorelDRAW 2021 VBA add-in scaffold for production-style dimension labels.

## Installation
1. Open the repository in VS Code.
2. Run `scripts\Build.ps1`.
3. Import the generated files into CorelDRAW VBA or use the `GMS` package as your source bundle.

## Features
- Modular VBA architecture
- Unit conversion helpers
- Selection measurement helpers
- Label templates and background box scaffold
- Settings persistence scaffold

## Requirements
- CorelDRAW Graphics Suite 2021 v23.3
- VBA enabled in CorelDRAW
- Windows PowerShell

## CorelDRAW Versions
- Tested target: CorelDRAW 2021 v23.3

## Known Issues
- CorelDRAW object-model behavior still needs in-app verification.
- `frmDimension` must be recreated as a real VBA form with the expected control names.

## Roadmap
See `docs/ROADMAP.md` and `docs/TODO.md`.
