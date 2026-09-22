# Macrix CorelDRAW Macros

This repository contains a collection of custom VBA macros for CorelDRAW, developed to enhance workflow and automate tasks.

## Repository Structure

To keep things organized and maintainable, each macro exists in its own isolated subfolder. This allows you to manage, sync, and develop multiple macros independently within the same repository.

```text
macrixcdrMacro/
│
├── macrixTools/       # Macrix Dimension Tools (Current Macro)
│   ├── src/               # VBA source files (.bas, .cls)
│   ├── forms/             # UserForm files (.frm, .frx)
│   ├── scripts/           # Sync scripts specific to this macro
│   └── GMS/               # Compiled CorelDRAW macro files
│
└── [future_macro]/        # Folder for your next macro project
    ├── src/
    ├── scripts/
    └── ...
```

## How to Sync Macros

Because the CorelDRAW VBA Editor stores code inside proprietary `.gms` files, we use PowerShell scripts to bridge the gap between CorelDRAW and Git. 

Each macro folder contains a `scripts` directory with two important PowerShell scripts:

- **`Sync-From-GMS.ps1`**: Extracts all modules, classes, and forms out of your active CorelDRAW `.gms` project and saves them as plain text files into the `src` and `forms` folders. Run this **before you commit** to Git.
- **`Sync-To-GMS.ps1`**: Takes the plain text files from your `src` and `forms` folders and injects them back into your CorelDRAW `.gms` project. Run this when you pull new updates from Git and want them to appear in CorelDRAW.

## Included Macros

### 1. Macrix Dimension Tools (`/macrixTools`)
A powerful, UI-driven dimensioning tool that automatically measures selected shapes and places formatted dimension labels (width, height, area, perimeter, and object count). Includes customizable padding, backgrounds, and a dynamic template engine for text formatting.

---
*Developed by Ketan for the Macrix workflow.*
