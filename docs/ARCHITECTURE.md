# Architecture

Greenfield scaffold of the target architecture (§3 of the master prompt). Adapted for the reality that
this is a **CorelDRAW VBA** project: services are standard modules operating on the host object model,
kept late-bound for cross-version compatibility.

```
CorelDRAW VBA Application
├── Bootstrap / Entry Points        modBootstrap
├── Core Framework
│   ├── Application Services        modAppServices
│   ├── Document  Services          modDocumentServices
│   ├── Selection Services          modSelectionServices
│   ├── Layer Services              modLayerServices
│   ├── Shape Services              modShapeServices
│   ├── Text Services               modTextServices        [planned — not yet built]
│   ├── Color Services              modColorServices        [planned — not yet built]
│   ├── Page Services               modPageServices
│   └── Export/Import Services      modExportImportServices
├── Infrastructure
│   ├── Configuration               modConfig
│   ├── Logging                     modLogger
│   ├── Error Handling              modErrorHandler   (+ session circuit breaker)
│   ├── Versioning                  modVersion
│   ├── Environment Detection         modEnvironment
│   ├── Path Management              modPathManager
│   └── COM Lifecycle                 modComLifecycle
├── UI
│   ├── UserForms                     frmAbout
│   ├── Dialog Helpers                modDialogHelper
│   └── UI Utilities                  modUiUtilities
├── Features
│   ├── feat_RoundedCorners
│   ├── feat_ExportText
│   └── …
├── Compatibility
│   ├── modVersionDetect
│   └── modCorelCompatibility
├── Security
│   ├── Signing / Trust               modTrust
│   └── Reference Integrity           modReferenceRegistry
└── Development / Build System
    ├── Validation                    build/validate.ps1
    ├── Packaging                     build/package.ps1
    ├── Installation                  build/installer/INNO_SETUP_TEMPLATE.iss
    └── Documentation                 docs/*
```

`modTextServices` and `modColorServices` are named here as target-state modules but don't exist yet —
`feat_ExportText` currently reads shape text directly via `modBinding.BND_HasMember`/`CallByName`
rather than through a dedicated service module. Extract that into `modTextServices` once a second
feature needs the same text access, rather than before there's a second caller to generalize from.

## Conventions that matter here

- **Late-bound (default):** services use `Object` variables and `CallByName` where API surface may
  differ by CorelDRAW version, per §42. Early binding (where chosen during development) must be
  declared in `modReferenceRegistry` with its `Name + GUID + Version` and a runtime detection path.
- **Single exit + paired restore:** every procedure that changes `Application.Optimization`,
  `EventsEnabled`, command groups, or transient references restores state on **both** the success path
  and the error path (a shared cleanup label) — §14, §39.
- **COM discipline:** §37 — release `ShapeRange`/`Selection`/`Layer`/`Page` refs inside loops; never
  cache a shape reference across invalidating operations (undo/delete/ungroup).
- **Feature gating:** features are discovered via `modFeatureRegistry`, not by hard-coded menus (§12).
- **Session circuit breaker:** §46 — a feature failing 3× in one session auto-disables until next
  launch, logged clearly.

## Not-yet-resolved (owner inputs, §48)

- [ ] Which CorelDRAW version(s) will the *first* supported release target?
- [ ] Is the VBA project digitally signed (and with what cert)? Falses for the trust/installer copies.
- [ ] Is a Windows CI runner available, or is all validation manual?
- [ ] Final installer technology selection (Inno template provided as the default).