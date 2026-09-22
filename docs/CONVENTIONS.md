# Coding conventions

These are the project-wide rules. Enforcement entry point: `build/validate.ps1` (static).

## Naming

| Prefix | Meaning | Examples |
| --- | --- | --- |
| `mod` | Standard module (global service / entry / infra) | `modLogger`, `modConfig` |
| `cls` | Class module (rare; prefer modules in this codebase) | `clsSessionBreaker` |
| `frm` | UserForm (.frm) | `frmAbout` |
| `pl` | Pure logic, **zero** CorelDRAW object references (§40) | `pl_StringUtils` |
| `feat` | End-user feature (registered) | `feat_RoundedCorners` |
| `modStart*` | Entry point(s) | `modStartBootstrap`, `modStartFactory` |

Any new `.bas` must use `Option Explicit`.

## Public member naming
- Public procedures/functions use `PascalCase` with a verb first: `SaveAll`, `ApplyToSelection`.
- Public members are prefixed with a short **module namespace** to avoid collisions across modules:
  e.g. `Logger_Log`, `CFG_Get`, `FR_Register`, `PL_TrimWS`.
- Private members are prefixed `p_` and otherwise PascalCase: `p_cleanup`, `p_releaseRefs`.
  **Do not use a leading underscore** (`_cleanup`) — VBA identifiers must start with a letter, so
  `Private Sub _cleanup()` is a compile error (this bit us once already; every leading-underscore
  identifier in the project had to be renamed — see CHANGELOG). `Private` already scopes the member,
  so `p_` is purely a visual cue, not a language requirement.

## The safe-run pattern (mandatory)
Every CorelDRAW-touching public entry point MUST be a thin entry sub with a single `On Error GoTo
Cleanup` and one cleanup label that runs on both the success and error paths — restoring
`Optimization`/`EventsEnabled`, releasing object references (`modComLifecycle.CM_Release`), and
reporting/counting any failure via `modErrorHandler`. See `docs/ADDING_A_MACRO.md` and
`feat_RoundedCorners.Main` / `feat_ExportText.Main` for the actual pattern in use.

```vb
Public Sub Main()
    Dim app As Object, doc As Object
    On Error GoTo Cleanup
    If modErrorHandler.EH_IsDisabled(F_BreakerId) Then
        Err.Raise 1001, "feat_MyFeature.Main", "feature session-disabled"
    End If
    Set app = modAppServices.SV_App_Host()
    ' ... do the work ...
    modErrorHandler.EH_RecordSuccess F_BreakerId

Cleanup:
    modComLifecycle.CM_Release app
    modComLifecycle.CM_Release doc
    If Err.Number <> 0 Then
        modErrorHandler.EH_ReportError Err.Number, "feat_MyFeature.Main", Err.Description
        If modErrorHandler.EH_RecordFailure(F_BreakerId) Then
            modUiUtilities.UU_MsgWarn "This feature is disabled for this session; see the log."
        End If
    End If
End Sub
```

There is deliberately no single generic `EH_Guard`/`EH_SafeRun` wrapper: VBA's `AddressOf` can only
target a fixed-signature callback (API/host callbacks), not a generic "invoke any Sub with any
parameter list" dispatcher, so a one-size-fits-all wrapper isn't achievable the way that might look
possible in other languages. Each feature inlines the pattern above instead — `modErrorHandler`
supplies the reporting/circuit-breaker pieces, the caller supplies the cleanup label.

## Error handling inside service helpers
Service modules (`modAppServices`, `modShapeServices`, etc.) use `On Error Resume Next` around
individual late-bound `CallByName` probes, since a missing member on an older/newer CorelDRAW version
is expected and should degrade gracefully rather than crash. Be deliberate about where that pattern
stops: a helper swallowing an error internally means the caller's `Err.Number` check at its `Cleanup`
label will never see it, so a genuine failure (wrong member name, unexpected object state) can be
silently reported as success. Reserve `On Error Resume Next` for member-existence probing; let file
I/O and other operations whose failure should actually stop the feature (see
`feat_ExportText.p_WriteDocumentText`) propagate to the caller instead.

## Comments & documentation
- Header comment on every module: purpose, dependencies, CorelDRAW-version assumptions, references.
- `TODO`/`FIXME`/`HACK` are allowed but must carry an ID tied to a tracking ticket.
- User-visible behavior changes → CHANGELOG entry (§43).
- Keep `Depends on:` comments accurate — a stale reference to a module that doesn't exist (or a
  since-renamed member) is worse than no comment; check it whenever you touch a file.
