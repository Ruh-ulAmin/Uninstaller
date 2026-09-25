# Uninstaller

A real, working Windows desktop application for uninstalling programs -
built to replace the ad-hoc PowerShell/batch scripts that used to live in
this repo (still kept under [`legacy/`](legacy) for reference).

It is a WPF (.NET 8) app with a classic Windows look: a top menu bar,
toolbar, sortable/searchable program list, a details pane, and a status
bar - similar in spirit to Control Panel's "Programs and Features", but
with real uninstall reliability features and leftover cleanup on top.

## Features

- **Full program inventory** - reads the same registry locations Windows
  itself uses (`HKLM`/`HKCU`, both 32-bit and 64-bit views) plus Store
  apps (AppX/MSIX) via PowerShell's `Get-AppxPackage`, so the list matches
  what's actually installed.
- **Search, sort, and filter** - live search box, sortable columns, and
  toggles to show/hide system components and Windows updates (hidden by
  default, like the real Control Panel).
- **Reliable uninstall** - runs each program's own uninstaller
  (`UninstallString`/`QuietUninstallString`), rewrites MSI commands for a
  clean silent uninstall when requested, and reports the exit code
  (including "reboot required").
- **Batch uninstall** - check multiple programs and uninstall them all in
  one queued run, with a live progress/log window and per-item
  cancellation between items.
- **Force Remove / leftover cleanup** - when a program's uninstaller is
  missing or broken, shows exactly what it would remove (registry key,
  install folder, shortcuts) and deletes only what you confirm.
- **System-wide Leftover Scanner** - a separate tool that looks for
  orphaned uninstall registry entries and folders under Program Files /
  ProgramData / AppData that no longer belong to any installed program.
  Nothing is ever pre-selected - you choose what to delete.
- **Modify/Repair support**, opening a program's install folder or
  registry key directly, and copying its raw uninstall command.
- **CSV export** of the full program list.
- **Light/Dark/System theme**, a full top menu (File, Edit, View, Actions,
  Tools, Help), toolbar, and status bar.
- **Logging** - every action is recorded under
  `%LOCALAPPDATA%\Uninstaller\Logs` for an audit trail.

## Elevation model

The app launches at the user's normal (non-elevated) integrity level -
its manifest deliberately does **not** request `requireAdministrator`.
Elevation is requested per-action, only when it's actually needed, the
same way Windows' own Programs and Features works:

- Uninstalling, modifying, or force-removing a program whose entry lives
  in `HKEY_LOCAL_MACHINE` (installed for all users) triggers a UAC prompt
  for that one action.
- Per-user programs (`HKEY_CURRENT_USER`) and Store apps run without
  elevation - they don't need it, and never silently inherit it.
- A toolbar/menu action ("File > Restart as Administrator") relaunches the
  whole app elevated if you'd rather not be prompted per-action (e.g. for
  a big batch uninstall or the system-wide Leftover Scanner).

This matters for more than convenience: `HKEY_CURRENT_USER` is writable by
*any* unprivileged process running as the logged-in user, no admin rights
needed. An app that runs fully elevated from launch and blindly executes
whatever `UninstallString`/`ModifyPath` a registry entry contains - HKCU
entries included - turns "click Uninstall" into arbitrary code execution
with administrator rights. Gating elevation by where the entry actually
came from closes that off. See [Security](#security) below for the full
list of hardening measures.

## Security

This app runs with real destructive power over the filesystem and
registry, some of it elevated, so a few things are worth calling out
explicitly:

- **No blanket elevation.** See [Elevation model](#elevation-model) above
  - this is the main one.
- **Recursive deletes are guarded.** Before force-removing a folder
  (`ForceRemovalService`/`PathSafetyGuard`), the app refuses to delete
  anything that isn't at least two path segments below a drive root, or
  that *is* a well-known top-level directory itself (`C:\Windows`,
  `C:\Program Files`, a drive root, etc.) - regardless of what a
  registry value claims an install location is.
- **Registry deletes are scoped.** The app will only ever delete a
  registry key that is a direct child of an `Uninstall` key - never
  anything shallower.
- **Force-remove doesn't pre-select the risky action.** Registry key and
  shortcut removal are pre-checked for convenience; recursive folder
  deletion always starts unchecked, and the confirmation dialog lists the
  actual paths about to be deleted, not just a count.
- **No shell command injection.** Every external process is started via
  argument arrays (`ProcessStartInfo.ArgumentList`), not by building a
  single command string - the classic uninstaller/Store-app removal path
  additionally validates the AppX package identity against an allow-list
  before it's used at all.
- **CSV export is formula-injection-safe.** A `DisplayName` starting with
  `=`, `+`, `-`, `@`, or a tab/CR is prefixed so Excel/Sheets can never
  treat it as a formula.
- **Running processes are detected before deletion**, both for a normal
  uninstall and for force-remove, since a locked file can't be deleted
  anyway and this avoids leaving a program half-removed.
- **Everything destructive requires explicit confirmation** - nothing in
  the Leftover Scanner is pre-selected, and every delete path shows what
  it's about to do before doing it.

If you're auditing this code, the places to look are
`Uninstaller.Core/Services/UninstallService.cs`,
`ForceRemovalService.cs`, `PathSafetyGuard.cs`, and
`LeftoverScannerService.cs`.

## Project layout

```
Uninstaller.sln
src/
  Uninstaller.Core/     Class library: program discovery, uninstall/removal
                        logic, settings, and logging - no UI dependencies.
  Uninstaller.App/      WPF application: MainWindow (menu/toolbar/grid/
                        details/status bar), dialogs, themes, view models.
legacy/                 The original PowerShell/batch scripts, kept for
                        reference only - superseded by the app above.
```

## Building and running

Requires the **.NET 8 SDK** and Windows (the app targets `net8.0-windows`
and uses WPF, Win32 registry APIs, and Windows-only interop).

```powershell
# Restore & build
dotnet build Uninstaller.sln -c Release

# Run (launches unelevated; individual actions prompt for elevation as needed)
dotnet run --project src\Uninstaller.App\Uninstaller.App.csproj -c Release
```

Or open `Uninstaller.sln` in Visual Studio 2022+ and press F5.

### Publishing a standalone executable

```powershell
dotnet publish src\Uninstaller.App\Uninstaller.App.csproj -c Release ^
  -r win-x64 --self-contained true ^
  -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true
```

The resulting `Uninstaller.exe` under
`src\Uninstaller.App\bin\Release\net8.0-windows\win-x64\publish\` can be
copied anywhere and run directly.

## Notes on scope

- The "Leftover Scanner" and general folder/registry matching are
  **heuristic** (name and location based) since Windows keeps no reliable
  mapping from arbitrary leftover files back to a program. Nothing is
  pre-selected for deletion anywhere in the app - you always review and
  confirm before anything is removed.
- Store apps (AppX/MSIX) are enumerated and removed via PowerShell's Appx
  module rather than the Windows.Management.Deployment API, keeping the
  app dependency-free while covering the common case.
