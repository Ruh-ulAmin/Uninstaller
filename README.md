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

## Why it needs Administrator

Machine-wide uninstall entries live in `HKEY_LOCAL_MACHINE`, and their
install folders are typically under `Program Files`/`ProgramData` - both
require an elevated token to modify or delete. The app's manifest requests
`requireAdministrator`, so Windows will prompt for elevation on launch. If
you decline, the app still runs and shows a warning banner, but some
uninstalls and cleanups may fail.

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

# Run (will prompt for elevation on launch)
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
copied anywhere and run directly (Windows will still prompt for
elevation, per the embedded manifest).

## Notes on scope

- The "Leftover Scanner" and general folder/registry matching are
  **heuristic** (name and location based) since Windows keeps no reliable
  mapping from arbitrary leftover files back to a program. Nothing is
  pre-selected for deletion anywhere in the app - you always review and
  confirm before anything is removed.
- Store apps (AppX/MSIX) are enumerated and removed via PowerShell's Appx
  module rather than the Windows.Management.Deployment API, keeping the
  app dependency-free while covering the common case.
