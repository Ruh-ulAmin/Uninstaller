using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Input;
using Microsoft.Win32;
using Uninstaller.App.ViewModels;
using Uninstaller.App.Views;
using Uninstaller.Core.Models;
using Uninstaller.Core.Services;

namespace Uninstaller.App;

public partial class MainWindow : Window
{
    private readonly MainViewModel _viewModel = new();

    public MainWindow()
    {
        InitializeComponent();
        DataContext = _viewModel;
        Loaded += async (_, _) => await _viewModel.RefreshAsync();
        PreviewKeyDown += MainWindow_PreviewKeyDown;
    }

    private async void MainWindow_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        var ctrl = (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control;

        if (ctrl && e.Key == Key.R)
        {
            await _viewModel.RefreshAsync();
            e.Handled = true;
        }
        else if (ctrl && e.Key == Key.E)
        {
            ExportProgramList();
            e.Handled = true;
        }
        else if (ctrl && e.Key == Key.A)
        {
            SelectAllChecked(true);
            e.Handled = true;
        }
        else if (ctrl && e.Key == Key.F)
        {
            FocusSearchBox();
            e.Handled = true;
        }
        else if (e.Key == Key.Delete && ProgramsGrid.IsKeyboardFocusWithin)
        {
            await UninstallSelectedAsync(silent: false);
            e.Handled = true;
        }
    }

    // ===================== File menu =====================

    private async void RefreshMenuItem_Click(object sender, RoutedEventArgs e) => await _viewModel.RefreshAsync();

    private void ExportMenuItem_Click(object sender, RoutedEventArgs e) => ExportProgramList();

    private void ExportProgramList()
    {
        var dialog = new SaveFileDialog
        {
            Filter = "CSV file (*.csv)|*.csv",
            FileName = $"InstalledPrograms_{DateTime.Now:yyyy-MM-dd}.csv"
        };

        if (dialog.ShowDialog(this) == true)
        {
            try
            {
                _viewModel.ExportListToCsv(dialog.FileName);
                _viewModel.StatusText = $"Exported program list to {dialog.FileName}";
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(this, $"Failed to export: {ex.Message}", "Uninstaller",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    private void OpenLogFolderMenuItem_Click(object sender, RoutedEventArgs e) =>
        MainViewModel.OpenFolder(_viewModel.Logger.LogFolder);

    private void ExitMenuItem_Click(object sender, RoutedEventArgs e) => System.Windows.Application.Current.Shutdown();

    private void RestartAsAdminButton_Click(object sender, RoutedEventArgs e)
    {
        var exePath = Process.GetCurrentProcess().MainModule?.FileName;
        if (string.IsNullOrWhiteSpace(exePath))
        {
            System.Windows.MessageBox.Show(this, "Could not determine the application path.", "Uninstaller",
                MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = exePath,
                UseShellExecute = true,
                Verb = "runas"
            });
            System.Windows.Application.Current.Shutdown();
        }
        catch (Win32Exception ex) when (ex.NativeErrorCode == 1223)
        {
            // User declined the UAC prompt - stay open exactly as before.
        }
    }

    // ===================== Edit menu =====================

    private void SelectAllMenuItem_Click(object sender, RoutedEventArgs e) => SelectAllChecked(true);

    private void SelectNoneMenuItem_Click(object sender, RoutedEventArgs e) => SelectAllChecked(false);

    private void SelectAllChecked(bool value)
    {
        foreach (var item in _viewModel.ProgramsView.Cast<ProgramViewModel>())
        {
            item.IsChecked = value;
        }
    }

    private void InvertSelectionMenuItem_Click(object sender, RoutedEventArgs e)
    {
        foreach (var item in _viewModel.ProgramsView.Cast<ProgramViewModel>())
        {
            item.IsChecked = !item.IsChecked;
        }
    }

    private void FindMenuItem_Click(object sender, RoutedEventArgs e) => FocusSearchBox();

    private void FocusSearchBox()
    {
        SearchBox.Focus();
        Keyboard.Focus(SearchBox);
        SearchBox.SelectAll();
    }

    // ===================== View menu (theme) =====================

    private void ThemeLight_Click(object sender, RoutedEventArgs e) => ApplyTheme(AppTheme.Light);

    private void ThemeDark_Click(object sender, RoutedEventArgs e) => ApplyTheme(AppTheme.Dark);

    private void ThemeSystem_Click(object sender, RoutedEventArgs e) => ApplyTheme(AppTheme.System);

    private void ApplyTheme(AppTheme theme)
    {
        _viewModel.Settings.Theme = theme;
        _viewModel.SaveSettings();
        App.ApplyTheme(theme);
    }

    // ===================== Actions menu =====================

    private async void UninstallMenuItem_Click(object sender, RoutedEventArgs e) => await UninstallSelectedAsync(silent: false);

    private async void UninstallSilentMenuItem_Click(object sender, RoutedEventArgs e) => await UninstallSelectedAsync(silent: true);

    private async void ProgramsGrid_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e) =>
        await UninstallSelectedAsync(silent: false);

    private async Task UninstallSelectedAsync(bool silent)
    {
        var target = _viewModel.SelectedProgram;
        if (target is null)
        {
            System.Windows.MessageBox.Show(this, "Select a program first.", "Uninstaller",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        // "Uninstall Silently" always means silent; otherwise honor the
        // user's "prefer silent uninstall" option from Settings.
        silent = silent || _viewModel.Settings.PreferSilentUninstall;

        if (_viewModel.ConfirmBeforeUninstall)
        {
            var confirm = System.Windows.MessageBox.Show(this,
                $"Uninstall '{target.DisplayName}'?",
                "Confirm Uninstall", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (confirm != MessageBoxResult.Yes)
            {
                return;
            }
        }

        if (!await WarnIfRunningAsync(target))
        {
            return;
        }

        var progressWindow = new ProgressLogWindow($"Uninstalling {target.DisplayName}...", allowCancel: false) { Owner = this };
        progressWindow.Show();

        var progress = new Progress<string>(progressWindow.AppendLine);
        var result = await _viewModel.UninstallAsync(target, silent, progress);

        progressWindow.SetComplete();
        _viewModel.StatusText = result.Message ?? (result.Success ? "Done." : "Failed.");

        if (!result.Success)
        {
            progressWindow.AppendLine("You can try 'Force Remove / Clean Leftovers...' if the uninstaller is broken or missing.");
        }
    }

    /// <summary>
    /// Checks whether the given program's install folder has a running
    /// process open, and if so, offers to close it before continuing.
    /// Returns false if the caller should abort (user chose Cancel).
    /// </summary>
    private async Task<bool> WarnIfRunningAsync(ProgramViewModel target)
    {
        if (string.IsNullOrWhiteSpace(target.Program.InstallLocation))
        {
            return true;
        }

        var runningProcesses = await Task.Run(() =>
            RunningProcessChecker.FindProcessesUnder(target.Program.InstallLocation));

        if (runningProcesses.Count == 0)
        {
            return true;
        }

        var names = string.Join(", ", runningProcesses.Select(p => $"{p.ProcessName} (PID {p.ProcessId})"));
        var choice = System.Windows.MessageBox.Show(this,
            $"'{target.DisplayName}' appears to still be running:\n\n{names}\n\n" +
            "Files it has open can't be removed while it's running. Close it now and continue?",
            "Program Is Running", MessageBoxButton.YesNoCancel, MessageBoxImage.Warning);

        if (choice == MessageBoxResult.Cancel)
        {
            return false;
        }

        if (choice == MessageBoxResult.Yes)
        {
            var stillRunning = await Task.Run(() =>
                RunningProcessChecker.TryStop(runningProcesses, TimeSpan.FromSeconds(5)));
            if (stillRunning.Count > 0)
            {
                System.Windows.MessageBox.Show(this,
                    $"Could not close: {string.Join(", ", stillRunning.Select(p => p.ProcessName))}. " +
                    "Some files may fail to remove.",
                    "Uninstaller", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // MessageBoxResult.No falls through and continues anyway.
        return true;
    }

    private async void UninstallCheckedMenuItem_Click(object sender, RoutedEventArgs e)
    {
        var targets = _viewModel.CheckedPrograms.ToList();
        if (targets.Count == 0)
        {
            System.Windows.MessageBox.Show(this, "Check one or more programs first using the checkboxes in the list.",
                "Uninstaller", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var confirm = System.Windows.MessageBox.Show(this,
            $"Uninstall {targets.Count} selected program(s)? Each will run its own uninstaller in turn.",
            "Confirm Batch Uninstall", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (confirm != MessageBoxResult.Yes)
        {
            return;
        }

        var progressWindow = new ProgressLogWindow($"Uninstalling {targets.Count} program(s)...") { Owner = this };
        progressWindow.Show();
        var progress = new Progress<string>(progressWindow.AppendLine);

        var completed = 0;
        foreach (var target in targets)
        {
            if (progressWindow.CancellationToken.IsCancellationRequested)
            {
                progressWindow.AppendLine("Batch uninstall cancelled.");
                break;
            }

            progressWindow.AppendLine($"--- {target.DisplayName} ---");
            await _viewModel.UninstallAsync(target, silent: true, progress);
            completed++;
            progressWindow.SetProgress((double)completed / targets.Count);
        }

        progressWindow.SetComplete();
        _viewModel.StatusText = $"Batch uninstall finished: {completed} of {targets.Count} processed.";
    }

    private void ModifyMenuItem_Click(object sender, RoutedEventArgs e)
    {
        var target = _viewModel.SelectedProgram;
        if (target is null || string.IsNullOrWhiteSpace(target.Program.ModifyPath))
        {
            System.Windows.MessageBox.Show(this, "This program does not support Modify.", "Uninstaller",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var (fileName, arguments) = CommandLineParser.Split(target.Program.ModifyPath);
        var psi = new ProcessStartInfo
        {
            FileName = fileName,
            Arguments = arguments,
            UseShellExecute = true
        };

        // Same reasoning as UninstallService: only elevate for entries that
        // came from HKEY_LOCAL_MACHINE. A per-user (HKCU) entry - writable
        // by any unprivileged process - must never be auto-elevated just
        // because the user clicked Modify on it.
        if (target.Program.Scope == ProgramScope.Machine)
        {
            psi.Verb = "runas";
        }

        try
        {
            Process.Start(psi);
        }
        catch (Win32Exception ex) when (ex.NativeErrorCode == 1223)
        {
            _viewModel.StatusText = $"Elevation was declined for '{target.DisplayName}'.";
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(this, $"Failed to launch Modify: {ex.Message}", "Uninstaller",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void ForceRemoveMenuItem_Click(object sender, RoutedEventArgs e)
    {
        var target = _viewModel.SelectedProgram;
        if (target is null)
        {
            System.Windows.MessageBox.Show(this, "Select a program first.", "Uninstaller",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        if (!await WarnIfRunningAsync(target))
        {
            return;
        }

        var candidates = _viewModel.ForceRemovalService.FindResidue(target.Program);
        var dialog = new ForceRemoveWindow(target.DisplayName, candidates) { Owner = this };
        var dialogResult = dialog.ShowDialog();

        if (dialogResult == true && dialog.Confirmed)
        {
            var results = _viewModel.ForceRemovalService.Remove(dialog.SelectedItems);
            var failures = results.Where(r => !r.Success).ToList();

            _viewModel.StatusText = failures.Count == 0
                ? $"Force-removed {results.Count} item(s) for '{target.DisplayName}'."
                : $"Force-removed {results.Count - failures.Count} of {results.Count} item(s); {failures.Count} failed.";

            await _viewModel.RefreshAsync();
        }
    }

    private void OpenInstallLocationMenuItem_Click(object sender, RoutedEventArgs e)
    {
        var target = _viewModel.SelectedProgram;
        if (target is null || string.IsNullOrWhiteSpace(target.Program.InstallLocation) ||
            !Directory.Exists(target.Program.InstallLocation))
        {
            System.Windows.MessageBox.Show(this, "No install location is available for this program.", "Uninstaller",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        MainViewModel.OpenFolder(target.Program.InstallLocation);
    }

    private void OpenRegistryKeyMenuItem_Click(object sender, RoutedEventArgs e)
    {
        var target = _viewModel.SelectedProgram;
        if (target is null || string.IsNullOrWhiteSpace(target.Program.RegistryKeyPath))
        {
            System.Windows.MessageBox.Show(this, "This program has no registry entry to open.", "Uninstaller",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        MainViewModel.OpenRegistryEditorAt(target.Program.RegistryKeyPath);
    }

    private void CopyUninstallCommandMenuItem_Click(object sender, RoutedEventArgs e)
    {
        var target = _viewModel.SelectedProgram;
        var command = target?.Program.UninstallString;
        if (string.IsNullOrWhiteSpace(command))
        {
            System.Windows.MessageBox.Show(this, "This program has no uninstall command to copy.", "Uninstaller",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        System.Windows.Clipboard.SetText(command);
        _viewModel.StatusText = "Uninstall command copied to clipboard.";
    }

    // ===================== Tools menu =====================

    private async void LeftoverScannerMenuItem_Click(object sender, RoutedEventArgs e)
    {
        var window = new LeftoverScannerWindow(
            _viewModel.LeftoverScannerService,
            _viewModel.ForceRemovalService,
            _viewModel.AllInstalledPrograms)
        { Owner = this };

        window.ShowDialog();
        await _viewModel.RefreshAsync();
    }

    private void OptionsMenuItem_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new SettingsWindow(_viewModel.Settings) { Owner = this };
        if (dialog.ShowDialog() == true)
        {
            _viewModel.ShowSystemComponents = dialog.Result.ShowSystemComponents;
            _viewModel.ShowWindowsUpdates = dialog.Result.ShowWindowsUpdates;
            _viewModel.ConfirmBeforeUninstall = dialog.Result.ConfirmBeforeUninstall;
            _viewModel.Settings.PreferSilentUninstall = dialog.Result.PreferSilentUninstall;
            _viewModel.Settings.Theme = dialog.Result.Theme;
            _viewModel.SaveSettings();
            App.ApplyTheme(dialog.Result.Theme);
        }
    }

    // ===================== Help menu =====================

    private void CheckForUpdatesMenuItem_Click(object sender, RoutedEventArgs e)
    {
        var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version?.ToString(3) ?? "1.0.0";
        System.Windows.MessageBox.Show(this, $"You're running Uninstaller {version}, the latest version.",
            "Check for Updates", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void AboutMenuItem_Click(object sender, RoutedEventArgs e) =>
        new AboutWindow { Owner = this }.ShowDialog();
}
