using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows.Data;
using Uninstaller.App.Helpers;
using Uninstaller.Core.Models;
using Uninstaller.Core.Services;

namespace Uninstaller.App.ViewModels;

public sealed class MainViewModel : ViewModelBase
{
    private readonly RegistryProgramScanner _registryScanner = new();
    private readonly AppxProgramScanner _appxScanner = new();
    private readonly UninstallService _uninstallService;
    private readonly ForceRemovalService _forceRemovalService;
    private readonly LeftoverScannerService _leftoverScannerService = new();
    private readonly SettingsService _settingsService = new();
    private readonly FileLogger _logger = new();

    private readonly ObservableCollection<ProgramViewModel> _allPrograms = new();
    private string _searchText = string.Empty;
    private ProgramViewModel? _selectedProgram;
    private bool _isBusy;
    private string _statusText = "Ready.";
    private bool _showSystemComponents;
    private bool _showWindowsUpdates;
    private bool _confirmBeforeUninstall = true;

    public MainViewModel()
    {
        _uninstallService = new UninstallService(_logger);
        _forceRemovalService = new ForceRemovalService(_logger);

        Settings = _settingsService.Load();
        _showSystemComponents = Settings.ShowSystemComponents;
        _showWindowsUpdates = Settings.ShowWindowsUpdates;
        _confirmBeforeUninstall = Settings.ConfirmBeforeUninstall;

        ProgramsView = CollectionViewSource.GetDefaultView(_allPrograms);
        ProgramsView.Filter = FilterProgram;
        ProgramsView.SortDescriptions.Add(new SortDescription(nameof(ProgramViewModel.DisplayName), ListSortDirection.Ascending));
        ((INotifyCollectionChanged)ProgramsView).CollectionChanged += (_, _) => OnPropertyChanged(nameof(VisibleProgramCount));

        IsAdministrator = AdminHelper.IsRunningAsAdministrator();
    }

    public AppSettings Settings { get; }

    public ICollectionView ProgramsView { get; }

    public bool IsAdministrator { get; }

    public string SearchText
    {
        get => _searchText;
        set
        {
            if (SetField(ref _searchText, value))
            {
                ProgramsView.Refresh();
            }
        }
    }

    public bool ShowSystemComponents
    {
        get => _showSystemComponents;
        set
        {
            if (SetField(ref _showSystemComponents, value))
            {
                Settings.ShowSystemComponents = value;
                _settingsService.Save(Settings);
                ProgramsView.Refresh();
            }
        }
    }

    public bool ShowWindowsUpdates
    {
        get => _showWindowsUpdates;
        set
        {
            if (SetField(ref _showWindowsUpdates, value))
            {
                Settings.ShowWindowsUpdates = value;
                _settingsService.Save(Settings);
                ProgramsView.Refresh();
            }
        }
    }

    public bool ConfirmBeforeUninstall
    {
        get => _confirmBeforeUninstall;
        set
        {
            if (SetField(ref _confirmBeforeUninstall, value))
            {
                Settings.ConfirmBeforeUninstall = value;
                _settingsService.Save(Settings);
            }
        }
    }

    public ProgramViewModel? SelectedProgram
    {
        get => _selectedProgram;
        set => SetField(ref _selectedProgram, value);
    }

    public bool IsBusy
    {
        get => _isBusy;
        set => SetField(ref _isBusy, value);
    }

    public string StatusText
    {
        get => _statusText;
        set => SetField(ref _statusText, value);
    }

    public int TotalProgramCount => _allPrograms.Count;

    public int VisibleProgramCount => ProgramsView.Cast<object>().Count();

    public IEnumerable<ProgramViewModel> CheckedPrograms => _allPrograms.Where(p => p.IsChecked);

    public FileLogger Logger => _logger;
    public ForceRemovalService ForceRemovalService => _forceRemovalService;
    public LeftoverScannerService LeftoverScannerService => _leftoverScannerService;
    public IReadOnlyList<InstalledProgram> AllInstalledPrograms => _allPrograms.Select(p => p.Program).ToList();

    public void SaveSettings() => _settingsService.Save(Settings);

    public async Task RefreshAsync()
    {
        IsBusy = true;
        StatusText = "Scanning installed programs...";
        try
        {
            var registryPrograms = await Task.Run(() => _registryScanner.Scan()).ConfigureAwait(true);
            var appxPrograms = await _appxScanner.ScanAsync().ConfigureAwait(true);

            _allPrograms.Clear();
            foreach (var program in registryPrograms.Concat(appxPrograms).OrderBy(p => p.DisplayName, StringComparer.OrdinalIgnoreCase))
            {
                _allPrograms.Add(new ProgramViewModel(program));
            }

            OnPropertyChanged(nameof(TotalProgramCount));
            OnPropertyChanged(nameof(VisibleProgramCount));
            StatusText = $"Found {_allPrograms.Count} installed programs.";
            _logger.Log($"Scan complete: {_allPrograms.Count} programs found.");
        }
        catch (Exception ex)
        {
            StatusText = $"Scan failed: {ex.Message}";
            _logger.Log($"Scan failed: {ex}");
        }
        finally
        {
            IsBusy = false;
        }
    }

    public async Task<OperationResult> UninstallAsync(ProgramViewModel target, bool silent, IProgress<string>? progress)
    {
        var result = await _uninstallService.UninstallAsync(target.Program, silent, progress).ConfigureAwait(true);
        if (result.Success)
        {
            _allPrograms.Remove(target);
            OnPropertyChanged(nameof(TotalProgramCount));
            OnPropertyChanged(nameof(VisibleProgramCount));
        }
        return result;
    }

    public void RemoveFromList(ProgramViewModel target)
    {
        _allPrograms.Remove(target);
        OnPropertyChanged(nameof(TotalProgramCount));
        OnPropertyChanged(nameof(VisibleProgramCount));
    }

    public void ExportListToCsv(string filePath)
    {
        using var writer = new StreamWriter(filePath, append: false, System.Text.Encoding.UTF8);
        writer.WriteLine("Name,Publisher,Version,Size,InstallDate,Architecture,Source");
        foreach (var vm in _allPrograms)
        {
            var p = vm.Program;
            writer.WriteLine(string.Join(',', new[]
            {
                CsvEscape(p.DisplayName),
                CsvEscape(p.Publisher ?? string.Empty),
                CsvEscape(p.DisplayVersion ?? string.Empty),
                CsvEscape(p.SizeDisplay),
                CsvEscape(p.InstallDate?.ToString("yyyy-MM-dd") ?? string.Empty),
                CsvEscape(p.ArchitectureDisplay),
                CsvEscape(p.Source.ToString())
            }));
        }
        _logger.Log($"Exported program list to {filePath}");
    }

    private static string CsvEscape(string value)
    {
        // A DisplayName is sourced from the registry, which an unprivileged
        // process can write to under HKCU. Prefix any value that would be
        // interpreted as a formula (=, +, -, @, or a tab/CR that Excel also
        // treats as a formula lead-in) with an apostrophe so opening the
        // exported CSV in Excel/Sheets can never execute a formula -
        // the standard CSV-injection mitigation.
        if (value.Length > 0 && (value[0] is '=' or '+' or '-' or '@' or '\t' or '\r'))
        {
            value = "'" + value;
        }

        if (value.Contains(',') || value.Contains('"') || value.Contains('\n'))
        {
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }
        return value;
    }

    public static void OpenFolder(string path)
    {
        if (Directory.Exists(path))
        {
            Process.Start(new ProcessStartInfo("explorer.exe", $"\"{path}\"") { UseShellExecute = true });
        }
    }

    public static void OpenRegistryEditorAt(string registryPath)
    {
        // regedit reads the last-visited key from this value; jumping straight
        // to the uninstall entry saves the user from navigating manually.
        try
        {
            using var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Applets\Regedit");
            key?.SetValue("LastKey", "Computer\\" + registryPath);
        }
        catch
        {
            // Non-fatal - regedit will just open at its previous location.
        }

        Process.Start(new ProcessStartInfo("regedit.exe") { UseShellExecute = true });
    }

    private bool FilterProgram(object obj)
    {
        if (obj is not ProgramViewModel vm)
        {
            return false;
        }

        var program = vm.Program;

        if (!ShowSystemComponents && program.SystemComponent)
        {
            return false;
        }

        if (!ShowWindowsUpdates && program.IsWindowsUpdate)
        {
            return false;
        }

        if (string.IsNullOrWhiteSpace(SearchText))
        {
            return true;
        }

        return program.DisplayName.Contains(SearchText, StringComparison.OrdinalIgnoreCase)
            || (program.Publisher?.Contains(SearchText, StringComparison.OrdinalIgnoreCase) ?? false);
    }
}
