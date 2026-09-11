using System.Collections.ObjectModel;
using System.Windows;
using Uninstaller.Core.Models;
using Uninstaller.Core.Services;

namespace Uninstaller.App.Views;

/// <summary>
/// System-wide leftover scan across all installed programs at once, as
/// opposed to <see cref="ForceRemoveWindow"/> which targets a single
/// program. Results are heuristic (name/location matching) so nothing is
/// pre-selected here, unlike the single-program force-remove flow.
/// </summary>
public partial class LeftoverScannerWindow : Window
{
    private readonly LeftoverScannerService _scannerService;
    private readonly ForceRemovalService _forceRemovalService;
    private readonly IReadOnlyList<InstalledProgram> _installedPrograms;
    private ObservableCollection<SelectableLeftoverItem> _items = new();

    public LeftoverScannerWindow(
        LeftoverScannerService scannerService,
        ForceRemovalService forceRemovalService,
        IReadOnlyList<InstalledProgram> installedPrograms)
    {
        InitializeComponent();
        _scannerService = scannerService;
        _forceRemovalService = forceRemovalService;
        _installedPrograms = installedPrograms;

        Loaded += async (_, _) => await RunScanAsync();
    }

    private async Task RunScanAsync()
    {
        ScanningOverlay.Visibility = Visibility.Visible;
        try
        {
            var results = await Task.Run(() =>
            {
                var orphanEntries = _scannerService.FindOrphanUninstallEntries(_installedPrograms);
                var orphanFolders = _scannerService.FindOrphanFolders(_installedPrograms);
                return orphanEntries.Concat(orphanFolders).ToList();
            });

            _items = new ObservableCollection<SelectableLeftoverItem>(
                results.Select(r => new SelectableLeftoverItem(r, preSelected: false)));
            ItemsList.ItemsSource = _items;
        }
        finally
        {
            ScanningOverlay.Visibility = Visibility.Collapsed;
        }
    }

    private async void RescanButton_Click(object sender, RoutedEventArgs e) => await RunScanAsync();

    private void RemoveButton_Click(object sender, RoutedEventArgs e)
    {
        var selected = _items.Where(i => i.IsSelected).Select(i => i.Item).ToList();
        if (selected.Count == 0)
        {
            System.Windows.MessageBox.Show(this, "No items are selected.", "Leftover Scanner",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var confirm = System.Windows.MessageBox.Show(this,
            $"Permanently delete {selected.Count} selected item(s)? This cannot be undone.",
            "Confirm Removal", MessageBoxButton.YesNo, MessageBoxImage.Warning);

        if (confirm != MessageBoxResult.Yes)
        {
            return;
        }

        var results = _forceRemovalService.Remove(selected);
        var failures = results.Where(r => !r.Success).ToList();

        foreach (var removed in selected)
        {
            var match = _items.FirstOrDefault(i => i.Item == removed);
            if (match is not null)
            {
                _items.Remove(match);
            }
        }

        var message = failures.Count == 0
            ? $"Removed {selected.Count} item(s) successfully."
            : $"Removed {selected.Count - failures.Count} item(s). {failures.Count} failed:\n" +
              string.Join('\n', failures.Select(f => f.Message));

        System.Windows.MessageBox.Show(this, message, "Leftover Scanner", MessageBoxButton.OK,
            failures.Count == 0 ? MessageBoxImage.Information : MessageBoxImage.Warning);
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e) => Close();
}
