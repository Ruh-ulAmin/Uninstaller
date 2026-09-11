using System.Collections.ObjectModel;
using System.Windows;
using Uninstaller.App.Helpers;
using Uninstaller.Core.Models;

namespace Uninstaller.App.Views;

public sealed class SelectableLeftoverItem : ViewModelBase
{
    private bool _isSelected;

    public SelectableLeftoverItem(LeftoverItem item, bool preSelected)
    {
        Item = item;
        _isSelected = preSelected;
    }

    public LeftoverItem Item { get; }

    public bool IsSelected
    {
        get => _isSelected;
        set => SetField(ref _isSelected, value);
    }

    public string KindDisplay => Item.Kind switch
    {
        LeftoverKind.Folder => "Folder",
        LeftoverKind.RegistryKey => "Registry",
        LeftoverKind.Shortcut => "Shortcut",
        LeftoverKind.OrphanUninstallEntry => "Orphan entry",
        _ => Item.Kind.ToString()
    };
}

/// <summary>
/// Lets the user review and confirm exactly which residual items to delete
/// for one specific program before <see cref="Core.Services.ForceRemovalService.Remove"/>
/// is invoked. Nothing is deleted unless the user explicitly presses "Remove Selected".
/// </summary>
public partial class ForceRemoveWindow : Window
{
    private readonly ObservableCollection<SelectableLeftoverItem> _items;

    public IReadOnlyList<LeftoverItem> SelectedItems { get; private set; } = Array.Empty<LeftoverItem>();
    public bool Confirmed { get; private set; }

    public ForceRemoveWindow(string programName, IReadOnlyList<LeftoverItem> candidates)
    {
        InitializeComponent();
        HeaderText.Text = candidates.Count == 0
            ? $"No leftover files or registry entries were found for '{programName}'."
            : $"Force remove leftovers for '{programName}'";

        _items = new ObservableCollection<SelectableLeftoverItem>(
            candidates.Select(c => new SelectableLeftoverItem(c, preSelected: true)));
        ItemsList.ItemsSource = _items;
    }

    private void RemoveButton_Click(object sender, RoutedEventArgs e)
    {
        SelectedItems = _items.Where(i => i.IsSelected).Select(i => i.Item).ToList();

        if (SelectedItems.Count == 0)
        {
            System.Windows.MessageBox.Show(this, "No items are selected.", "Force Remove",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var confirm = System.Windows.MessageBox.Show(this,
            $"Permanently delete {SelectedItems.Count} selected item(s)? This cannot be undone.",
            "Confirm Force Remove", MessageBoxButton.YesNo, MessageBoxImage.Warning);

        if (confirm != MessageBoxResult.Yes)
        {
            return;
        }

        Confirmed = true;
        DialogResult = true;
        Close();
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
