using System.Collections.ObjectModel;
using System.Windows;

namespace Uninstaller.App.Views;

/// <summary>
/// A simple modal progress + live log window used while uninstall / cleanup
/// operations run, so the user can see exactly what is happening and cancel
/// a queued batch before it moves on to the next item.
/// </summary>
public partial class ProgressLogWindow : Window
{
    private readonly ObservableCollection<string> _lines = new();
    private readonly CancellationTokenSource _cts = new();

    public CancellationToken CancellationToken => _cts.Token;

    public ProgressLogWindow(string title, bool allowCancel = true)
    {
        InitializeComponent();
        TitleText.Text = title;
        LogItems.ItemsSource = _lines;

        if (!allowCancel)
        {
            // A single uninstall runs one external process to completion -
            // killing it mid-way could leave the target program half-removed,
            // so only batch runs (which cancel *between* items) offer Cancel.
            CancelButton.Visibility = Visibility.Collapsed;
        }
    }

    public void AppendLine(string line)
    {
        Dispatcher.Invoke(() =>
        {
            _lines.Add(line);
            LogScrollViewer.ScrollToEnd();
        });
    }

    public void SetProgress(double fraction)
    {
        Dispatcher.Invoke(() =>
        {
            ProgressIndicator.IsIndeterminate = false;
            ProgressIndicator.Value = Math.Clamp(fraction, 0, 1) * 100;
        });
    }

    public void SetComplete()
    {
        Dispatcher.Invoke(() =>
        {
            ProgressIndicator.IsIndeterminate = false;
            ProgressIndicator.Value = 100;
            CancelButton.IsEnabled = false;
            CloseButton.IsEnabled = true;
        });
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        _cts.Cancel();
        CancelButton.IsEnabled = false;
        AppendLine("Cancellation requested - finishing the current item, then stopping...");
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
    {
        if (!CloseButton.IsEnabled)
        {
            // Operation still running - treat the window close (X button) like Cancel.
            _cts.Cancel();
        }
        base.OnClosing(e);
    }
}
