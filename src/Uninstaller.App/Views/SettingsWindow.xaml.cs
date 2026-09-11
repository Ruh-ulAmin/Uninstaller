using System.Windows;
using Uninstaller.Core.Models;

namespace Uninstaller.App.Views;

public partial class SettingsWindow : Window
{
    public AppSettings Result { get; }

    public SettingsWindow(AppSettings current)
    {
        InitializeComponent();

        // Work on a copy so cancelling never mutates the caller's settings.
        Result = new AppSettings
        {
            ShowSystemComponents = current.ShowSystemComponents,
            ShowWindowsUpdates = current.ShowWindowsUpdates,
            ConfirmBeforeUninstall = current.ConfirmBeforeUninstall,
            PreferSilentUninstall = current.PreferSilentUninstall,
            Theme = current.Theme
        };

        ShowSystemComponentsCheck.IsChecked = Result.ShowSystemComponents;
        ShowWindowsUpdatesCheck.IsChecked = Result.ShowWindowsUpdates;
        ConfirmBeforeUninstallCheck.IsChecked = Result.ConfirmBeforeUninstall;
        PreferSilentUninstallCheck.IsChecked = Result.PreferSilentUninstall;

        switch (Result.Theme)
        {
            case AppTheme.Light:
                ThemeLightRadio.IsChecked = true;
                break;
            case AppTheme.Dark:
                ThemeDarkRadio.IsChecked = true;
                break;
            default:
                ThemeSystemRadio.IsChecked = true;
                break;
        }
    }

    private void SaveButton_Click(object sender, RoutedEventArgs e)
    {
        Result.ShowSystemComponents = ShowSystemComponentsCheck.IsChecked ?? false;
        Result.ShowWindowsUpdates = ShowWindowsUpdatesCheck.IsChecked ?? false;
        Result.ConfirmBeforeUninstall = ConfirmBeforeUninstallCheck.IsChecked ?? true;
        Result.PreferSilentUninstall = PreferSilentUninstallCheck.IsChecked ?? false;
        Result.Theme = ThemeDarkRadio.IsChecked == true ? AppTheme.Dark
            : ThemeLightRadio.IsChecked == true ? AppTheme.Light
            : AppTheme.System;

        DialogResult = true;
        Close();
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
