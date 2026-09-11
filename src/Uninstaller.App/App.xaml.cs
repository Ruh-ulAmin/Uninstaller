using System.Windows;
using System.Windows.Threading;
using Uninstaller.Core.Models;
using Uninstaller.Core.Services;

namespace Uninstaller.App;

public partial class App : System.Windows.Application
{
    private const int LightThemeDictionaryIndex = 0;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        DispatcherUnhandledException += OnDispatcherUnhandledException;
        AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;

        var settingsService = new SettingsService();
        var settings = settingsService.Load();
        ApplyTheme(settings.Theme);
    }

    /// <summary>
    /// Swaps the first merged resource dictionary (the color palette) for the
    /// requested theme. "System" follows the current Windows app theme.
    /// </summary>
    public static void ApplyTheme(AppTheme theme)
    {
        var effective = theme == AppTheme.System ? DetectSystemTheme() : theme;
        var uri = effective == AppTheme.Dark
            ? new Uri("Themes/DarkTheme.xaml", UriKind.Relative)
            : new Uri("Themes/LightTheme.xaml", UriKind.Relative);

        var app = Current;
        var newDictionary = new ResourceDictionary { Source = uri };
        app.Resources.MergedDictionaries[LightThemeDictionaryIndex] = newDictionary;
    }

    private static AppTheme DetectSystemTheme()
    {
        try
        {
            using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
            var value = key?.GetValue("AppsUseLightTheme");
            if (value is int i)
            {
                return i == 0 ? AppTheme.Dark : AppTheme.Light;
            }
        }
        catch
        {
            // Fall through to the default below.
        }
        return AppTheme.Light;
    }

    private void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        LogAndShow(e.Exception);
        e.Handled = true;
    }

    private void OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        if (e.ExceptionObject is Exception ex)
        {
            LogAndShow(ex);
        }
    }

    private static void LogAndShow(Exception ex)
    {
        try
        {
            new FileLogger().Log($"Unhandled exception: {ex}");
        }
        catch
        {
            // Ignore logging failures here - we still want the message box below.
        }

        System.Windows.MessageBox.Show(
            $"An unexpected error occurred:\n\n{ex.Message}",
            "Uninstaller",
            MessageBoxButton.OK,
            MessageBoxImage.Error);
    }
}
