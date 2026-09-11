namespace Uninstaller.Core.Models;

public enum AppTheme
{
    System,
    Light,
    Dark
}

/// <summary>
/// Persisted user preferences for the application.
/// </summary>
public sealed class AppSettings
{
    public bool ShowSystemComponents { get; set; }
    public bool ShowWindowsUpdates { get; set; }
    public bool ConfirmBeforeUninstall { get; set; } = true;
    public bool PreferSilentUninstall { get; set; }
    public AppTheme Theme { get; set; } = AppTheme.System;
    public string SortColumn { get; set; } = "DisplayName";
    public bool SortAscending { get; set; } = true;
}
