namespace Uninstaller.Core.Models;

/// <summary>
/// Represents a single installed application, sourced either from the classic
/// Win32 "Uninstall" registry keys or from the AppX/MSIX package catalog.
/// </summary>
public sealed class InstalledProgram
{
    public required string DisplayName { get; init; }
    public string? DisplayVersion { get; init; }
    public string? Publisher { get; init; }
    public DateTime? InstallDate { get; init; }
    public long? EstimatedSizeKb { get; init; }
    public string? InstallLocation { get; init; }
    public string? UninstallString { get; init; }
    public string? QuietUninstallString { get; init; }
    public string? ModifyPath { get; init; }
    public string? DisplayIconPath { get; init; }
    public string? Comments { get; init; }
    public string? HelpLink { get; init; }
    public string? UrlInfoAbout { get; init; }

    public bool SystemComponent { get; init; }
    public bool IsWindowsUpdate { get; init; }
    public bool IsMsi { get; init; }

    public ProgramSource Source { get; init; } = ProgramSource.Win32Registry;
    public ProgramScope Scope { get; init; } = ProgramScope.Machine;
    public ProgramArchitecture Architecture { get; init; } = ProgramArchitecture.Unknown;

    /// <summary>
    /// Full registry path (e.g. HKEY_LOCAL_MACHINE\Software\...\Uninstall\{GUID})
    /// for Win32 entries, used for direct removal of the uninstall key itself.
    /// Null for AppX entries.
    /// </summary>
    public string? RegistryKeyPath { get; init; }

    /// <summary>
    /// PackageFullName for AppX/MSIX entries. Null for Win32 entries.
    /// </summary>
    public string? PackageFullName { get; init; }

    /// <summary>
    /// Stable identity used for de-duplication and for matching leftovers.
    /// </summary>
    public string Key => PackageFullName ?? RegistryKeyPath ?? $"{DisplayName}|{DisplayVersion}|{Publisher}";

    public string SizeDisplay => EstimatedSizeKb is > 0
        ? FormatSize(EstimatedSizeKb.Value * 1024)
        : "Unknown";

    public string ArchitectureDisplay => Architecture switch
    {
        ProgramArchitecture.Bit64 => "64-bit",
        ProgramArchitecture.Bit32 => "32-bit",
        _ => Source == ProgramSource.AppxPackage ? "Store app" : "Unknown"
    };

    public static string FormatSize(long bytes)
    {
        string[] units = { "B", "KB", "MB", "GB", "TB" };
        double size = bytes;
        int unit = 0;
        while (size >= 1024 && unit < units.Length - 1)
        {
            size /= 1024;
            unit++;
        }
        return $"{size:0.##} {units[unit]}";
    }
}
