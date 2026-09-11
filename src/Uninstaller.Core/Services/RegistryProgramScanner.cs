using System.Text.RegularExpressions;
using Microsoft.Win32;
using Uninstaller.Core.Models;

namespace Uninstaller.Core.Services;

/// <summary>
/// Enumerates installed Win32 applications by reading the standard
/// "Uninstall" registry keys across both registry views (32/64-bit) and
/// both HKLM (machine-wide) and HKCU (per-user) hives, which together is
/// the same data source Windows' own "Programs and Features" applet uses.
/// </summary>
public sealed class RegistryProgramScanner
{
    private const string UninstallSubKey = @"Software\Microsoft\Windows\CurrentVersion\Uninstall";
    private static readonly Regex KbPattern = new(@"^KB\d{6,}", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public IReadOnlyList<InstalledProgram> Scan()
    {
        var results = new List<InstalledProgram>();
        var seenKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        ScanRoot(RegistryHive.LocalMachine, RegistryView.Registry64, ProgramScope.Machine, ProgramArchitecture.Bit64, results, seenKeys);
        ScanRoot(RegistryHive.LocalMachine, RegistryView.Registry32, ProgramScope.Machine, ProgramArchitecture.Bit32, results, seenKeys);
        ScanRoot(RegistryHive.CurrentUser, RegistryView.Registry64, ProgramScope.User, ProgramArchitecture.Bit64, results, seenKeys);

        return results;
    }

    private static void ScanRoot(
        RegistryHive hive,
        RegistryView view,
        ProgramScope scope,
        ProgramArchitecture architecture,
        List<InstalledProgram> results,
        HashSet<string> seenKeys)
    {
        RegistryKey? baseKey = null;
        RegistryKey? uninstallKey = null;
        try
        {
            baseKey = RegistryKey.OpenBaseKey(hive, view);
            uninstallKey = baseKey.OpenSubKey(UninstallSubKey);
            if (uninstallKey is null)
            {
                return;
            }

            foreach (var subKeyName in uninstallKey.GetSubKeyNames())
            {
                using var subKey = uninstallKey.OpenSubKey(subKeyName);
                if (subKey is null)
                {
                    continue;
                }

                var fullPath = subKey.Name; // e.g. HKEY_LOCAL_MACHINE\Software\...\Uninstall\{GUID}
                if (!seenKeys.Add($"{fullPath}|{view}"))
                {
                    continue;
                }

                var program = TryReadProgram(subKey, fullPath, scope, architecture);
                if (program is not null)
                {
                    results.Add(program);
                }
            }
        }
        catch (System.Security.SecurityException)
        {
            // No permission to read this hive/view - skip silently.
        }
        catch (UnauthorizedAccessException)
        {
            // Same as above.
        }
        finally
        {
            uninstallKey?.Dispose();
            baseKey?.Dispose();
        }
    }

    private static InstalledProgram? TryReadProgram(RegistryKey key, string fullPath, ProgramScope scope, ProgramArchitecture architecture)
    {
        var displayName = key.GetValue("DisplayName") as string;
        if (string.IsNullOrWhiteSpace(displayName))
        {
            // Entries without a DisplayName are not meant to be shown to end users
            // (e.g. shared components, hotfix containers).
            return null;
        }

        var uninstallString = key.GetValue("UninstallString") as string;
        var quietUninstallString = key.GetValue("QuietUninstallString") as string;
        var releaseType = key.GetValue("ReleaseType") as string;
        var parentDisplayName = key.GetValue("ParentDisplayName") as string;

        var systemComponent = ToBool(key.GetValue("SystemComponent"));
        var isWindowsUpdate =
            KbPattern.IsMatch(displayName) ||
            string.Equals(releaseType, "Security Update", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(releaseType, "Update", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(releaseType, "Hotfix", StringComparison.OrdinalIgnoreCase) ||
            !string.IsNullOrEmpty(parentDisplayName);

        return new InstalledProgram
        {
            DisplayName = displayName.Trim(),
            DisplayVersion = key.GetValue("DisplayVersion") as string,
            Publisher = key.GetValue("Publisher") as string,
            InstallDate = ParseInstallDate(key.GetValue("InstallDate") as string),
            EstimatedSizeKb = ToLong(key.GetValue("EstimatedSize")),
            InstallLocation = NormalizePath(key.GetValue("InstallLocation") as string),
            UninstallString = uninstallString,
            QuietUninstallString = quietUninstallString,
            ModifyPath = key.GetValue("ModifyPath") as string,
            DisplayIconPath = key.GetValue("DisplayIcon") as string,
            Comments = key.GetValue("Comments") as string,
            HelpLink = key.GetValue("HelpLink") as string,
            UrlInfoAbout = key.GetValue("URLInfoAbout") as string,
            SystemComponent = systemComponent,
            IsWindowsUpdate = isWindowsUpdate,
            IsMsi = ToBool(key.GetValue("WindowsInstaller")),
            Source = ProgramSource.Win32Registry,
            Scope = scope,
            Architecture = architecture,
            RegistryKeyPath = fullPath
        };
    }

    private static bool ToBool(object? value) => value is int i && i != 0;

    private static long? ToLong(object? value) => value switch
    {
        int i => i,
        long l => l,
        _ => null
    };

    private static string? NormalizePath(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return null;
        }
        return path.Trim('"', ' ');
    }

    private static DateTime? ParseInstallDate(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
        {
            return null;
        }

        // Most installers write InstallDate as "yyyyMMdd".
        if (raw.Length == 8 && DateTime.TryParseExact(raw, "yyyyMMdd", null,
                System.Globalization.DateTimeStyles.None, out var exact))
        {
            return exact;
        }

        return DateTime.TryParse(raw, out var parsed) ? parsed : null;
    }
}
