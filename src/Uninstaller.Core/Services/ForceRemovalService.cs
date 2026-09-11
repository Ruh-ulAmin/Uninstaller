using Microsoft.Win32;
using Uninstaller.Core.Models;

namespace Uninstaller.Core.Services;

/// <summary>
/// Handles the "force remove" path used when a program's own uninstaller is
/// missing, broken, or refuses to run: it discovers what a normal uninstall
/// would have removed (registry key, install folder, shortcuts) and deletes
/// each piece only after the caller has confirmed it.
/// </summary>
public sealed class ForceRemovalService
{
    private readonly FileLogger _logger;

    public ForceRemovalService(FileLogger logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Finds everything associated with a program that a force-removal could
    /// clean up, without deleting anything yet.
    /// </summary>
    public IReadOnlyList<LeftoverItem> FindResidue(InstalledProgram program)
    {
        var items = new List<LeftoverItem>();

        if (!string.IsNullOrWhiteSpace(program.RegistryKeyPath))
        {
            items.Add(new LeftoverItem
            {
                Kind = LeftoverKind.RegistryKey,
                Path = program.RegistryKeyPath,
                Description = "Uninstall registry entry"
            });
        }

        if (!string.IsNullOrWhiteSpace(program.InstallLocation) && Directory.Exists(program.InstallLocation))
        {
            items.Add(new LeftoverItem
            {
                Kind = LeftoverKind.Folder,
                Path = program.InstallLocation,
                Description = "Installation folder",
                SizeBytes = TryGetDirectorySize(program.InstallLocation)
            });
        }

        foreach (var shortcut in ShortcutScanner.FindShortcuts(program.DisplayName))
        {
            items.Add(new LeftoverItem
            {
                Kind = LeftoverKind.Shortcut,
                Path = shortcut,
                Description = "Shortcut"
            });
        }

        return items;
    }

    /// <summary>
    /// Deletes exactly the items passed in - the caller is responsible for
    /// only including items the user explicitly selected.
    /// </summary>
    public IReadOnlyList<OperationResult> Remove(IEnumerable<LeftoverItem> items)
    {
        var results = new List<OperationResult>();

        foreach (var item in items)
        {
            try
            {
                switch (item.Kind)
                {
                    case LeftoverKind.RegistryKey:
                        RemoveRegistryKey(item.Path);
                        break;
                    case LeftoverKind.Folder:
                        Directory.Delete(item.Path, recursive: true);
                        break;
                    case LeftoverKind.Shortcut:
                    case LeftoverKind.OrphanUninstallEntry:
                        File.Delete(item.Path);
                        break;
                }

                var msg = $"Removed {item.Description ?? item.Kind.ToString()}: {item.Path}";
                _logger.Log(msg);
                results.Add(OperationResult.Ok(item.Path, msg));
            }
            catch (Exception ex)
            {
                var msg = $"Failed to remove {item.Path}: {ex.Message}";
                _logger.Log(msg);
                results.Add(OperationResult.Fail(item.Path, msg));
            }
        }

        return results;
    }

    private static void RemoveRegistryKey(string fullPath)
    {
        var (hive, view, subPath) = SplitRegistryPath(fullPath);
        using var baseKey = RegistryKey.OpenBaseKey(hive, view);
        baseKey.DeleteSubKeyTree(subPath, throwOnMissingSubKey: false);
    }

    internal static (RegistryHive Hive, RegistryView View, string SubPath) SplitRegistryPath(string fullPath)
    {
        var separatorIndex = fullPath.IndexOf('\\');
        var rootName = separatorIndex > 0 ? fullPath[..separatorIndex] : fullPath;
        var subPath = separatorIndex > 0 ? fullPath[(separatorIndex + 1)..] : string.Empty;

        var hive = rootName switch
        {
            "HKEY_LOCAL_MACHINE" => RegistryHive.LocalMachine,
            "HKEY_CURRENT_USER" => RegistryHive.CurrentUser,
            "HKEY_CLASSES_ROOT" => RegistryHive.ClassesRoot,
            "HKEY_USERS" => RegistryHive.Users,
            _ => RegistryHive.LocalMachine
        };

        // Registry64 always resolves the "real" path this string already points at.
        return (hive, RegistryView.Registry64, subPath);
    }

    private static long? TryGetDirectorySize(string path)
    {
        try
        {
            long total = 0;
            foreach (var file in Directory.EnumerateFiles(path, "*", SearchOption.AllDirectories))
            {
                try
                {
                    total += new FileInfo(file).Length;
                }
                catch
                {
                    // Inaccessible file - skip it, size is only informational.
                }
            }
            return total;
        }
        catch
        {
            return null;
        }
    }
}
