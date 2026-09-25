using Uninstaller.Core.Models;

namespace Uninstaller.Core.Services;

/// <summary>
/// System-wide scan for residual data left behind by programs that were
/// already uninstalled (or uninstalled outside this app). Everything found
/// here is a *candidate* for review - nothing is ever deleted by the scan
/// itself, and results should always be presented to the user with nothing
/// pre-selected.
/// </summary>
public sealed class LeftoverScannerService
{
    // Well-known top-level folder names that are never program leftovers.
    private static readonly HashSet<string> ExcludedFolderNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "Common Files", "Windows", "WindowsApps", "WindowsPowerShell",
        "Microsoft", "Microsoft.NET", "Internet Explorer", "ModifiableWindowsApps",
        "Package Cache", "Packages", "PackageManagement", "Temp", "Temporary Internet Files",
        "Google", "Mozilla", ".dotnet", ".nuget", "Comms", "Connected Devices Platform"
    };

    public IReadOnlyList<LeftoverItem> FindOrphanUninstallEntries(IEnumerable<InstalledProgram> installedPrograms)
    {
        var items = new List<LeftoverItem>();

        foreach (var program in installedPrograms)
        {
            if (program.Source != ProgramSource.Win32Registry || string.IsNullOrWhiteSpace(program.RegistryKeyPath))
            {
                continue;
            }

            var installLocationMissing =
                !string.IsNullOrWhiteSpace(program.InstallLocation) && !Directory.Exists(program.InstallLocation);

            var uninstallerMissing = false;
            if (!string.IsNullOrWhiteSpace(program.UninstallString))
            {
                var (fileName, _) = CommandLineParser.Split(program.UninstallString);
                if (!CommandLineParser.IsMsiExec(fileName) && !string.IsNullOrWhiteSpace(fileName))
                {
                    uninstallerMissing = !File.Exists(fileName);
                }
            }

            if (installLocationMissing && uninstallerMissing)
            {
                items.Add(new LeftoverItem
                {
                    Kind = LeftoverKind.OrphanUninstallEntry,
                    Path = program.RegistryKeyPath,
                    Description = $"Orphaned registry entry for '{program.DisplayName}' (install folder and uninstaller both missing)",
                    RegistryArchitecture = program.Architecture
                });
            }
        }

        return items;
    }

    public IReadOnlyList<LeftoverItem> FindOrphanFolders(IEnumerable<InstalledProgram> installedPrograms)
    {
        var installedTokens = BuildInstalledNameTokens(installedPrograms);
        var installedLocations = installedPrograms
            .Select(p => p.InstallLocation)
            .Where(p => !string.IsNullOrWhiteSpace(p))
            .Select(p => p!.TrimEnd('\\'))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var items = new List<LeftoverItem>();

        foreach (var root in CandidateRoots())
        {
            if (!Directory.Exists(root))
            {
                continue;
            }

            IEnumerable<string> subDirs;
            try
            {
                subDirs = Directory.EnumerateDirectories(root);
            }
            catch
            {
                continue;
            }

            foreach (var dir in subDirs)
            {
                var name = Path.GetFileName(dir);
                if (ExcludedFolderNames.Contains(name))
                {
                    continue;
                }

                if (installedLocations.Contains(dir.TrimEnd('\\')))
                {
                    continue;
                }

                if (installedTokens.Any(token => name.Contains(token, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                if (!PathSafetyGuard.IsSafeToDeleteRecursively(dir))
                {
                    continue;
                }

                items.Add(new LeftoverItem
                {
                    Kind = LeftoverKind.Folder,
                    Path = dir,
                    Description = $"Folder not linked to any installed program ({root})"
                });
            }
        }

        return items;
    }

    private static HashSet<string> BuildInstalledNameTokens(IEnumerable<InstalledProgram> installedPrograms)
    {
        var tokens = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var program in installedPrograms)
        {
            foreach (var token in program.DisplayName.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                if (token.Length > 2)
                {
                    tokens.Add(token);
                }
            }

            if (!string.IsNullOrWhiteSpace(program.Publisher))
            {
                foreach (var token in program.Publisher.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
                {
                    if (token.Length > 2)
                    {
                        tokens.Add(token);
                    }
                }
            }
        }
        return tokens;
    }

    private static IEnumerable<string> CandidateRoots()
    {
        yield return Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        yield return Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
        yield return Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData); // ProgramData
        yield return Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        yield return Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData); // Roaming
    }
}
