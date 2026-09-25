namespace Uninstaller.Core.Services;

/// <summary>
/// Defense-in-depth checks applied immediately before any destructive
/// filesystem or registry delete. Every path/registry key this app deletes
/// ultimately traces back to a string read from the registry (an
/// InstallLocation, an UninstallString-derived folder, an uninstall key
/// path) - all of which HKEY_CURRENT_USER entries let an unprivileged,
/// non-admin process set to whatever it likes. These checks make sure a
/// malformed or malicious value can never cause a catastrophic delete
/// (a drive root, "C:\Windows", "C:\Program Files" itself, etc.),
/// independent of who could have influenced the value or why.
/// </summary>
public static class PathSafetyGuard
{
    // Top-level directories where NOTHING is ever eligible for deletion,
    // at any depth - a legitimate program InstallLocation is never inside
    // any of these, and the damage potential (OS files, another user's
    // entire profile) is severe enough that no depth is deep enough to
    // trust. Matched against just the first path segment below the drive
    // root, case-insensitively.
    private static readonly string[] NeverAllowedTopLevel =
    {
        "windows", "system32", "syswow64", "perflogs", "$recycle.bin",
        "boot", "recovery", "windowsapps", "system volume information"
    };

    // Top-level directories that ARE legitimate install locations (so a
    // subfolder within them is fine), but where the directory itself, or
    // anything too shallow beneath it, must never be deleted.
    private static readonly string[] ProtectedDirectoryNames =
    {
        "windows", "program files", "program files (x86)", "programdata", "users"
    };

    /// <summary>
    /// True if it is safe to recursively delete <paramref name="path"/>:
    /// it must be a fully-rooted path with meaningful depth below a drive
    /// root, and must not itself be one of the well-known top-level system
    /// directories (deleting a subfolder *within* Program Files is fine -
    /// deleting Program Files itself, or the drive root, is not).
    /// </summary>
    public static bool IsSafeToDeleteRecursively(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return false;
        }

        string fullPath;
        try
        {
            fullPath = Path.GetFullPath(path);
        }
        catch
        {
            return false;
        }

        var root = Path.GetPathRoot(fullPath);
        if (string.IsNullOrEmpty(root) || !fullPath.StartsWith(root, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var relative = fullPath[root.Length..].Trim(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        if (relative.Length == 0)
        {
            // The drive root itself (e.g. "C:\").
            return false;
        }

        var segments = relative.Split(
            new[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar },
            StringSplitOptions.RemoveEmptyEntries);

        if (segments.Length == 0)
        {
            return false;
        }

        // C:\Windows and friends: never eligible, at any depth. A genuine
        // program is never installed inside any of these.
        if (NeverAllowedTopLevel.Contains(segments[0], StringComparer.OrdinalIgnoreCase))
        {
            return false;
        }

        // Refuse to delete a well-known top-level directory itself
        // (e.g. "C:\Program Files") - a subfolder inside one, like
        // "C:\Program Files\Vendor\App", is a normal install location.
        if (segments.Length == 1 && ProtectedDirectoryNames.Contains(segments[0], StringComparer.OrdinalIgnoreCase))
        {
            return false;
        }

        // "C:\Users\<name>" is an entire user profile, not a program - an
        // InstallLocation this shallow under Users would delete someone's
        // whole profile (documents, desktop, everything). Require a real
        // folder within the profile instead.
        if (string.Equals(segments[0], "users", StringComparison.OrdinalIgnoreCase) && segments.Length < 3)
        {
            return false;
        }

        // Otherwise require at least two path segments below the drive
        // root (e.g. "Program Files\Vendor App") so a value that somehow
        // collapsed to just a top-level folder can't wipe out an entire
        // category of machine-wide data even under a name not listed above.
        return segments.Length >= 2;
    }

    /// <summary>
    /// True if it is safe to delete the registry key at <paramref name="subPath"/>
    /// (relative to its hive root). Only keys that are actually nested under
    /// an "Uninstall" key - the one place this app ever discovers keys to
    /// remove - are ever eligible, which rules out a malformed or truncated
    /// path collapsing to something shallow like "Software" or
    /// "Software\Microsoft\Windows".
    /// </summary>
    public static bool IsSafeToDeleteRegistryKey(string subPath)
    {
        if (string.IsNullOrWhiteSpace(subPath))
        {
            return false;
        }

        var segments = subPath.Split('\\', StringSplitOptions.RemoveEmptyEntries);

        // Must end in .../Uninstall/<something>, with the key itself (the
        // last segment) non-empty - i.e. a specific program's subkey, never
        // the Uninstall key itself.
        return segments.Length >= 2 &&
               string.Equals(segments[^2], "Uninstall", StringComparison.OrdinalIgnoreCase);
    }
}
