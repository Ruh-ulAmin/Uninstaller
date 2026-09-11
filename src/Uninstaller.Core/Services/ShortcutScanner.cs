namespace Uninstaller.Core.Services;

/// <summary>
/// Looks for Start Menu / Desktop shortcuts (.lnk / .url) whose file name
/// suggests they belong to a given program. Matching is name-based since
/// shortcuts carry no product identifier - callers must treat results as
/// candidates for user confirmation, not a guaranteed link to the program.
/// </summary>
public static class ShortcutScanner
{
    public static IEnumerable<string> FindShortcuts(string productName)
    {
        var nameTokens = productName
            .Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(t => t.Length > 2)
            .ToArray();

        if (nameTokens.Length == 0)
        {
            yield break;
        }

        foreach (var folder in ShortcutFolders())
        {
            if (!Directory.Exists(folder))
            {
                continue;
            }

            IEnumerable<string> files;
            try
            {
                files = Directory.EnumerateFiles(folder, "*.lnk", SearchOption.AllDirectories)
                    .Concat(Directory.EnumerateFiles(folder, "*.url", SearchOption.AllDirectories));
            }
            catch
            {
                continue;
            }

            foreach (var file in files)
            {
                var fileNameNoExt = Path.GetFileNameWithoutExtension(file);
                if (nameTokens.Any(token => fileNameNoExt.Contains(token, StringComparison.OrdinalIgnoreCase)))
                {
                    yield return file;
                }
            }
        }
    }

    private static IEnumerable<string> ShortcutFolders()
    {
        yield return Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu);
        yield return Environment.GetFolderPath(Environment.SpecialFolder.StartMenu);
        yield return Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        yield return Environment.GetFolderPath(Environment.SpecialFolder.CommonDesktopDirectory);
    }
}
