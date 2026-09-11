using System.Text;

namespace Uninstaller.Core.Services;

/// <summary>
/// Splits a raw uninstall command line (as stored in the registry) into an
/// executable path and an argument string, honoring quoted paths.
/// </summary>
public static class CommandLineParser
{
    public static (string FileName, string Arguments) Split(string commandLine)
    {
        commandLine = commandLine.Trim();

        if (commandLine.Length == 0)
        {
            return (string.Empty, string.Empty);
        }

        if (commandLine[0] == '"')
        {
            var closingQuote = commandLine.IndexOf('"', 1);
            if (closingQuote > 0)
            {
                var file = commandLine.Substring(1, closingQuote - 1);
                var rest = closingQuote + 1 < commandLine.Length
                    ? commandLine[(closingQuote + 1)..].TrimStart()
                    : string.Empty;
                return (file, rest);
            }
        }

        // Unquoted, with no closing quote found above: some installers (older
        // NSIS/InstallShield ones especially) register UninstallString without
        // quoting the path even when it contains spaces, e.g.
        // "C:\Program Files\Vendor App\uninstall.exe -s". Splitting on the
        // first space would misparse that as fileName="C:\Program". Instead,
        // try the space-delimited prefixes from longest to shortest and use
        // the first one that actually exists on disk.
        var spaceIndices = new List<int>();
        for (var i = 0; i < commandLine.Length; i++)
        {
            if (commandLine[i] == ' ')
            {
                spaceIndices.Add(i);
            }
        }

        for (var i = spaceIndices.Count - 1; i >= 0; i--)
        {
            var candidate = commandLine[..spaceIndices[i]];
            if (File.Exists(candidate))
            {
                return (candidate, commandLine[(spaceIndices[i] + 1)..].TrimStart());
            }
        }

        if (spaceIndices.Count == 0)
        {
            return (commandLine, string.Empty);
        }

        return (commandLine[..spaceIndices[0]], commandLine[(spaceIndices[0] + 1)..].TrimStart());
    }

    /// <summary>
    /// Rewrites an MSI-based uninstall command to run fully silently
    /// (adds /qn and suppresses restarts) without altering the product code.
    /// </summary>
    public static string MakeMsiSilent(string arguments)
    {
        var sb = new StringBuilder(arguments);
        if (!arguments.Contains("/qn", StringComparison.OrdinalIgnoreCase) &&
            !arguments.Contains("/quiet", StringComparison.OrdinalIgnoreCase))
        {
            sb.Append(" /qn");
        }
        if (!arguments.Contains("/norestart", StringComparison.OrdinalIgnoreCase))
        {
            sb.Append(" /norestart");
        }
        return sb.ToString();
    }

    public static bool IsMsiExec(string fileName) =>
        Path.GetFileName(fileName).Equals("msiexec.exe", StringComparison.OrdinalIgnoreCase);
}
