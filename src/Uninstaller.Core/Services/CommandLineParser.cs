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

        // Unquoted: split on the first space that is followed by something
        // that looks like a flag or another path, falling back to the first
        // whitespace if we can't be clever about it.
        var spaceIndex = commandLine.IndexOf(' ');
        if (spaceIndex < 0)
        {
            return (commandLine, string.Empty);
        }

        return (commandLine[..spaceIndex], commandLine[(spaceIndex + 1)..].TrimStart());
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
