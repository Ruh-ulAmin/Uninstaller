using System.Diagnostics;

namespace Uninstaller.Core.Services;

public sealed record RunningProcessInfo(int ProcessId, string ProcessName, string ExecutablePath);

/// <summary>
/// Detects processes currently running from inside a given folder, so the
/// UI can warn before deleting it - files locked by a running process can't
/// be deleted anyway, and force-removing a program that's still open can
/// leave it in a broken, half-uninstalled state.
/// </summary>
public static class RunningProcessChecker
{
    public static IReadOnlyList<RunningProcessInfo> FindProcessesUnder(string folderPath)
    {
        var results = new List<RunningProcessInfo>();
        if (string.IsNullOrWhiteSpace(folderPath))
        {
            return results;
        }

        string normalizedFolder;
        try
        {
            normalizedFolder = Path.GetFullPath(folderPath).TrimEnd('\\', '/') + Path.DirectorySeparatorChar;
        }
        catch
        {
            return results;
        }

        foreach (var process in Process.GetProcesses())
        {
            try
            {
                var path = process.MainModule?.FileName;
                if (!string.IsNullOrEmpty(path) && path.StartsWith(normalizedFolder, StringComparison.OrdinalIgnoreCase))
                {
                    results.Add(new RunningProcessInfo(process.Id, process.ProcessName, path));
                }
            }
            catch
            {
                // Processes we don't own, are cross-session, or are system
                // processes throw here (access denied) - not something we
                // could act on anyway, so they're simply skipped.
            }
            finally
            {
                process.Dispose();
            }
        }

        return results;
    }

    /// <summary>
    /// Attempts to close the given processes, trying a graceful close first
    /// and only force-killing ones that don't exit promptly. Returns the
    /// processes that could not be stopped.
    /// </summary>
    public static IReadOnlyList<RunningProcessInfo> TryStop(IEnumerable<RunningProcessInfo> processes, TimeSpan gracePeriod)
    {
        var stillRunning = new List<RunningProcessInfo>();

        foreach (var info in processes)
        {
            try
            {
                using var process = Process.GetProcessById(info.ProcessId);
                process.CloseMainWindow();
                if (!process.WaitForExit((int)gracePeriod.TotalMilliseconds))
                {
                    process.Kill(entireProcessTree: true);
                    process.WaitForExit(2000);
                }
            }
            catch (ArgumentException)
            {
                // Already exited between the scan and now - fine.
            }
            catch
            {
                stillRunning.Add(info);
            }
        }

        return stillRunning;
    }
}
