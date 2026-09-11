namespace Uninstaller.Core.Services;

/// <summary>
/// Simple thread-safe append-only text logger. Every uninstall / cleanup
/// action taken by the app is recorded here so the user has an audit trail.
/// </summary>
public sealed class FileLogger
{
    private readonly object _lock = new();
    private readonly string _logFilePath;

    public string LogFilePath => _logFilePath;
    public string LogFolder { get; }

    public FileLogger()
    {
        LogFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Uninstaller", "Logs");
        Directory.CreateDirectory(LogFolder);
        _logFilePath = Path.Combine(LogFolder, $"uninstaller-{DateTime.Now:yyyy-MM-dd}.log");
    }

    public void Log(string message)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] {message}";
        lock (_lock)
        {
            try
            {
                File.AppendAllLines(_logFilePath, new[] { line });
            }
            catch
            {
                // Logging must never crash the app.
            }
        }
    }
}
