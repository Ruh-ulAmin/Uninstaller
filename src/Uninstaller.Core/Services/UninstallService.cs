using System.Diagnostics;
using Uninstaller.Core.Models;

namespace Uninstaller.Core.Services;

/// <summary>
/// Runs the actual uninstall command for a program: either its
/// vendor-provided uninstaller, the MSI engine, or (for Store apps)
/// PowerShell's AppX cmdlets.
/// </summary>
public sealed class UninstallService
{
    private readonly FileLogger _logger;

    public UninstallService(FileLogger logger)
    {
        _logger = logger;
    }

    public async Task<OperationResult> UninstallAsync(InstalledProgram program, bool silent, IProgress<string>? progress = null, CancellationToken cancellationToken = default)
    {
        progress?.Report($"Starting uninstall of '{program.DisplayName}'...");
        _logger.Log($"Uninstall requested: {program.DisplayName} (silent={silent}, source={program.Source})");

        try
        {
            return program.Source switch
            {
                ProgramSource.AppxPackage => await UninstallAppxAsync(program, progress, cancellationToken).ConfigureAwait(false),
                _ => await UninstallWin32Async(program, silent, progress, cancellationToken).ConfigureAwait(false)
            };
        }
        catch (OperationCanceledException)
        {
            var msg = $"Uninstall of '{program.DisplayName}' was cancelled.";
            progress?.Report(msg);
            _logger.Log(msg);
            return OperationResult.Fail(program.DisplayName, msg);
        }
        catch (Exception ex)
        {
            var msg = $"Error uninstalling '{program.DisplayName}': {ex.Message}";
            progress?.Report(msg);
            _logger.Log(msg);
            return OperationResult.Fail(program.DisplayName, msg);
        }
    }

    private async Task<OperationResult> UninstallWin32Async(InstalledProgram program, bool silent, IProgress<string>? progress, CancellationToken cancellationToken)
    {
        var commandLine = silent && !string.IsNullOrWhiteSpace(program.QuietUninstallString)
            ? program.QuietUninstallString!
            : program.UninstallString;

        if (string.IsNullOrWhiteSpace(commandLine))
        {
            var msg = $"'{program.DisplayName}' has no uninstall command registered.";
            progress?.Report(msg);
            _logger.Log(msg);
            return OperationResult.Fail(program.DisplayName, msg);
        }

        var (fileName, arguments) = CommandLineParser.Split(commandLine);

        if (CommandLineParser.IsMsiExec(fileName) && silent)
        {
            arguments = CommandLineParser.MakeMsiSilent(arguments);
        }

        progress?.Report($"Running: {fileName} {arguments}");
        _logger.Log($"Command: \"{fileName}\" {arguments}");

        var psi = new ProcessStartInfo
        {
            FileName = fileName,
            Arguments = arguments,
            UseShellExecute = true,
            Verb = "runas",
            WindowStyle = silent ? ProcessWindowStyle.Hidden : ProcessWindowStyle.Normal
        };

        using var process = new Process { StartInfo = psi };
        process.Start();
        await process.WaitForExitAsync(cancellationToken).ConfigureAwait(false);

        var exitCode = process.ExitCode;
        var rebootRequired = exitCode == 3010;
        var success = exitCode == 0 || rebootRequired;

        var resultMsg = success
            ? $"'{program.DisplayName}' uninstalled successfully.{(rebootRequired ? " A reboot is required to finish." : string.Empty)}"
            : $"Uninstaller for '{program.DisplayName}' exited with code {exitCode}.";

        progress?.Report(resultMsg);
        _logger.Log(resultMsg);

        return success
            ? OperationResult.Ok(program.DisplayName, resultMsg, exitCode, rebootRequired)
            : OperationResult.Fail(program.DisplayName, resultMsg, exitCode);
    }

    private async Task<OperationResult> UninstallAppxAsync(InstalledProgram program, IProgress<string>? progress, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(program.PackageFullName))
        {
            var msg = $"'{program.DisplayName}' is missing its package identity and cannot be removed.";
            progress?.Report(msg);
            return OperationResult.Fail(program.DisplayName, msg);
        }

        var psCommand = $"Remove-AppxPackage -Package '{program.PackageFullName}' -ErrorAction Stop";
        progress?.Report($"Running: powershell.exe -Command \"{psCommand}\"");
        _logger.Log($"AppX removal command: {psCommand}");

        var psi = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = $"-NoProfile -NonInteractive -Command \"{psCommand}\"",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        using var process = new Process { StartInfo = psi };
        process.Start();
        var stdErrTask = process.StandardError.ReadToEndAsync(cancellationToken);
        var stdOutTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
        await process.WaitForExitAsync(cancellationToken).ConfigureAwait(false);
        var stdErr = await stdErrTask.ConfigureAwait(false);
        var stdOut = await stdOutTask.ConfigureAwait(false);

        var success = process.ExitCode == 0 && string.IsNullOrWhiteSpace(stdErr);
        var resultMsg = success
            ? $"'{program.DisplayName}' (Store app) removed successfully."
            : $"Failed to remove '{program.DisplayName}': {(string.IsNullOrWhiteSpace(stdErr) ? stdOut : stdErr)}";

        progress?.Report(resultMsg);
        _logger.Log(resultMsg);

        return success
            ? OperationResult.Ok(program.DisplayName, resultMsg, process.ExitCode)
            : OperationResult.Fail(program.DisplayName, resultMsg, process.ExitCode);
    }
}
