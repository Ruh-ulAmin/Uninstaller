namespace Uninstaller.Core.Models;

/// <summary>
/// Outcome of a single uninstall / force-removal / cleanup operation.
/// </summary>
public sealed class OperationResult
{
    public required string TargetName { get; init; }
    public bool Success { get; init; }
    public int? ExitCode { get; init; }
    public string? Message { get; init; }
    public bool RebootRequired { get; init; }

    public static OperationResult Ok(string target, string message, int? exitCode = null, bool rebootRequired = false) =>
        new() { TargetName = target, Success = true, Message = message, ExitCode = exitCode, RebootRequired = rebootRequired };

    public static OperationResult Fail(string target, string message, int? exitCode = null) =>
        new() { TargetName = target, Success = false, Message = message, ExitCode = exitCode };
}
