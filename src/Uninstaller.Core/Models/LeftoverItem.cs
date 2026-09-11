namespace Uninstaller.Core.Models;

public enum LeftoverKind
{
    Folder,
    RegistryKey,
    Shortcut,
    OrphanUninstallEntry
}

/// <summary>
/// A candidate piece of residual data found by the leftover scanner.
/// Nothing represented by this type is ever deleted automatically -
/// the user must explicitly select and confirm removal in the UI.
/// </summary>
public sealed class LeftoverItem
{
    public required LeftoverKind Kind { get; init; }
    public required string Path { get; init; }
    public string? Description { get; init; }
    public long? SizeBytes { get; init; }
    public string SizeDisplay => SizeBytes is > 0 ? InstalledProgram.FormatSize(SizeBytes.Value) : "-";
}
