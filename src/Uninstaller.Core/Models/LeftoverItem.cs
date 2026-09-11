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

    /// <summary>
    /// For <see cref="LeftoverKind.RegistryKey"/> and
    /// <see cref="LeftoverKind.OrphanUninstallEntry"/> only: which registry
    /// view (32/64-bit) the key was actually read from. WOW64 redirection
    /// means a 32-bit key's reported path looks identical to its 64-bit
    /// counterpart, so the view must travel with the item or deletion will
    /// silently target the wrong (non-existent) physical key.
    /// </summary>
    public ProgramArchitecture RegistryArchitecture { get; init; } = ProgramArchitecture.Bit64;

    public string SizeDisplay => SizeBytes is > 0 ? InstalledProgram.FormatSize(SizeBytes.Value) : "-";
}
