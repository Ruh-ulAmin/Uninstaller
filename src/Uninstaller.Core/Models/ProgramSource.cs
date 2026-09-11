namespace Uninstaller.Core.Models;

/// <summary>
/// Where an installed program entry was discovered.
/// </summary>
public enum ProgramSource
{
    Win32Registry,
    AppxPackage
}

/// <summary>
/// Registry scope a Win32 uninstall entry was found under.
/// </summary>
public enum ProgramScope
{
    Machine,
    User
}

/// <summary>
/// Bitness of the registry view an entry was read from.
/// </summary>
public enum ProgramArchitecture
{
    Unknown,
    Bit32,
    Bit64
}
