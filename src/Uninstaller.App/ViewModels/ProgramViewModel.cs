using System.Windows.Media.Imaging;
using Uninstaller.App.Helpers;
using Uninstaller.Core.Models;

namespace Uninstaller.App.ViewModels;

/// <summary>
/// UI-facing wrapper around <see cref="InstalledProgram"/> that adds
/// selection state and a lazily-resolved icon for the programs grid.
/// </summary>
public sealed class ProgramViewModel : ViewModelBase
{
    private bool _isChecked;
    private BitmapSource? _icon;

    public ProgramViewModel(InstalledProgram program)
    {
        Program = program;
    }

    public InstalledProgram Program { get; }

    public bool IsChecked
    {
        get => _isChecked;
        set => SetField(ref _isChecked, value);
    }

    public string DisplayName => Program.DisplayName;
    public string? Publisher => string.IsNullOrWhiteSpace(Program.Publisher) ? "Unknown" : Program.Publisher;
    public string? DisplayVersion => string.IsNullOrWhiteSpace(Program.DisplayVersion) ? "-" : Program.DisplayVersion;
    public string SizeDisplay => Program.SizeDisplay;
    public string ArchitectureDisplay => Program.ArchitectureDisplay;
    public string InstallDateDisplay => Program.InstallDate?.ToString("yyyy-MM-dd") ?? "-";
    public bool IsStoreApp => Program.Source == ProgramSource.AppxPackage;
    public bool HasModify => !string.IsNullOrWhiteSpace(Program.ModifyPath);

    public BitmapSource Icon => _icon ??= IconExtractor.GetIconFor(ResolveIconSourcePath());

    private string? ResolveIconSourcePath()
    {
        if (!string.IsNullOrWhiteSpace(Program.DisplayIconPath))
        {
            // DisplayIcon can be "path,index" - take just the path portion.
            var commaIndex = Program.DisplayIconPath.LastIndexOf(',');
            var path = commaIndex > 0 ? Program.DisplayIconPath[..commaIndex] : Program.DisplayIconPath;
            return path.Trim('"');
        }

        return Program.InstallLocation;
    }
}
