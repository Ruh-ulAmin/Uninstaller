using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media.Imaging;

namespace Uninstaller.App.Helpers;

/// <summary>
/// Extracts a small icon for a file/executable using Shell32, avoiding a
/// dependency on System.Drawing.Common. Results are cached per source path
/// for the lifetime of the process since icon lookups hit the file system.
/// </summary>
public static class IconExtractor
{
    private const uint SHGFI_ICON = 0x000000100;
    private const uint SHGFI_SMALLICON = 0x000000001;
    private const uint SHGFI_LARGEICON = 0x000000000;
    private const uint SHGFI_USEFILEATTRIBUTES = 0x000000010;
    private const uint FILE_ATTRIBUTE_NORMAL = 0x80;

    [StructLayout(LayoutKind.Sequential)]
    private struct SHFILEINFO
    {
        public IntPtr hIcon;
        public int iIcon;
        public uint dwAttributes;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string szDisplayName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
        public string szTypeName;
    }

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr SHGetFileInfo(string pszPath, uint dwFileAttributes, ref SHFILEINFO psfi, uint cbFileInfo, uint uFlags);

    [DllImport("user32.dll")]
    private static extern bool DestroyIcon(IntPtr hIcon);

    private static readonly Dictionary<string, BitmapSource?> Cache = new(StringComparer.OrdinalIgnoreCase);
    private static readonly BitmapSource? DefaultIcon = LoadGenericIcon();

    public static BitmapSource GetIconFor(string? filePath, bool small = true)
    {
        var key = $"{filePath}|{small}";
        if (Cache.TryGetValue(key, out var cached) && cached is not null)
        {
            return cached;
        }

        var resolved = TryExtract(filePath, small) ?? DefaultIcon;
        Cache[key] = resolved;
        return resolved ?? BuildFallbackBitmap();
    }

    private static BitmapSource? TryExtract(string? filePath, bool small)
    {
        var useFileAttributes = string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath);
        var path = useFileAttributes ? ".exe" : filePath!;
        var flags = SHGFI_ICON | (small ? SHGFI_SMALLICON : SHGFI_LARGEICON) | (useFileAttributes ? SHGFI_USEFILEATTRIBUTES : 0);

        var info = new SHFILEINFO();
        var result = SHGetFileInfo(path, FILE_ATTRIBUTE_NORMAL, ref info, (uint)Marshal.SizeOf<SHFILEINFO>(), flags);

        if (result == IntPtr.Zero || info.hIcon == IntPtr.Zero)
        {
            return null;
        }

        try
        {
            var bitmapSource = Imaging.CreateBitmapSourceFromHIcon(info.hIcon, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
            bitmapSource.Freeze();
            return bitmapSource;
        }
        catch
        {
            return null;
        }
        finally
        {
            DestroyIcon(info.hIcon);
        }
    }

    private static BitmapSource? LoadGenericIcon() => TryExtract(null, small: true);

    private static BitmapSource BuildFallbackBitmap()
    {
        // Last-resort 16x16 transparent bitmap so bindings never fail even if
        // shell32 lookups are unavailable (e.g. locked-down environments).
        var bitmap = new WriteableBitmap(16, 16, 96, 96, System.Windows.Media.PixelFormats.Bgra32, null);
        bitmap.Freeze();
        return bitmap;
    }
}
