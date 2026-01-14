using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace DeskGrid.Services;

/// <summary>
/// Extracts high-resolution icons from files and folders
/// </summary>
public static class IconLoader
{
    private static readonly Dictionary<string, ImageSource> _iconCache = new();
    private static readonly object _cacheLock = new();

    /// <summary>
    /// Gets an icon for a file or folder path
    /// </summary>
    public static ImageSource? GetIcon(string path, int size = 48)
    {
        if (string.IsNullOrEmpty(path))
            return null;

        // Check cache first
        string cacheKey = $"{path}_{size}";
        lock (_cacheLock)
        {
            if (_iconCache.TryGetValue(cacheKey, out var cached))
                return cached;
        }

        ImageSource? icon = null;

        try
        {
            // Try IShellItemImageFactory for high-quality icons
            icon = GetIconViaShellItem(path, size);
        }
        catch
        {
            // Fallback to SHGetFileInfo
            try
            {
                icon = GetIconViaSHGetFileInfo(path, size > 32);
            }
            catch
            {
                // Last resort - return null
            }
        }

        // Cache the result
        if (icon != null)
        {
            lock (_cacheLock)
            {
                _iconCache[cacheKey] = icon;
            }
        }

        return icon;
    }

    private static ImageSource? GetIconViaShellItem(string path, int size)
    {
        // Create shell item
        var guid = typeof(IShellItemImageFactory).GUID;
        int hr = SHCreateItemFromParsingName(path, IntPtr.Zero, ref guid, out var imageFactory);

        if (hr != 0 || imageFactory == null)
            return null;

        try
        {
            // Get the icon bitmap
            hr = imageFactory.GetImage(new SIZE(size, size), SIIGBF.SIIGBF_RESIZETOFIT, out var hBitmap);

            if (hr != 0 || hBitmap == IntPtr.Zero)
                return null;

            try
            {
                // Convert HBITMAP to ImageSource
                return ConvertHBitmapToImageSource(hBitmap);
            }
            finally
            {
                DeleteObject(hBitmap);
            }
        }
        finally
        {
            Marshal.ReleaseComObject(imageFactory);
        }
    }

    private static ImageSource? GetIconViaSHGetFileInfo(string path, bool largeIcon)
    {
        var shinfo = new SHFILEINFO();
        uint flags = SHGFI_ICON | (largeIcon ? SHGFI_LARGEICON : SHGFI_SMALLICON);

        var result = SHGetFileInfo(path, 0, ref shinfo, (uint)Marshal.SizeOf(shinfo), flags);

        if (result == IntPtr.Zero || shinfo.hIcon == IntPtr.Zero)
            return null;

        try
        {
            var icon = System.Drawing.Icon.FromHandle(shinfo.hIcon);
            return Imaging.CreateBitmapSourceFromHIcon(
                icon.Handle,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());
        }
        finally
        {
            DestroyIcon(shinfo.hIcon);
        }
    }

    private static ImageSource ConvertHBitmapToImageSource(IntPtr hBitmap)
    {
        try
        {
            // Get bitmap info
            var bmp = new BITMAP();
            GetObject(hBitmap, Marshal.SizeOf(typeof(BITMAP)), ref bmp);
            
            if (bmp.bmWidth == 0 || bmp.bmHeight == 0)
                return null!;
            
            // Create a WPF-compatible bitmap with proper alpha support
            var bitmapSource = Imaging.CreateBitmapSourceFromHBitmap(
                hBitmap,
                IntPtr.Zero,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());
            
            // Convert to BGRA format to ensure alpha is preserved
            var formattedBitmap = new FormatConvertedBitmap(bitmapSource, PixelFormats.Bgra32, null, 0);
            formattedBitmap.Freeze();
            
            return formattedBitmap;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[IconLoader] Icon conversion failed: {ex.Message}");
            return null!;
        }
    }
    
    [StructLayout(LayoutKind.Sequential)]
    private struct BITMAP
    {
        public int bmType;
        public int bmWidth;
        public int bmHeight;
        public int bmWidthBytes;
        public short bmPlanes;
        public short bmBitsPixel;
        public IntPtr bmBits;
    }
    
    [DllImport("gdi32.dll")]
    private static extern int GetObject(IntPtr hObject, int nCount, ref BITMAP lpObject);

    /// <summary>
    /// Clears the icon cache
    /// </summary>
    public static void ClearCache()
    {
        lock (_cacheLock)
        {
            _iconCache.Clear();
        }
    }

    #region P/Invoke

    private const uint SHGFI_ICON = 0x000000100;
    private const uint SHGFI_LARGEICON = 0x000000000;
    private const uint SHGFI_SMALLICON = 0x000000001;

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
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

    [DllImport("shell32.dll", CharSet = CharSet.Auto)]
    private static extern IntPtr SHGetFileInfo(string pszPath, uint dwFileAttributes, 
        ref SHFILEINFO psfi, uint cbFileInfo, uint uFlags);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool DestroyIcon(IntPtr hIcon);

    [DllImport("gdi32.dll")]
    private static extern bool DeleteObject(IntPtr hObject);

    [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
    private static extern int SHCreateItemFromParsingName(
        [MarshalAs(UnmanagedType.LPWStr)] string pszPath,
        IntPtr pbc,
        ref Guid riid,
        [MarshalAs(UnmanagedType.Interface)] out IShellItemImageFactory ppv);

    [ComImport]
    [Guid("bcc18b79-ba16-442f-80c4-8a59c30c463b")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellItemImageFactory
    {
        [PreserveSig]
        int GetImage(SIZE size, SIIGBF flags, out IntPtr phbm);
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct SIZE
    {
        public int cx;
        public int cy;

        public SIZE(int width, int height)
        {
            cx = width;
            cy = height;
        }
    }

    [Flags]
    private enum SIIGBF
    {
        SIIGBF_RESIZETOFIT = 0x00000000,
        SIIGBF_BIGGERSIZEOK = 0x00000001,
        SIIGBF_MEMORYONLY = 0x00000002,
        SIIGBF_ICONONLY = 0x00000004,
        SIIGBF_THUMBNAILONLY = 0x00000008,
        SIIGBF_INCACHEONLY = 0x00000010,
    }

    #endregion
}
