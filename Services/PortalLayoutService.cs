using System.Diagnostics;
using System.IO;
using System.Text.Json;
using DeskGrid.Models;

namespace DeskGrid.Services;

/// <summary>
/// Service for saving and loading portal layouts
/// </summary>
public class PortalLayoutService
{
    private readonly string _configPath;
    private static readonly JsonSerializerOptions _jsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true
    };

    public PortalLayoutService()
    {
        // Use AppData for reliable persistence across sessions
        var appDataDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        var configDir = Path.Combine(appDataDir, "DeskGrid", "config");
        Directory.CreateDirectory(configDir);
        _configPath = Path.Combine(configDir, "layout.json");
        System.Diagnostics.Debug.WriteLine($"[PortalLayoutService] Config path: {_configPath}");
    }

    /// <summary>
    /// Loads the portal configuration from disk
    /// </summary>
    public DeskGridConfig LoadConfig()
    {
        Debug.WriteLine($"[PortalLayoutService] === LOADING CONFIG ===");
        Debug.WriteLine($"[PortalLayoutService] Config path: {_configPath}");
        Debug.WriteLine($"[PortalLayoutService] File exists: {File.Exists(_configPath)}");
        
        try
        {
            if (File.Exists(_configPath))
            {
                var json = File.ReadAllText(_configPath);
                Debug.WriteLine($"[PortalLayoutService] JSON length: {json.Length} chars");
                Debug.WriteLine($"[PortalLayoutService] JSON preview: {json.Substring(0, Math.Min(500, json.Length))}...");
                
                var config = JsonSerializer.Deserialize<DeskGridConfig>(json, _jsonOptions);
                if (config != null)
                {
                    Debug.WriteLine($"[PortalLayoutService] Loaded {config.Portals.Count} portals, portalsVisible: {config.PortalsVisible}");
                    foreach (var portal in config.Portals)
                    {
                        Debug.WriteLine($"[PortalLayoutService]   - Portal: '{portal.Title}' type={portal.PortalType} pos=({portal.X},{portal.Y}) size={portal.Width}x{portal.Height}");
                    }
                    return config;
                }
                else
                {
                    Debug.WriteLine("[PortalLayoutService] Deserialize returned null!");
                }
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[PortalLayoutService] Failed to load config: {ex.Message}");
            Debug.WriteLine($"[PortalLayoutService] Stack trace: {ex.StackTrace}");
        }

        Debug.WriteLine("[PortalLayoutService] No config found, returning default (empty)");
        return new DeskGridConfig();
    }

    /// <summary>
    /// Saves the portal configuration to disk
    /// </summary>
    public void SaveConfig(DeskGridConfig config)
    {
        Debug.WriteLine($"[PortalLayoutService] === SAVING CONFIG ===");
        Debug.WriteLine($"[PortalLayoutService] Config path: {_configPath}");
        Debug.WriteLine($"[PortalLayoutService] Portals count: {config.Portals.Count}, portalsVisible: {config.PortalsVisible}");
        
        foreach (var portal in config.Portals)
        {
            Debug.WriteLine($"[PortalLayoutService]   - Saving portal: '{portal.Title}' type={portal.PortalType} pos=({portal.X},{portal.Y})");
        }
        
        try
        {
            var json = JsonSerializer.Serialize(config, _jsonOptions);
            Debug.WriteLine($"[PortalLayoutService] JSON to write ({json.Length} chars): {json.Substring(0, Math.Min(500, json.Length))}...");
            File.WriteAllText(_configPath, json);
            Debug.WriteLine($"[PortalLayoutService] Save SUCCESSFUL");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[PortalLayoutService] Failed to save config: {ex.Message}");
            Debug.WriteLine($"[PortalLayoutService] Stack trace: {ex.StackTrace}");
        }
    }

    /// <summary>
    /// Creates a PortalConfig from a PortalWindow
    /// </summary>
    public PortalConfig CreateConfigFromWindow(Views.PortalWindow window, string id)
    {
        var config = new PortalConfig
        {
            Id = id,
            PortalType = "folder",
            Title = window.Title
        };

        // Get physical pixel position
        var hwnd = new System.Windows.Interop.WindowInteropHelper(window).Handle;
        if (hwnd != IntPtr.Zero && GetWindowRect(hwnd, out RECT rect))
        {
            config.X = rect.Left;
            config.Y = rect.Top;
            config.Width = rect.Right - rect.Left;
            config.Height = rect.Bottom - rect.Top;
        }
        else
        {
            config.X = (int)window.Left;
            config.Y = (int)window.Top;
            config.Width = (int)window.Width;
            config.Height = (int)window.Height;
        }

        if (window.DataContext is ViewModels.PortalViewModel vm)
        {
            config.FolderPath = vm.FolderPath;
            config.Title = vm.Title;
            config.IsRolledUp = vm.IsRolledUp;
            config.SortMode = vm.CurrentSortMode.ToString();
            config.HeaderColor = vm.HeaderColor;
            config.BackgroundColor = vm.BackgroundColor;
            config.TitleAlignment = vm.TitleAlignment;
        }

        return config;
    }

    /// <summary>
    /// Creates a PortalConfig from an AppPortalWindow
    /// </summary>
    public PortalConfig CreateConfigFromWindow(Views.AppPortalWindow window, string id)
    {
        var config = new PortalConfig
        {
            Id = id,
            PortalType = "app",
            Title = window.Title
        };

        // Get physical pixel position
        var hwnd = new System.Windows.Interop.WindowInteropHelper(window).Handle;
        if (hwnd != IntPtr.Zero && GetWindowRect(hwnd, out RECT rect))
        {
            config.X = rect.Left;
            config.Y = rect.Top;
            config.Width = rect.Right - rect.Left;
            config.Height = rect.Bottom - rect.Top;
        }

        if (window.DataContext is ViewModels.AppPortalViewModel vm)
        {
            config.ExecutablePath = vm.App.ExecutablePath;
            config.Title = vm.Title;
            config.IsRolledUp = vm.IsRolledUp;
        }

        return config;
    }

    /// <summary>
    /// Creates a PortalConfig from a UrlPortalWindow
    /// </summary>
    public PortalConfig CreateConfigFromWindow(Views.UrlPortalWindow window, string id)
    {
        var config = new PortalConfig
        {
            Id = id,
            PortalType = "url",
            Title = window.Title
        };

        // Get physical pixel position
        var hwnd = new System.Windows.Interop.WindowInteropHelper(window).Handle;
        if (hwnd != IntPtr.Zero && GetWindowRect(hwnd, out RECT rect))
        {
            config.X = rect.Left;
            config.Y = rect.Top;
            config.Width = rect.Right - rect.Left;
            config.Height = rect.Bottom - rect.Top;
        }

        // Save bookmarks
        config.Bookmarks = window.Bookmarks.Select(b => new BookmarkConfig
        {
            Title = b.Title,
            Url = b.Url
        }).ToList();

        return config;
    }

    #region Win32

    [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left, Top, Right, Bottom;
    }

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    #endregion
}
