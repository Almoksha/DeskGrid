using System.Text.Json.Serialization;

namespace DeskGrid.Models;

/// <summary>
/// Configuration data for a single portal (for persistence)
/// </summary>
public class PortalConfig
{
    /// <summary>
    /// Unique identifier for this portal
    /// </summary>
    public string Id { get; set; } = Guid.NewGuid().ToString();

    /// <summary>
    /// Type of portal: "folder", "app", "widget", "url"
    /// </summary>
    public string PortalType { get; set; } = "folder";

    /// <summary>
    /// Display title of the portal
    /// </summary>
    public string Title { get; set; } = "New Portal";

    /// <summary>
    /// Folder path for folder portals
    /// </summary>
    public string? FolderPath { get; set; }

    /// <summary>
    /// Executable path for app portals
    /// </summary>
    public string? ExecutablePath { get; set; }

    /// <summary>
    /// Screen X position (physical pixels)
    /// </summary>
    public int X { get; set; }

    /// <summary>
    /// Screen Y position (physical pixels)
    /// </summary>
    public int Y { get; set; }

    /// <summary>
    /// Width in physical pixels
    /// </summary>
    public int Width { get; set; } = 300;

    /// <summary>
    /// Height in physical pixels
    /// </summary>
    public int Height { get; set; } = 400;

    /// <summary>
    /// Whether the portal is rolled up (minimized to header only)
    /// </summary>
    public bool IsRolledUp { get; set; }

    /// <summary>
    /// Custom header color (hex format, e.g. "#7C3AED")
    /// </summary>
    public string? HeaderColor { get; set; }

    /// <summary>
    /// Custom background color (hex format)
    /// </summary>
    public string? BackgroundColor { get; set; }

    /// <summary>
    /// Sort mode for folder portals
    /// </summary>
    public string SortMode { get; set; } = "Name";

    /// <summary>
    /// Bookmarks for URL portals
    /// </summary>
    public List<BookmarkConfig>? Bookmarks { get; set; }

    /// <summary>
    /// Title alignment for folder portals ("Left", "Center", "Right")
    /// </summary>
    public string? TitleAlignment { get; set; }
}

/// <summary>
/// Simplified bookmark config for serialization (doesn't use ObservableObject)
/// </summary>
public class BookmarkConfig
{
    public string Title { get; set; } = string.Empty;
    public string Url { get; set; } = string.Empty;
}

/// <summary>
/// Root configuration containing all portal layouts
/// </summary>
public class DeskGridConfig
{
    /// <summary>
    /// Configuration version for migration support
    /// </summary>
    public int Version { get; set; } = 1;

    /// <summary>
    /// All portal configurations
    /// </summary>
    public List<PortalConfig> Portals { get; set; } = new();

    /// <summary>
    /// Whether portals are currently visible
    /// </summary>
    public bool PortalsVisible { get; set; } = true;

    /// <summary>
    /// Active profile name (for profile system)
    /// </summary>
    public string ActiveProfile { get; set; } = "Default";
}
