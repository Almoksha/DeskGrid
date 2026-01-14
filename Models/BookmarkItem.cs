using CommunityToolkit.Mvvm.ComponentModel;

namespace DeskGrid.Models;

/// <summary>
/// Represents a web bookmark for URL Portals
/// </summary>
public partial class BookmarkItem : ObservableObject
{
    [ObservableProperty]
    private string _title = string.Empty;

    [ObservableProperty]
    private string _url = string.Empty;

    [ObservableProperty]
    private string? _iconPath;

    [ObservableProperty]
    private string? _faviconUrl;

    /// <summary>
    /// Opens the URL in the default browser
    /// </summary>
    public void Open()
    {
        try
        {
            if (!string.IsNullOrEmpty(Url))
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = Url,
                    UseShellExecute = true
                });
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[BookmarkItem] Open failed: {ex.Message}");
        }
    }
}

/// <summary>
/// Configuration for a URL Portal
/// </summary>
public class UrlPortalConfig
{
    public string Title { get; set; } = "Bookmarks";
    public List<BookmarkItem> Bookmarks { get; set; } = new();
}
