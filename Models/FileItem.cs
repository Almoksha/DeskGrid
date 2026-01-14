using System.Windows.Media;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace DeskGrid.Models;

/// <summary>
/// Represents a file or folder item displayed in a Portal
/// </summary>
public partial class FileItem : ObservableObject
{
    [ObservableProperty]
    private string _name = string.Empty;

    [ObservableProperty]
    private string _fullPath = string.Empty;

    [ObservableProperty]
    private ImageSource? _icon;

    [ObservableProperty]
    private bool _isDirectory;

    [ObservableProperty]
    private string _fileType = string.Empty;

    [ObservableProperty]
    private DateTime _modifiedDate;

    [ObservableProperty]
    private long _fileSize;

    [ObservableProperty]
    private bool _isSelected;

    /// <summary>
    /// Opens the file or folder with the default application
    /// </summary>
    [RelayCommand]
    private void Open()
    {
        try
        {
            var startInfo = new System.Diagnostics.ProcessStartInfo
            {
                FileName = FullPath,
                UseShellExecute = true
            };
            System.Diagnostics.Process.Start(startInfo);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Failed to open {FullPath}: {ex.Message}");
        }
    }

    /// <summary>
    /// Gets a formatted file size string
    /// </summary>
    public string FormattedSize
    {
        get
        {
            if (IsDirectory) return "";
            
            const long KB = 1024;
            const long MB = KB * 1024;
            const long GB = MB * 1024;

            return FileSize switch
            {
                >= GB => $"{FileSize / (double)GB:F1} GB",
                >= MB => $"{FileSize / (double)MB:F1} MB",
                >= KB => $"{FileSize / (double)KB:F1} KB",
                _ => $"{FileSize} bytes"
            };
        }
    }

    /// <summary>
    /// Gets a formatted modified date string
    /// </summary>
    public string FormattedDate => ModifiedDate.ToString("MMM dd, yyyy");
}
