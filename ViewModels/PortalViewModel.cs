using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using DeskGrid.Models;
using DeskGrid.Services;

namespace DeskGrid.ViewModels;

/// <summary>
/// ViewModel for a Portal container
/// </summary>
public partial class PortalViewModel : ObservableObject, IDisposable
{
    private FileSystemWatcher? _watcher;
    private bool _disposed;
    private readonly SynchronizationContext? _syncContext;

    [ObservableProperty]
    private string _title = "New Portal";

    [ObservableProperty]
    private string _folderPath = string.Empty;

    [ObservableProperty]
    private bool _isRolledUp;

    [ObservableProperty]
    private bool _isLoading;

    [ObservableProperty]
    private SortMode _currentSortMode = SortMode.Name;

    [ObservableProperty]
    private string _searchQuery = string.Empty;

    [ObservableProperty]
    private bool _isSearchVisible;

    [ObservableProperty]
    private string? _headerColor; // Hex color like "#7C3AED"

    [ObservableProperty]
    private string? _backgroundColor; // Hex color like "#30553498"

    [ObservableProperty]
    private string _titleAlignment = "Left"; // "Left", "Center", or "Right"

    /// <summary>
    /// All items in the folder (unfiltered)
    /// </summary>
    public ObservableCollection<FileItem> Items { get; } = new();

    /// <summary>
    /// Items filtered by search query
    /// </summary>
    public IEnumerable<FileItem> FilteredItems => string.IsNullOrWhiteSpace(SearchQuery)
        ? Items
        : Items.Where(i => i.Name.Contains(SearchQuery, StringComparison.OrdinalIgnoreCase));

    partial void OnSearchQueryChanged(string value)
    {
        OnPropertyChanged(nameof(FilteredItems));
    }

    public PortalViewModel()
    {
        _syncContext = SynchronizationContext.Current;
    }

    /// <summary>
    /// Sets the folder path and loads items
    /// </summary>
    public void SetFolder(string path)
    {
        if (string.IsNullOrEmpty(path) || !Directory.Exists(path))
            return;

        // Stop existing watcher
        StopWatcher();

        FolderPath = path;

        // Load items
        LoadItems();

        // Start watching for changes
        StartWatcher();
    }

    private void LoadItems()
    {
        IsLoading = true;

        try
        {
            Items.Clear();

            if (!Directory.Exists(FolderPath))
                return;

            // Load directories first
            foreach (var dir in Directory.GetDirectories(FolderPath))
            {
                try
                {
                    var dirInfo = new DirectoryInfo(dir);
                    if ((dirInfo.Attributes & FileAttributes.Hidden) != 0)
                        continue;

                    Items.Add(new FileItem
                    {
                        Name = dirInfo.Name,
                        FullPath = dir,
                        IsDirectory = true,
                        ModifiedDate = dirInfo.LastWriteTime,
                        Icon = IconLoader.GetIcon(dir)
                    });
                }
                catch { /* Skip inaccessible directories */ }
            }

            // Load files
            foreach (var file in Directory.GetFiles(FolderPath))
            {
                try
                {
                    var fileInfo = new FileInfo(file);
                    if ((fileInfo.Attributes & FileAttributes.Hidden) != 0)
                        continue;

                    Items.Add(new FileItem
                    {
                        Name = fileInfo.Name,
                        FullPath = file,
                        IsDirectory = false,
                        ModifiedDate = fileInfo.LastWriteTime,
                        FileSize = fileInfo.Length,
                        FileType = fileInfo.Extension,
                        Icon = IconLoader.GetIcon(file)
                    });
                }
                catch { /* Skip inaccessible files */ }
            }

            ApplySort();
        }
        finally
        {
            IsLoading = false;
        }
    }

    private void StartWatcher()
    {
        if (string.IsNullOrEmpty(FolderPath) || !Directory.Exists(FolderPath))
            return;

        _watcher = new FileSystemWatcher(FolderPath)
        {
            NotifyFilter = NotifyFilters.FileName | NotifyFilters.DirectoryName | 
                          NotifyFilters.LastWrite | NotifyFilters.CreationTime,
            EnableRaisingEvents = true
        };

        _watcher.Created += OnFileCreated;
        _watcher.Deleted += OnFileDeleted;
        _watcher.Renamed += OnFileRenamed;
        _watcher.Changed += OnFileChanged;
    }

    private void StopWatcher()
    {
        if (_watcher != null)
        {
            _watcher.EnableRaisingEvents = false;
            _watcher.Created -= OnFileCreated;
            _watcher.Deleted -= OnFileDeleted;
            _watcher.Renamed -= OnFileRenamed;
            _watcher.Changed -= OnFileChanged;
            _watcher.Dispose();
            _watcher = null;
        }
    }

    private void OnFileCreated(object sender, FileSystemEventArgs e)
    {
        RunOnUI(() =>
        {
            try
            {
                var isDir = Directory.Exists(e.FullPath);
                var item = new FileItem
                {
                    Name = Path.GetFileName(e.FullPath),
                    FullPath = e.FullPath,
                    IsDirectory = isDir
                };

                if (isDir)
                {
                    var dirInfo = new DirectoryInfo(e.FullPath);
                    item.ModifiedDate = dirInfo.LastWriteTime;
                }
                else
                {
                    var fileInfo = new FileInfo(e.FullPath);
                    item.ModifiedDate = fileInfo.LastWriteTime;
                    item.FileSize = fileInfo.Length;
                    item.FileType = fileInfo.Extension;
                }

                item.Icon = IconLoader.GetIcon(e.FullPath);
                Items.Add(item);
                ApplySort();
            }
            catch { }
        });
    }

    private void OnFileDeleted(object sender, FileSystemEventArgs e)
    {
        RunOnUI(() =>
        {
            var item = Items.FirstOrDefault(i => i.FullPath == e.FullPath);
            if (item != null)
                Items.Remove(item);
        });
    }

    private void OnFileRenamed(object sender, RenamedEventArgs e)
    {
        RunOnUI(() =>
        {
            var item = Items.FirstOrDefault(i => i.FullPath == e.OldFullPath);
            if (item != null)
            {
                item.Name = Path.GetFileName(e.FullPath);
                item.FullPath = e.FullPath;
                item.Icon = IconLoader.GetIcon(e.FullPath);
                ApplySort();
            }
        });
    }

    private void OnFileChanged(object sender, FileSystemEventArgs e)
    {
        RunOnUI(() =>
        {
            var item = Items.FirstOrDefault(i => i.FullPath == e.FullPath);
            if (item != null)
            {
                try
                {
                    if (File.Exists(e.FullPath))
                    {
                        var fileInfo = new FileInfo(e.FullPath);
                        item.ModifiedDate = fileInfo.LastWriteTime;
                        item.FileSize = fileInfo.Length;
                    }
                }
                catch { }
            }
        });
    }

    private void RunOnUI(Action action)
    {
        if (_syncContext != null)
            _syncContext.Post(_ => action(), null);
        else
            action();
    }

    [RelayCommand]
    private void Sort(SortMode mode)
    {
        CurrentSortMode = mode;
        ApplySort();
    }

    private void ApplySort()
    {
        var sorted = CurrentSortMode switch
        {
            SortMode.Name => Items.OrderBy(i => !i.IsDirectory).ThenBy(i => i.Name),
            SortMode.Date => Items.OrderBy(i => !i.IsDirectory).ThenByDescending(i => i.ModifiedDate),
            SortMode.Size => Items.OrderBy(i => !i.IsDirectory).ThenByDescending(i => i.FileSize),
            SortMode.Type => Items.OrderBy(i => !i.IsDirectory).ThenBy(i => i.FileType).ThenBy(i => i.Name),
            _ => Items.OrderBy(i => i.Name)
        };

        var sortedList = sorted.ToList();
        Items.Clear();
        foreach (var item in sortedList)
            Items.Add(item);
    }

    [RelayCommand]
    private void Refresh()
    {
        LoadItems();
    }

    [RelayCommand]
    private void ToggleRollUp()
    {
        IsRolledUp = !IsRolledUp;
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        StopWatcher();
        GC.SuppressFinalize(this);
    }
}

public enum SortMode
{
    Name,
    Date,
    Size,
    Type
}
