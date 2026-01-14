using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Input;
using DeskGrid.Models;

namespace DeskGrid.Views;

public partial class UrlPortalWindow : Window
{
    private bool _isDragging;
    private Point _dragStartPoint;
    private Point _windowStartPosition;

    public ObservableCollection<BookmarkItem> Bookmarks { get; } = new();
    public string Title { get; set; } = "Bookmarks";
    public bool HasNoBookmarks => Bookmarks.Count == 0;

    public UrlPortalWindow()
    {
        InitializeComponent();
        DataContext = this;
        
        Bookmarks.CollectionChanged += (s, e) => 
        {
            OnPropertyChanged(nameof(HasNoBookmarks));
        };
    }

    private void OnPropertyChanged(string propertyName)
    {
        // Simple property change notification
    }

    #region Drag Move

    private void Header_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        _isDragging = true;
        _dragStartPoint = PointToScreen(e.GetPosition(this));
        _windowStartPosition = new Point(Left, Top);
        HeaderBorder.CaptureMouse();
    }

    private void Header_MouseMove(object sender, System.Windows.Input.MouseEventArgs e)
    {
        if (!_isDragging) return;
        var currentScreenPoint = PointToScreen(e.GetPosition(this));
        Left = _windowStartPosition.X + (currentScreenPoint.X - _dragStartPoint.X);
        Top = _windowStartPosition.Y + (currentScreenPoint.Y - _dragStartPoint.Y);
    }

    private void Header_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (_isDragging)
        {
            _isDragging = false;
            HeaderBorder.ReleaseMouseCapture();
        }
    }

    #endregion

    #region Bookmark Actions

    private void AddBookmark_Click(object sender, RoutedEventArgs e)
    {
        var url = Microsoft.VisualBasic.Interaction.InputBox(
            "Enter the URL:", "Add Bookmark", "https://");
        
        if (!string.IsNullOrWhiteSpace(url) && url != "https://")
        {
            var title = Microsoft.VisualBasic.Interaction.InputBox(
                "Enter a title:", "Bookmark Title", GetDomainFromUrl(url));
            
            if (!string.IsNullOrWhiteSpace(title))
            {
                Bookmarks.Add(new BookmarkItem { Title = title, Url = url });
            }
        }
    }

    private void Bookmark_Click(object sender, MouseButtonEventArgs e)
    {
        if (sender is FrameworkElement element && element.DataContext is BookmarkItem bookmark)
        {
            bookmark.Open();
        }
    }

    private void RemoveBookmark_Click(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement element && element.Tag is BookmarkItem bookmark)
        {
            Bookmarks.Remove(bookmark);
        }
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private string GetDomainFromUrl(string url)
    {
        try
        {
            var uri = new Uri(url);
            return uri.Host;
        }
        catch
        {
            return "Bookmark";
        }
    }

    #endregion
}
