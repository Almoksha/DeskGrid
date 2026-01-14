using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Animation;
using DeskGrid.ViewModels;
using DeskGrid.Services;

namespace DeskGrid.Views;

/// <summary>
/// Portal Window - Desktop container with roll-up and blur
/// </summary>
public partial class PortalWindow : Window
{
    private bool _isDragging;
    private System.Windows.Point _dragStartPoint;
    private System.Windows.Point _windowStartPosition;
    private double _expandedHeight;
    private bool _isRolledUp;
    
    // For drag and drop
    private System.Windows.Point _dragStartPosition;
    private bool _isDragInProgress;
    
    // Snapping service
    private readonly SnapService _snapService = new();
    
    // Auto-expand on hover (user configurable)
    private bool _autoExpandOnHover = true;

    // For blur effect
    private const int WM_DWMCOMPOSITIONCHANGED = 0x031E;
    
    // Win32 for focus management (needed for desktop-embedded windows)
    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    private static extern IntPtr SetFocus(IntPtr hWnd);

    public PortalWindow()
    {
        System.Diagnostics.Debug.WriteLine($"[PortalWindow] === CONSTRUCTOR START ===");
        InitializeComponent();
        Loaded += PortalWindow_Loaded;
        MouseEnter += PortalWindow_MouseEnter;
        MouseLeave += PortalWindow_MouseLeave;

        // Use Preview events to handle clicks before they reach the desktop
        ItemsListBox.PreviewMouseLeftButtonDown += ItemsListBox_PreviewMouseLeftButtonDown;
        ItemsListBox.PreviewMouseDoubleClick += ItemsListBox_PreviewMouseDoubleClick;
        ItemsListBox.PreviewMouseRightButtonDown += ItemsListBox_PreviewMouseRightButtonDown;
        ItemsListBox.PreviewMouseMove += ItemsListBox_PreviewMouseMove;
        
        // Fix focus for TextBoxes in desktop-embedded windows
        TitleEditBox.GotFocus += TextBox_GotFocus;
        SearchTextBox.GotFocus += TextBox_GotFocus;
        System.Diagnostics.Debug.WriteLine($"[PortalWindow] === CONSTRUCTOR COMPLETE ===");
    }
    
    /// <summary>
    /// Forces window focus when a TextBox gets focus (fixes keyboard input in desktop-embedded windows)
    /// </summary>
    private void TextBox_GotFocus(object sender, RoutedEventArgs e)
    {
        System.Diagnostics.Debug.WriteLine($"[PortalWindow] TextBox got focus, forcing window focus");
        var hwnd = new WindowInteropHelper(this).Handle;
        if (hwnd != IntPtr.Zero)
        {
            SetForegroundWindow(hwnd);
            SetFocus(hwnd);
        }
    }

    private void PortalWindow_Loaded(object sender, RoutedEventArgs e)
    {
        System.Diagnostics.Debug.WriteLine($"[PortalWindow] === LOADED EVENT ===");
        System.Diagnostics.Debug.WriteLine($"[PortalWindow] Title: '{Title}'");
        if (DataContext is PortalViewModel vm)
        {
            System.Diagnostics.Debug.WriteLine($"[PortalWindow] FolderPath: '{vm.FolderPath}'");
            System.Diagnostics.Debug.WriteLine($"[PortalWindow] HeaderColor: '{vm.HeaderColor}'");
            System.Diagnostics.Debug.WriteLine($"[PortalWindow] BackgroundColor: '{vm.BackgroundColor}'");
            System.Diagnostics.Debug.WriteLine($"[PortalWindow] TitleAlignment: '{vm.TitleAlignment}'");
        }
        _expandedHeight = Height;
        ApplyBlurEffect();
        PlayFadeInAnimation();
        
        // Explicitly register for OLE drops - needed for embedded desktop windows
        // This ensures we receive drops from desktop icons and Explorer
        RegisterOleDropTarget();
    }
    
    /// <summary>
    /// Registers this window as an OLE drop target to receive drops from desktop icons
    /// </summary>
    private void RegisterOleDropTarget()
    {
        try
        {
            var hwnd = new WindowInteropHelper(this).Handle;
            if (hwnd != IntPtr.Zero)
            {
                // Register OLE drop target using Win32
                int result = RegisterDragDrop(hwnd, new OleDropTarget(this));
                System.Diagnostics.Debug.WriteLine($"[PortalWindow] RegisterDragDrop result: {result} (0=success)");
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[PortalWindow] Failed to register OLE drop target: {ex.Message}");
        }
    }
    
    [DllImport("ole32.dll")]
    private static extern int RegisterDragDrop(IntPtr hwnd, IDropTarget pDropTarget);
    
    [DllImport("ole32.dll")]
    private static extern int RevokeDragDrop(IntPtr hwnd);

    /// <summary>
    /// Plays a fade-in animation when the portal is created
    /// </summary>
    private void PlayFadeInAnimation()
    {
        // Start from invisible
        Opacity = 0;
        MainBorder.RenderTransform = new ScaleTransform(0.95, 0.95);
        MainBorder.RenderTransformOrigin = new System.Windows.Point(0.5, 0.5);

        // Fade in opacity
        var fadeIn = new DoubleAnimation
        {
            From = 0,
            To = 1,
            Duration = TimeSpan.FromMilliseconds(200),
            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
        };
        BeginAnimation(OpacityProperty, fadeIn);

        // Scale up slightly
        var scaleXAnim = new DoubleAnimation
        {
            From = 0.95,
            To = 1,
            Duration = TimeSpan.FromMilliseconds(200),
            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
        };
        var scaleYAnim = new DoubleAnimation
        {
            From = 0.95,
            To = 1,
            Duration = TimeSpan.FromMilliseconds(200),
            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
        };
        
        if (MainBorder.RenderTransform is ScaleTransform scale)
        {
            scale.BeginAnimation(ScaleTransform.ScaleXProperty, scaleXAnim);
            scale.BeginAnimation(ScaleTransform.ScaleYProperty, scaleYAnim);
        }
    }

    /// <summary>
    /// Sets the folder this Portal displays
    /// </summary>
    public void SetFolder(string path)
    {
        System.Diagnostics.Debug.WriteLine($"[PortalWindow] SetFolder called with path: '{path}'");
        if (DataContext is PortalViewModel vm)
        {
            vm.SetFolder(path);
            System.Diagnostics.Debug.WriteLine($"[PortalWindow] SetFolder complete, items loaded: {vm.Items.Count}");
        }
        else
        {
            System.Diagnostics.Debug.WriteLine($"[PortalWindow] SetFolder FAILED: DataContext is not PortalViewModel");
        }
    }

    #region Blur Effect

    private void ApplyBlurEffect()
    {
        try
        {
            var hwnd = new WindowInteropHelper(this).Handle;
            if (hwnd == IntPtr.Zero) return;

            // Enable blur behind
            var accent = new AccentPolicy
            {
                AccentState = AccentState.ACCENT_ENABLE_BLURBEHIND,
                AccentFlags = 2,
                GradientColor = unchecked((int)0x99000000) // Semi-transparent black
            };

            var accentSize = Marshal.SizeOf(accent);
            var accentPtr = Marshal.AllocHGlobal(accentSize);
            Marshal.StructureToPtr(accent, accentPtr, false);

            var data = new WindowCompositionAttributeData
            {
                Attribute = WindowCompositionAttribute.WCA_ACCENT_POLICY,
                SizeOfData = accentSize,
                Data = accentPtr
            };

            SetWindowCompositionAttribute(hwnd, ref data);
            Marshal.FreeHGlobal(accentPtr);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Blur effect failed: {ex.Message}");
        }
    }

    [DllImport("user32.dll")]
    private static extern int SetWindowCompositionAttribute(IntPtr hwnd, ref WindowCompositionAttributeData data);

    private enum AccentState
    {
        ACCENT_DISABLED = 0,
        ACCENT_ENABLE_GRADIENT = 1,
        ACCENT_ENABLE_TRANSPARENTGRADIENT = 2,
        ACCENT_ENABLE_BLURBEHIND = 3,
        ACCENT_ENABLE_ACRYLICBLURBEHIND = 4,
        ACCENT_INVALID_STATE = 5
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct AccentPolicy
    {
        public AccentState AccentState;
        public int AccentFlags;
        public int GradientColor;
        public int AnimationId;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct WindowCompositionAttributeData
    {
        public WindowCompositionAttribute Attribute;
        public IntPtr Data;
        public int SizeOfData;
    }

    private enum WindowCompositionAttribute
    {
        WCA_ACCENT_POLICY = 19
    }

    #endregion

    #region Drag and Move

    private void Header_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ClickCount == 2)
        {
            // Double-click to toggle roll-up
            ToggleRollUp();
            return;
        }

        _isDragging = true;
        
        // Get mouse position in screen coordinates
        _dragStartPoint = PointToScreen(e.GetPosition(this));
        
        // Get current window position using Win32 API
        var hwnd = new WindowInteropHelper(this).Handle;
        if (hwnd != IntPtr.Zero && GetWindowRect(hwnd, out RECT rect))
        {
            _windowStartPosition = new System.Windows.Point(rect.Left, rect.Top);
        }
        else
        {
            _windowStartPosition = new System.Windows.Point(Left, Top);
        }
        
        HeaderBorder.CaptureMouse();
    }

    private void Header_MouseMove(object sender, System.Windows.Input.MouseEventArgs e)
    {
        if (!_isDragging) return;

        // Get current mouse position in screen coordinates
        var currentScreenPoint = PointToScreen(e.GetPosition(this));
        
        // Calculate new window position
        var newX = (int)(_windowStartPosition.X + (currentScreenPoint.X - _dragStartPoint.X));
        var newY = (int)(_windowStartPosition.Y + (currentScreenPoint.Y - _dragStartPoint.Y));

        // Get current window size for snapping
        var hwnd = new WindowInteropHelper(this).Handle;
        if (hwnd != IntPtr.Zero && GetWindowRect(hwnd, out RECT rect))
        {
            var width = rect.Right - rect.Left;
            var height = rect.Bottom - rect.Top;
            
            // Get bounds of other portals for snapping
            var otherPortals = GetOtherPortalBounds();
            
            // Apply snapping
            var currentBounds = new System.Drawing.Rectangle(newX, newY, width, height);
            var snappedPos = _snapService.GetSnappedPosition(currentBounds, otherPortals);
            newX = snappedPos.X;
            newY = snappedPos.Y;
            
            SetWindowPos(hwnd, IntPtr.Zero, newX, newY, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
        }
    }

    private void Header_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (_isDragging)
        {
            _isDragging = false;
            HeaderBorder.ReleaseMouseCapture();
            
            // Save layout after drag (position change)
            TriggerSaveLayout();
        }
    }
    
    /// <summary>
    /// Triggers a layout save via the DesktopManager
    /// </summary>
    private void TriggerSaveLayout()
    {
        if (System.Windows.Application.Current is App app)
        {
            app.DesktopManager?.SaveLayout();
        }
    }

    /// <summary>
    /// Gets the bounds of all other portal windows (for snapping)
    /// </summary>
    private IEnumerable<System.Drawing.Rectangle> GetOtherPortalBounds()
    {
        var results = new List<System.Drawing.Rectangle>();
        
        if (System.Windows.Application.Current is not App app || app.DesktopManager == null)
            return results;

        // Get all portal windows from the main window's collection via reflection or public accessor
        // For now, we iterate through all open windows of our types
        foreach (Window window in System.Windows.Application.Current.Windows)
        {
            if (window == this) continue; // Skip self
            if (window is not PortalWindow && window is not AppPortalWindow) continue;
            
            var hwnd = new WindowInteropHelper(window).Handle;
            if (hwnd != IntPtr.Zero && GetWindowRect(hwnd, out RECT rect))
            {
                results.Add(new System.Drawing.Rectangle(
                    rect.Left, rect.Top,
                    rect.Right - rect.Left,
                    rect.Bottom - rect.Top));
            }
        }
        
        return results;
    }

    #endregion

    #region Win32 Declarations for Drag

    private const uint SWP_NOSIZE = 0x0001;
    private const uint SWP_NOZORDER = 0x0004;
    private const uint SWP_NOACTIVATE = 0x0010;

    [DllImport("user32.dll")]
    [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
    private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

    [DllImport("user32.dll")]
    [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left, Top, Right, Bottom;
    }

    #endregion

    #region Roll-up Animation

    private void RollUpButton_Click(object sender, RoutedEventArgs e)
    {
        ToggleRollUp();
    }

    private void ToggleRollUp()
    {
        _isRolledUp = !_isRolledUp;

        if (DataContext is PortalViewModel vm)
        {
            vm.IsRolledUp = _isRolledUp;
        }

        if (_isRolledUp)
        {
            RollUp();
        }
        else
        {
            Expand();
        }
    }

    private void RollUp()
    {
        _expandedHeight = Height;

        // Animate height
        var heightAnim = new DoubleAnimation
        {
            From = Height,
            To = 32, // Just the header
            Duration = TimeSpan.FromMilliseconds(200),
            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
        };
        BeginAnimation(HeightProperty, heightAnim);

        // Rotate arrow
        var rotateAnim = new DoubleAnimation
        {
            To = 180,
            Duration = TimeSpan.FromMilliseconds(200)
        };
        RollUpArrow.BeginAnimation(System.Windows.Media.RotateTransform.AngleProperty, rotateAnim);

        // Hide content
        ContentGrid.Visibility = Visibility.Collapsed;
    }

    private void Expand()
    {
        // Show content first
        ContentGrid.Visibility = Visibility.Visible;

        // Animate height
        var heightAnim = new DoubleAnimation
        {
            From = Height,
            To = _expandedHeight,
            Duration = TimeSpan.FromMilliseconds(200),
            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
        };
        BeginAnimation(HeightProperty, heightAnim);

        // Rotate arrow back
        var rotateAnim = new DoubleAnimation
        {
            To = 0,
            Duration = TimeSpan.FromMilliseconds(200)
        };
        RollUpArrow.BeginAnimation(System.Windows.Media.RotateTransform.AngleProperty, rotateAnim);
    }

    #endregion

    #region Auto Roll-up on Mouse Leave

    private System.Timers.Timer? _autoRollUpTimer;

    private void PortalWindow_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
    {
        _autoRollUpTimer?.Stop();

        // Expand if rolled up and auto-expand is enabled
        if (_isRolledUp && _autoExpandOnHover)
        {
            Expand();
            _isRolledUp = false;
        }
    }

    private void PortalWindow_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
    {
        // Start timer for auto roll-up (if enabled)
        // For now, we don't auto roll-up, user must click the button
    }

    #endregion

    #region Search

    private void SearchButton_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is PortalViewModel vm)
        {
            vm.IsSearchVisible = !vm.IsSearchVisible;
            if (vm.IsSearchVisible)
            {
                SearchTextBox.Focus();
            }
            else
            {
                vm.SearchQuery = string.Empty;
            }
        }
    }

    private void ClearSearchButton_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is PortalViewModel vm)
        {
            vm.SearchQuery = string.Empty;
            vm.IsSearchVisible = false;
        }
    }

    private void SearchTextBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key == Key.Escape)
        {
            if (DataContext is PortalViewModel vm)
            {
                vm.SearchQuery = string.Empty;
                vm.IsSearchVisible = false;
            }
            e.Handled = true;
        }
    }

    #endregion

    #region Title Rename

    private bool _isEditingTitle;
    private string? _originalTitle;

    private void TitleText_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ClickCount == 2) // Double-click to edit
        {
            StartTitleEdit();
            e.Handled = true;
        }
    }

    private void StartTitleEdit()
    {
        System.Diagnostics.Debug.WriteLine($"[PortalWindow] StartTitleEdit called");
        if (DataContext is PortalViewModel vm)
        {
            _originalTitle = vm.Title;
            _isEditingTitle = true;
            TitleText.Visibility = Visibility.Collapsed;
            TitleEditBox.Visibility = Visibility.Visible;
            TitleEditBox.Focus();
            TitleEditBox.SelectAll();
            System.Diagnostics.Debug.WriteLine($"[PortalWindow] Title edit started, current title: '{vm.Title}'");
        }
    }

    private void EndTitleEdit(bool save)
    {
        if (!_isEditingTitle) return;
        _isEditingTitle = false;

        if (DataContext is PortalViewModel vm)
        {
            if (!save && _originalTitle != null)
            {
                System.Diagnostics.Debug.WriteLine($"[PortalWindow] Title edit cancelled, reverting to: '{_originalTitle}'");
                vm.Title = _originalTitle;
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"[PortalWindow] Title edit saved, new title: '{vm.Title}'");
            }
        }

        TitleEditBox.Visibility = Visibility.Collapsed;
        TitleText.Visibility = Visibility.Visible;
        _originalTitle = null;
    }

    private void TitleEditBox_LostFocus(object sender, RoutedEventArgs e)
    {
        System.Diagnostics.Debug.WriteLine($"[PortalWindow] TitleEditBox lost focus, saving title");
        EndTitleEdit(save: true);
        TriggerSaveLayout(); // Save title change
    }

    private void TitleEditBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            System.Diagnostics.Debug.WriteLine($"[PortalWindow] Enter pressed, confirming title edit");
            EndTitleEdit(save: true);
            TriggerSaveLayout();
            e.Handled = true;
        }
        else if (e.Key == Key.Escape)
        {
            System.Diagnostics.Debug.WriteLine($"[PortalWindow] Escape pressed, cancelling title edit");
            EndTitleEdit(save: false);
            e.Handled = true;
        }
    }

    #endregion

    #region Menu

    private void MenuButton_Click(object sender, RoutedEventArgs e)
    {
        var menu = new ContextMenu();

        var sortMenu = new MenuItem { Header = "Sort by" };
        sortMenu.Items.Add(CreateSortMenuItem("Name", SortMode.Name));
        sortMenu.Items.Add(CreateSortMenuItem("Date Modified", SortMode.Date));
        sortMenu.Items.Add(CreateSortMenuItem("Size", SortMode.Size));
        sortMenu.Items.Add(CreateSortMenuItem("Type", SortMode.Type));
        menu.Items.Add(sortMenu);

        menu.Items.Add(new Separator());

        var refreshItem = new MenuItem { Header = "Refresh" };
        refreshItem.Click += (s, args) =>
        {
            if (DataContext is PortalViewModel vm)
                vm.RefreshCommand.Execute(null);
        };
        menu.Items.Add(refreshItem);

        var changeFolderItem = new MenuItem { Header = "Change Folder..." };
        changeFolderItem.Click += ChangeFolderItem_Click;
        menu.Items.Add(changeFolderItem);

        // Header Color submenu
        var colorMenu = new MenuItem { Header = "Header Color" };
        colorMenu.Items.Add(CreateColorMenuItem("Default", null));
        colorMenu.Items.Add(new Separator());
        colorMenu.Items.Add(CreateColorMenuItem("🟣 Purple", "#7C3AED"));
        colorMenu.Items.Add(CreateColorMenuItem("🔵 Blue", "#3B82F6"));
        colorMenu.Items.Add(CreateColorMenuItem("🟢 Green", "#10B981"));
        colorMenu.Items.Add(CreateColorMenuItem("🟡 Yellow", "#F59E0B"));
        colorMenu.Items.Add(CreateColorMenuItem("🔴 Red", "#EF4444"));
        colorMenu.Items.Add(CreateColorMenuItem("🟤 Brown", "#78716C"));
        menu.Items.Add(colorMenu);

        // Background Color submenu
        var bgColorMenu = new MenuItem { Header = "Background Color" };
        bgColorMenu.Items.Add(CreateBgColorMenuItem("Default", null));
        bgColorMenu.Items.Add(new Separator());
        bgColorMenu.Items.Add(CreateBgColorMenuItem("🟣 Purple", "#30553498"));
        bgColorMenu.Items.Add(CreateBgColorMenuItem("🔵 Blue", "#302563EB"));
        bgColorMenu.Items.Add(CreateBgColorMenuItem("🟢 Green", "#30059669"));
        bgColorMenu.Items.Add(CreateBgColorMenuItem("🟡 Yellow", "#30D97706"));
        bgColorMenu.Items.Add(CreateBgColorMenuItem("🔴 Red", "#30DC2626"));
        bgColorMenu.Items.Add(CreateBgColorMenuItem("⬛ Darker", "#E0101010"));
        menu.Items.Add(bgColorMenu);

        // Title Alignment submenu
        var alignMenu = new MenuItem { Header = "Title Alignment" };
        alignMenu.Items.Add(CreateTitleAlignMenuItem("Left", System.Windows.HorizontalAlignment.Left));
        alignMenu.Items.Add(CreateTitleAlignMenuItem("Center", System.Windows.HorizontalAlignment.Center));
        alignMenu.Items.Add(CreateTitleAlignMenuItem("Right", System.Windows.HorizontalAlignment.Right));
        menu.Items.Add(alignMenu);

        menu.Items.Add(new Separator());
        
        // Auto-expand on hover toggle
        var autoExpandItem = new MenuItem 
        { 
            Header = "Auto-expand on Hover",
            IsCheckable = true,
            IsChecked = _autoExpandOnHover
        };
        autoExpandItem.Click += (s, args) =>
        {
            _autoExpandOnHover = autoExpandItem.IsChecked;
            TriggerSaveLayout(); // Save setting
        };
        menu.Items.Add(autoExpandItem);

        menu.Items.Add(new Separator());

        var closeItem = new MenuItem { Header = "Close Portal" };
        closeItem.Click += (s, args) => Close();
        menu.Items.Add(closeItem);

        menu.IsOpen = true;
    }

    private MenuItem CreateSortMenuItem(string header, SortMode mode)
    {
        var item = new MenuItem { Header = header };
        item.Click += (s, e) =>
        {
            if (DataContext is PortalViewModel vm)
                vm.SortCommand.Execute(mode);
        };
        return item;
    }

    private void ChangeFolderItem_Click(object sender, RoutedEventArgs e)
    {
        using var dialog = new System.Windows.Forms.FolderBrowserDialog();
        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            SetFolder(dialog.SelectedPath);
            TriggerSaveLayout(); // Save folder change
        }
    }

    private MenuItem CreateColorMenuItem(string header, string? colorHex)
    {
        var item = new MenuItem { Header = header };
        item.Click += (s, e) =>
        {
            if (DataContext is PortalViewModel vm)
            {
                vm.HeaderColor = colorHex;
                ApplyHeaderColor(colorHex);
                TriggerSaveLayout(); // Save color change
            }
        };
        return item;
    }

    private void ApplyHeaderColor(string? colorHex)
    {
        if (string.IsNullOrEmpty(colorHex))
        {
            // Reset to default theme color
            HeaderBorder.Background = (System.Windows.Media.Brush)FindResource("PortalHeaderBrush");
        }
        else
        {
            try
            {
                var color = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(colorHex);
                HeaderBorder.Background = new System.Windows.Media.SolidColorBrush(color);
            }
            catch
            {
                HeaderBorder.Background = (System.Windows.Media.Brush)FindResource("PortalHeaderBrush");
            }
        }
    }

    private MenuItem CreateBgColorMenuItem(string header, string? colorHex)
    {
        var item = new MenuItem { Header = header };
        item.Click += (s, e) =>
        {
            ApplyBackgroundColor(colorHex);
            TriggerSaveLayout(); // Save bg color change
        };
        return item;
    }

    private void ApplyBackgroundColor(string? colorHex)
    {
        // Store in ViewModel for persistence
        if (DataContext is PortalViewModel vm)
        {
            vm.BackgroundColor = colorHex;
        }
        
        if (string.IsNullOrEmpty(colorHex))
        {
            MainBorder.Background = (System.Windows.Media.Brush)FindResource("PortalBackgroundBrush");
        }
        else
        {
            try
            {
                var color = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(colorHex);
                MainBorder.Background = new System.Windows.Media.SolidColorBrush(color);
            }
            catch
            {
                MainBorder.Background = (System.Windows.Media.Brush)FindResource("PortalBackgroundBrush");
            }
        }
    }

    private MenuItem CreateTitleAlignMenuItem(string header, System.Windows.HorizontalAlignment alignment)
    {
        var item = new MenuItem { Header = header };
        item.Click += (s, e) =>
        {
            TitleText.HorizontalAlignment = alignment;
            if (alignment == System.Windows.HorizontalAlignment.Center)
                TitleText.Margin = new Thickness(0);
            else if (alignment == System.Windows.HorizontalAlignment.Right)
                TitleText.Margin = new Thickness(0, 0, 8, 0);
            else
                TitleText.Margin = new Thickness(8, 0, 0, 0);
            
            // Store in ViewModel for persistence
            if (DataContext is PortalViewModel vm)
            {
                vm.TitleAlignment = alignment switch
                {
                    System.Windows.HorizontalAlignment.Center => "Center",
                    System.Windows.HorizontalAlignment.Right => "Right",
                    _ => "Left"
                };
            }
            TriggerSaveLayout();
        };
        return item;
    }

    #endregion

    #region Item Interaction

    /// <summary>
    /// Find the FileItem at the given mouse position using visual tree hit testing
    /// </summary>
    private Models.FileItem? GetItemAtPosition(System.Windows.Point position)
    {
        var hitTestResult = VisualTreeHelper.HitTest(ItemsListBox, position);
        if (hitTestResult?.VisualHit == null) return null;

        // Walk up the visual tree to find the ListBoxItem
        DependencyObject? current = hitTestResult.VisualHit;
        while (current != null && current is not ListBoxItem)
        {
            current = VisualTreeHelper.GetParent(current);
        }

        if (current is ListBoxItem listBoxItem)
        {
            return listBoxItem.DataContext as Models.FileItem;
        }
        return null;
    }

    private void ItemsListBox_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        _dragStartPosition = e.GetPosition(ItemsListBox);
        
        var position = e.GetPosition(ItemsListBox);
        var item = GetItemAtPosition(position);
        
        if (item != null)
        {
            ItemsListBox.SelectedItem = item;
            System.Diagnostics.Debug.WriteLine($"[PortalWindow] Selected: {item.Name}");
        }
        // Don't mark as handled - allow selection to work normally
    }

    private void ItemsListBox_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        var position = e.GetPosition(ItemsListBox);
        var item = GetItemAtPosition(position);
        
        System.Diagnostics.Debug.WriteLine($"[PortalWindow] PreviewMouseDoubleClick - Item: {item?.Name}");
        
        if (item != null)
        {
            // Double-click on item: open it
            System.Diagnostics.Debug.WriteLine($"[PortalWindow] Opening: {item.FullPath}");
            item.OpenCommand.Execute(null);
            e.Handled = true; // IMPORTANT: Prevent event from going to desktop
        }
        else
        {
            // Double-click on empty area: Quick Launch - open folder in Explorer
            if (DataContext is PortalViewModel vm && !string.IsNullOrEmpty(vm.FolderPath))
            {
                try
                {
                    System.Diagnostics.Process.Start("explorer.exe", vm.FolderPath);
                    e.Handled = true;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[PortalWindow] Quick Launch failed: {ex.Message}");
                }
            }
        }
    }

    private void ItemsListBox_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
    {
        var position = e.GetPosition(ItemsListBox);
        var item = GetItemAtPosition(position);
        
        System.Diagnostics.Debug.WriteLine($"[PortalWindow] PreviewMouseRightButtonDown - Item: {item?.Name}");
        
        if (item != null)
        {
            ItemsListBox.SelectedItem = item;
            
            // Show item context menu
            var contextMenu = new ContextMenu();
            
            var openItem = new MenuItem { Header = "Open" };
            openItem.Click += (s, args) => item.OpenCommand.Execute(null);
            contextMenu.Items.Add(openItem);
            
            var openLocationItem = new MenuItem { Header = "Open File Location" };
            openLocationItem.Click += (s, args) =>
            {
                try
                {
                    System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{item.FullPath}\"");
                }
                catch { }
            };
            contextMenu.Items.Add(openLocationItem);
            
            contextMenu.Items.Add(new Separator());
            
            var deleteItem = new MenuItem { Header = "Delete" };
            deleteItem.Click += (s, args) =>
            {
                try
                {
                    if (item.IsDirectory)
                        System.IO.Directory.Delete(item.FullPath, true);
                    else
                        System.IO.File.Delete(item.FullPath);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Delete failed: {ex.Message}");
                }
            };
            contextMenu.Items.Add(deleteItem);
            
            contextMenu.IsOpen = true;
            e.Handled = true; // IMPORTANT: Prevent desktop context menu
        }
        else
        {
            // Right-click on empty area - show Portal context menu
            e.Handled = true; // Still prevent desktop context menu
        }
    }

    #endregion

    #region Drag and Drop

    private void ContentGrid_DragEnter(object sender, System.Windows.DragEventArgs e)
    {
        if (e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop))
        {
            DropHighlight.Visibility = Visibility.Visible;
            e.Effects = System.Windows.DragDropEffects.Copy;
        }
        else
        {
            e.Effects = System.Windows.DragDropEffects.None;
        }
        e.Handled = true;
    }

    private void ContentGrid_DragLeave(object sender, System.Windows.DragEventArgs e)
    {
        DropHighlight.Visibility = Visibility.Collapsed;
        e.Handled = true;
    }

    private void ContentGrid_DragOver(object sender, System.Windows.DragEventArgs e)
    {
        if (e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop))
        {
            // Use Move if from same volume, Copy otherwise (or if Ctrl held)
            if ((e.KeyStates & System.Windows.DragDropKeyStates.ControlKey) == System.Windows.DragDropKeyStates.ControlKey)
                e.Effects = System.Windows.DragDropEffects.Copy;
            else
                e.Effects = System.Windows.DragDropEffects.Move;
        }
        else
        {
            e.Effects = System.Windows.DragDropEffects.None;
        }
        e.Handled = true;
    }

    private void ContentGrid_Drop(object sender, System.Windows.DragEventArgs e)
    {
        System.Diagnostics.Debug.WriteLine($"[PortalWindow] === DROP EVENT START ===");
        System.Diagnostics.Debug.WriteLine($"[PortalWindow] Portal Title: '{Title}'");
        
        DropHighlight.Visibility = Visibility.Collapsed;

        if (!e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop))
        {
            System.Diagnostics.Debug.WriteLine($"[PortalWindow] Drop rejected: No FileDrop data present");
            System.Diagnostics.Debug.WriteLine($"[PortalWindow] Available formats: {string.Join(", ", e.Data.GetFormats())}");
            return;
        }

        var files = (string[])e.Data.GetData(System.Windows.DataFormats.FileDrop);
        if (files == null || files.Length == 0)
        {
            System.Diagnostics.Debug.WriteLine($"[PortalWindow] Drop rejected: files array is null or empty");
            return;
        }

        System.Diagnostics.Debug.WriteLine($"[PortalWindow] Files being dropped: {files.Length}");
        foreach (var f in files)
        {
            System.Diagnostics.Debug.WriteLine($"[PortalWindow]   - {f}");
        }

        if (DataContext is not PortalViewModel vm)
        {
            System.Diagnostics.Debug.WriteLine($"[PortalWindow] Drop rejected: DataContext is not PortalViewModel");
            return;
        }
            
        System.Diagnostics.Debug.WriteLine($"[PortalWindow] Current FolderPath: '{vm.FolderPath}'");
        
        // If no folder path set, check if user dropped a folder to use as the path
        if (string.IsNullOrEmpty(vm.FolderPath))
        {
            System.Diagnostics.Debug.WriteLine($"[PortalWindow] Portal is EMPTY - checking if folder was dropped");
            
            // Check if any dropped item is a folder
            foreach (var path in files)
            {
                if (System.IO.Directory.Exists(path))
                {
                    System.Diagnostics.Debug.WriteLine($"[PortalWindow] Setting folder path from dropped folder: {path}");
                    SetFolder(path);
                    TriggerSaveLayout();
                    e.Handled = true;
                    System.Diagnostics.Debug.WriteLine($"[PortalWindow] === DROP EVENT COMPLETE (folder set) ===");
                    return;
                }
            }
            
            // No folder was dropped - show user feedback
            System.Diagnostics.Debug.WriteLine($"[PortalWindow] Drop failed: No folder path set and no folder was dropped");
            System.Windows.MessageBox.Show(
                "This portal doesn't have a folder set yet.\n\n" +
                "To set a folder, either:\n" +
                "• Drop a folder onto this portal\n" +
                "• Right-click the portal header and select 'Change Folder'",
                "Portal Empty",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            e.Handled = true;
            return;
        }

        var targetFolder = vm.FolderPath;
        var doCopy = (e.KeyStates & System.Windows.DragDropKeyStates.ControlKey) == System.Windows.DragDropKeyStates.ControlKey;
        System.Diagnostics.Debug.WriteLine($"[PortalWindow] Target folder: {targetFolder}");
        System.Diagnostics.Debug.WriteLine($"[PortalWindow] Operation: {(doCopy ? "COPY" : "MOVE")}");

        // Initialize rule engine for auto-sorting
        var ruleEngine = new Services.RuleEngine();

        int successCount = 0;
        int failCount = 0;
        
        foreach (var sourcePath in files)
        {
            try
            {
                var fileName = System.IO.Path.GetFileName(sourcePath);
                var destPath = System.IO.Path.Combine(targetFolder, fileName);

                System.Diagnostics.Debug.WriteLine($"[PortalWindow] Processing: {sourcePath}");
                System.Diagnostics.Debug.WriteLine($"[PortalWindow]   -> Destination: {destPath}");

                // Avoid dropping onto itself
                if (string.Equals(sourcePath, destPath, StringComparison.OrdinalIgnoreCase))
                {
                    System.Diagnostics.Debug.WriteLine($"[PortalWindow]   Skipped: Source and destination are the same");
                    continue;
                }

                // Handle duplicate names
                destPath = GetUniqueFilePath(destPath);
                System.Diagnostics.Debug.WriteLine($"[PortalWindow]   Final destination: {destPath}");

                if (System.IO.Directory.Exists(sourcePath))
                {
                    System.Diagnostics.Debug.WriteLine($"[PortalWindow]   Type: Directory");
                    if (doCopy)
                        CopyDirectory(sourcePath, destPath);
                    else
                        MoveDirectoryWithFallback(sourcePath, destPath);
                }
                else if (System.IO.File.Exists(sourcePath))
                {
                    System.Diagnostics.Debug.WriteLine($"[PortalWindow]   Type: File");
                    if (doCopy)
                        System.IO.File.Copy(sourcePath, destPath, false);
                    else
                        MoveFileWithFallback(sourcePath, destPath);
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[PortalWindow]   ERROR: Source path does not exist!");
                    failCount++;
                    continue;
                }

                System.Diagnostics.Debug.WriteLine($"[PortalWindow]   SUCCESS: {(doCopy ? "Copied" : "Moved")}");
                successCount++;

                // Apply auto-sort rules after file is in this portal
                ApplyAutoSortRules(ruleEngine, destPath);
            }
            catch (Exception ex)
            {
                failCount++;
                System.Diagnostics.Debug.WriteLine($"[PortalWindow]   FAILED: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[PortalWindow]   Stack: {ex.StackTrace}");
            }
        }

        System.Diagnostics.Debug.WriteLine($"[PortalWindow] === DROP EVENT COMPLETE: {successCount} succeeded, {failCount} failed ===");
        e.Handled = true;
    }

    /// <summary>
    /// Apply auto-sort rules to move file to target portal if a rule matches
    /// </summary>
    private void ApplyAutoSortRules(Services.RuleEngine ruleEngine, string filePath)
    {
        var fileName = System.IO.Path.GetFileName(filePath);
        var matchingRule = ruleEngine.FindMatchingRule(fileName);

        if (matchingRule == null || string.IsNullOrEmpty(matchingRule.TargetPortalId))
            return;

        // Find the target portal's folder
        if (System.Windows.Application.Current is App app && app.DesktopManager != null)
        {
            var targetFolder = app.DesktopManager.GetPortalFolderById(matchingRule.TargetPortalId);
            
            if (!string.IsNullOrEmpty(targetFolder) && System.IO.Directory.Exists(targetFolder))
            {
                try
                {
                    var destPath = System.IO.Path.Combine(targetFolder, fileName);
                    
                    // Don't move if already in target folder
                    if (string.Equals(System.IO.Path.GetDirectoryName(filePath), targetFolder, StringComparison.OrdinalIgnoreCase))
                        return;

                    destPath = GetUniqueFilePath(destPath);
                    System.IO.File.Move(filePath, destPath);
                    System.Diagnostics.Debug.WriteLine($"[AutoSort] Applied rule '{matchingRule.Name}': {filePath} -> {destPath}");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[AutoSort] Failed to apply rule: {ex.Message}");
                }
            }
        }
    }

    private string GetUniqueFilePath(string path)
    {
        if (!System.IO.File.Exists(path) && !System.IO.Directory.Exists(path))
            return path;

        var dir = System.IO.Path.GetDirectoryName(path) ?? "";
        var name = System.IO.Path.GetFileNameWithoutExtension(path);
        var ext = System.IO.Path.GetExtension(path);
        var counter = 1;

        string newPath;
        do
        {
            newPath = System.IO.Path.Combine(dir, $"{name} ({counter}){ext}");
            counter++;
        } while (System.IO.File.Exists(newPath) || System.IO.Directory.Exists(newPath));

        return newPath;
    }

    private void CopyDirectory(string sourceDir, string destDir)
    {
        System.IO.Directory.CreateDirectory(destDir);

        foreach (var file in System.IO.Directory.GetFiles(sourceDir))
        {
            var destFile = System.IO.Path.Combine(destDir, System.IO.Path.GetFileName(file));
            System.IO.File.Copy(file, destFile);
        }

        foreach (var subDir in System.IO.Directory.GetDirectories(sourceDir))
        {
            var destSubDir = System.IO.Path.Combine(destDir, System.IO.Path.GetFileName(subDir));
            CopyDirectory(subDir, destSubDir);
        }
    }

    /// <summary>
    /// Moves a file using Windows Shell API first (handles desktop files properly),
    /// falling back to standard IO, then copy+delete as last resort
    /// </summary>
    private void MoveFileWithFallback(string sourcePath, string destPath)
    {
        // First try: Windows Shell API (same as Explorer, handles desktop files)
        try
        {
            using var shellOps = new Services.ShellFileOperations();
            if (shellOps.MoveFile(sourcePath, destPath))
            {
                System.Diagnostics.Debug.WriteLine($"[PortalWindow] Moved via Shell API: {sourcePath}");
                return;
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[PortalWindow] Shell API move failed: {ex.Message}");
        }

        // Second try: Standard .NET File.Move
        try
        {
            System.IO.File.Move(sourcePath, destPath);
            System.Diagnostics.Debug.WriteLine($"[PortalWindow] Moved via File.Move: {sourcePath}");
            return;
        }
        catch (UnauthorizedAccessException)
        {
            System.Diagnostics.Debug.WriteLine($"[PortalWindow] File.Move failed (UnauthorizedAccess), trying copy+delete");
        }
        catch (System.IO.IOException)
        {
            System.Diagnostics.Debug.WriteLine($"[PortalWindow] File.Move failed (IOException), trying copy+delete");
        }

        // Last resort: Copy then delete
        try
        {
            System.IO.File.Copy(sourcePath, destPath, false);
            System.Diagnostics.Debug.WriteLine($"[PortalWindow] Copied file, now deleting source");
            try
            {
                System.IO.File.Delete(sourcePath);
                System.Diagnostics.Debug.WriteLine($"[PortalWindow] Deleted source file successfully");
            }
            catch (Exception delEx)
            {
                System.Diagnostics.Debug.WriteLine($"[PortalWindow] Delete after copy failed: {delEx.Message}");
                // File is in destination, original couldn't be deleted - acceptable
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[PortalWindow] All move attempts failed: {ex.Message}");
            throw; // Re-throw to let caller handle
        }
    }

    /// <summary>
    /// Moves a directory using Windows Shell API first, falling back to standard IO
    /// </summary>
    private void MoveDirectoryWithFallback(string sourceDir, string destDir)
    {
        // First try: Windows Shell API
        try
        {
            using var shellOps = new Services.ShellFileOperations();
            if (shellOps.MoveFile(sourceDir, destDir)) // Shell API handles both files and folders
            {
                System.Diagnostics.Debug.WriteLine($"[PortalWindow] Moved directory via Shell API: {sourceDir}");
                return;
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[PortalWindow] Shell API directory move failed: {ex.Message}");
        }

        // Second try: Standard .NET Directory.Move
        try
        {
            System.IO.Directory.Move(sourceDir, destDir);
            System.Diagnostics.Debug.WriteLine($"[PortalWindow] Moved via Directory.Move: {sourceDir}");
            return;
        }
        catch (UnauthorizedAccessException)
        {
            System.Diagnostics.Debug.WriteLine($"[PortalWindow] Directory.Move failed (UnauthorizedAccess), trying copy+delete");
        }
        catch (System.IO.IOException)
        {
            System.Diagnostics.Debug.WriteLine($"[PortalWindow] Directory.Move failed (IOException), trying copy+delete");
        }

        // Last resort: Copy then delete
        try
        {
            CopyDirectory(sourceDir, destDir);
            System.Diagnostics.Debug.WriteLine($"[PortalWindow] Copied directory, now deleting source");
            try
            {
                System.IO.Directory.Delete(sourceDir, true);
                System.Diagnostics.Debug.WriteLine($"[PortalWindow] Deleted source directory successfully");
            }
            catch (Exception delEx)
            {
                System.Diagnostics.Debug.WriteLine($"[PortalWindow] Delete after copy failed: {delEx.Message}");
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[PortalWindow] All directory move attempts failed: {ex.Message}");
            throw;
        }
    }

    private void ItemsListBox_PreviewMouseMove(object sender, System.Windows.Input.MouseEventArgs e)
    {
        if (e.LeftButton != MouseButtonState.Pressed || _isDragInProgress) return;

        var position = e.GetPosition(ItemsListBox);
        var diff = position - _dragStartPosition;

        // Check if moved enough to start drag
        if (Math.Abs(diff.X) < SystemParameters.MinimumHorizontalDragDistance &&
            Math.Abs(diff.Y) < SystemParameters.MinimumVerticalDragDistance)
            return;

        var item = ItemsListBox.SelectedItem as Models.FileItem;
        if (item == null || string.IsNullOrEmpty(item.FullPath)) return;

        try
        {
            _isDragInProgress = true;
            
            var dataObject = new System.Windows.DataObject(System.Windows.DataFormats.FileDrop, new[] { item.FullPath });
            System.Windows.DragDrop.DoDragDrop(ItemsListBox, dataObject, System.Windows.DragDropEffects.Move | System.Windows.DragDropEffects.Copy);
            
            _isDragInProgress = false;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[PortalWindow] Drag start failed: {ex.Message}");
            _isDragInProgress = false;
        }
    }

    #endregion
    
    #region OLE Drop Helpers
    
    /// <summary>
    /// Shows or hides the drop highlight (called from OleDropTarget)
    /// </summary>
    public void ShowDropHighlight(bool visible)
    {
        DropHighlight.Visibility = visible ? Visibility.Visible : Visibility.Collapsed;
    }
    
    /// <summary>
    /// Handles files dropped via OLE (from desktop icons)
    /// </summary>
    public void HandleOleDrop(string[] files)
    {
        if (files == null || files.Length == 0) return;
        
        if (DataContext is not PortalViewModel vm)
            return;
            
        // If no folder path set, check if user dropped a folder to use as the path
        if (string.IsNullOrEmpty(vm.FolderPath))
        {
            foreach (var path in files)
            {
                if (System.IO.Directory.Exists(path))
                {
                    System.Diagnostics.Debug.WriteLine($"[PortalWindow] OLE Drop: Setting folder path from dropped folder: {path}");
                    SetFolder(path);
                    TriggerSaveLayout();
                    return;
                }
            }
            
            System.Diagnostics.Debug.WriteLine("[PortalWindow] OLE Drop failed: No folder path set");
            return;
        }
        
        // Process the dropped files (move them to this portal's folder)
        var targetFolder = vm.FolderPath;
        var ruleEngine = new Services.RuleEngine();
        
        foreach (var sourcePath in files)
        {
            try
            {
                var fileName = System.IO.Path.GetFileName(sourcePath);
                var destPath = System.IO.Path.Combine(targetFolder, fileName);
                
                // Avoid dropping onto itself
                if (string.Equals(sourcePath, destPath, StringComparison.OrdinalIgnoreCase))
                    continue;
                
                destPath = GetUniqueFilePath(destPath);
                
                if (System.IO.Directory.Exists(sourcePath))
                {
                    MoveDirectoryWithFallback(sourcePath, destPath);
                    System.Diagnostics.Debug.WriteLine($"[PortalWindow] OLE Drop: Moved folder {sourcePath} -> {destPath}");
                }
                else if (System.IO.File.Exists(sourcePath))
                {
                    MoveFileWithFallback(sourcePath, destPath);
                    System.Diagnostics.Debug.WriteLine($"[PortalWindow] OLE Drop: Moved file {sourcePath} -> {destPath}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PortalWindow] OLE Drop failed for {sourcePath}: {ex.Message}");
            }
        }
        
        // Refresh the view
        vm.RefreshCommand.Execute(null);
    }
    
    #endregion

    protected override void OnClosed(EventArgs e)
    {
        _autoRollUpTimer?.Dispose();
        
        // Revoke OLE drop target
        try
        {
            var hwnd = new WindowInteropHelper(this).Handle;
            if (hwnd != IntPtr.Zero)
            {
                RevokeDragDrop(hwnd);
            }
        }
        catch { }
        
        if (DataContext is PortalViewModel vm)
            vm.Dispose();

        // Notify desktop manager
        if (System.Windows.Application.Current is App app)
        {
            app.DesktopManager?.DetachWindow(this);
        }

        base.OnClosed(e);
    }
}

/// <summary>
/// Converts 0 to Visible, non-zero to Collapsed
/// </summary>
public class ZeroToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is int count)
            return count == 0 ? Visibility.Visible : Visibility.Collapsed;
        return Visibility.Collapsed;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}

/// <summary>
/// COM IDropTarget interface for OLE drag-drop
/// </summary>
[ComImport]
[Guid("00000122-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IDropTarget
{
    int DragEnter([In] System.Runtime.InteropServices.ComTypes.IDataObject pDataObj, [In] int grfKeyState, [In] long pt, [In, Out] ref int pdwEffect);
    int DragOver([In] int grfKeyState, [In] long pt, [In, Out] ref int pdwEffect);
    int DragLeave();
    int Drop([In] System.Runtime.InteropServices.ComTypes.IDataObject pDataObj, [In] int grfKeyState, [In] long pt, [In, Out] ref int pdwEffect);
}

/// <summary>
/// OLE drop target implementation that forwards to the PortalWindow
/// </summary>
public class OleDropTarget : IDropTarget
{
    private readonly PortalWindow _window;
    private string[]? _droppedFiles;
    
    public OleDropTarget(PortalWindow window)
    {
        _window = window;
    }
    
    public int DragEnter(System.Runtime.InteropServices.ComTypes.IDataObject pDataObj, int grfKeyState, long pt, ref int pdwEffect)
    {
        try
        {
            System.Diagnostics.Debug.WriteLine("[OleDropTarget] DragEnter");
            _droppedFiles = GetFilesFromDataObject(pDataObj);
            
            if (_droppedFiles != null && _droppedFiles.Length > 0)
            {
                pdwEffect = 1; // DROPEFFECT_COPY
                _window.Dispatcher.Invoke(() => _window.ShowDropHighlight(true));
                return 0; // S_OK
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[OleDropTarget] DragEnter error: {ex.Message}");
        }
        
        pdwEffect = 0;
        return 0;
    }
    
    public int DragOver(int grfKeyState, long pt, ref int pdwEffect)
    {
        pdwEffect = _droppedFiles != null ? 1 : 0; // DROPEFFECT_COPY or NONE
        return 0;
    }
    
    public int DragLeave()
    {
        System.Diagnostics.Debug.WriteLine("[OleDropTarget] DragLeave");
        _droppedFiles = null;
        _window.Dispatcher.Invoke(() => _window.ShowDropHighlight(false));
        return 0;
    }
    
    public int Drop(System.Runtime.InteropServices.ComTypes.IDataObject pDataObj, int grfKeyState, long pt, ref int pdwEffect)
    {
        try
        {
            System.Diagnostics.Debug.WriteLine("[OleDropTarget] Drop");
            _window.Dispatcher.Invoke(() => _window.ShowDropHighlight(false));
            
            var files = GetFilesFromDataObject(pDataObj);
            if (files != null && files.Length > 0)
            {
                System.Diagnostics.Debug.WriteLine($"[OleDropTarget] Dropping {files.Length} files");
                _window.Dispatcher.Invoke(() => _window.HandleOleDrop(files));
                pdwEffect = 1; // DROPEFFECT_COPY
                return 0;
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[OleDropTarget] Drop error: {ex.Message}");
        }
        
        pdwEffect = 0;
        return 0;
    }
    
    private string[]? GetFilesFromDataObject(System.Runtime.InteropServices.ComTypes.IDataObject dataObj)
    {
        try
        {
            var format = new System.Runtime.InteropServices.ComTypes.FORMATETC
            {
                cfFormat = 15, // CF_HDROP
                dwAspect = System.Runtime.InteropServices.ComTypes.DVASPECT.DVASPECT_CONTENT,
                lindex = -1,
                tymed = System.Runtime.InteropServices.ComTypes.TYMED.TYMED_HGLOBAL
            };
            
            dataObj.GetData(ref format, out var medium);
            
            if (medium.unionmember != IntPtr.Zero)
            {
                try
                {
                    uint count = DragQueryFile(medium.unionmember, 0xFFFFFFFF, null, 0);
                    var files = new string[count];
                    
                    for (uint i = 0; i < count; i++)
                    {
                        var sb = new System.Text.StringBuilder(260);
                        DragQueryFile(medium.unionmember, i, sb, 260);
                        files[i] = sb.ToString();
                    }
                    
                    return files;
                }
                finally
                {
                    ReleaseStgMedium(ref medium);
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[OleDropTarget] GetFilesFromDataObject error: {ex.Message}");
        }
        
        return null;
    }
    
    [DllImport("shell32.dll", CharSet = CharSet.Auto)]
    private static extern uint DragQueryFile(IntPtr hDrop, uint iFile, System.Text.StringBuilder? lpszFile, int cch);
    
    [DllImport("ole32.dll")]
    private static extern void ReleaseStgMedium(ref System.Runtime.InteropServices.ComTypes.STGMEDIUM pmedium);
}
