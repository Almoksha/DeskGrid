using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Animation;
using DeskGrid.Models;
using DeskGrid.ViewModels;

namespace DeskGrid.Views;

/// <summary>
/// App Portal Window - Hosts an embedded application
/// </summary>
public partial class AppPortalWindow : Window
{
    private bool _isDragging;
    private Point _dragStartPoint;
    private Point _windowStartPosition;
    private double _expandedHeight;
    private bool _isRolledUp;
    private HwndSource? _hwndSource;

    public AppPortalWindow()
    {
        InitializeComponent();
        Loaded += AppPortalWindow_Loaded;
        SizeChanged += AppPortalWindow_SizeChanged;
        
        // Use Preview event to catch clicks before they go to desktop
        PreviewMouseLeftButtonDown += AppPortalWindow_PreviewMouseLeftButtonDown;
    }

    public AppPortalWindow(string executablePath, string? name = null) : this()
    {
        DataContext = new AppPortalViewModel(executablePath, name);
    }

    private void AppPortalWindow_Loaded(object sender, RoutedEventArgs e)
    {
        _expandedHeight = Height;
        
        System.Diagnostics.Debug.WriteLine("[AppPortal] Window loaded, attaching mouse events to ContentGrid and MainBorder");
        
        // Attach Preview events to specific elements that might work better
        ContentGrid.PreviewMouseLeftButtonDown += ContentGrid_PreviewMouseLeftButtonDown;
        MainBorder.PreviewMouseLeftButtonDown += ContentGrid_PreviewMouseLeftButtonDown;
        
        // Get the HwndSource for the content area
        _hwndSource = HwndSource.FromVisual(AppHostBorder) as HwndSource;
        
        if (DataContext is AppPortalViewModel vm)
        {
            // Set up the content handle for embedding
            var contentHwnd = GetContentAreaHandle();
            vm.ContentHandle = contentHwnd;
            vm.ContentSize = GetContentSize();

            // Set the desktop handle (SHELLDLL_DefView) for embedding apps as siblings
            vm.DesktopHandle = FindShellDefView();
            System.Diagnostics.Debug.WriteLine($"[AppPortal] DesktopHandle: {vm.DesktopHandle}");

            // Get actual physical pixel position using GetWindowRect (DPI-aware)
            var hwnd = new WindowInteropHelper(this).Handle;
            if (GetWindowRect(hwnd, out RECT rect))
            {
                vm.PortalScreenPosition = new System.Drawing.Point(rect.Left, rect.Top);
                System.Diagnostics.Debug.WriteLine($"[AppPortal] PortalScreenPosition (physical): {rect.Left}, {rect.Top}");
            }
            else
            {
                // Fallback to WPF properties (may not be accurate with DPI scaling)
                vm.PortalScreenPosition = new System.Drawing.Point((int)Left, (int)Top);
            }

            // Subscribe to property changes for UI updates
            vm.PropertyChanged += ViewModel_PropertyChanged;
            
            // Subscribe to Portal resize requests (for auto-expanding to fit app)
            vm.PortalResizeRequested += OnPortalResizeRequested;
        }

        ApplyBlurEffect();
    }

    private void OnPortalResizeRequested(int contentWidth, int contentHeight)
    {
        // Resize the Portal window to accommodate the app's minimum size
        // Add some padding for borders
        int newWidth = contentWidth + 4; // 2px border on each side
        int newHeight = contentHeight + 4;
        
        System.Diagnostics.Debug.WriteLine($"[AppPortal] Resizing Portal to ({newWidth}, {newHeight})");
        
        // Update Portal size using SetWindowPos for physical pixels
        var hwnd = new WindowInteropHelper(this).Handle;
        if (hwnd != IntPtr.Zero)
        {
            // Get current position
            if (GetWindowRect(hwnd, out RECT rect))
            {
                SetWindowPos(hwnd, IntPtr.Zero, rect.Left, rect.Top, newWidth, newHeight, 
                    SWP_NOZORDER | SWP_NOACTIVATE);
                
                // Update ViewModel's Portal position and content size
                if (DataContext is AppPortalViewModel vm)
                {
                    vm.PortalScreenPosition = new System.Drawing.Point(rect.Left, rect.Top);
                    vm.ContentSize = GetContentSize();
                }
            }
        }
    }

    private void AppPortalWindow_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        if (DataContext is AppPortalViewModel vm)
        {
            vm.ContentSize = GetContentSize();
            vm.OnContentResized(vm.ContentSize.Width, vm.ContentSize.Height);
        }
    }

    private IntPtr GetContentAreaHandle()
    {
        // Get the window handle of this WPF window
        var hwnd = new WindowInteropHelper(this).Handle;
        return hwnd;
    }

    private System.Drawing.Size GetContentSize()
    {
        // Calculate content area size (window size minus header)
        var width = (int)(ActualWidth - 2); // Account for border
        var height = (int)(ActualHeight - 32 - 2); // Subtract header and border
        
        // Convert from WPF logical units to physical pixels for Win32 API
        var source = PresentationSource.FromVisual(this);
        if (source?.CompositionTarget != null)
        {
            var dpiX = source.CompositionTarget.TransformToDevice.M11;
            var dpiY = source.CompositionTarget.TransformToDevice.M22;
            width = (int)(width * dpiX);
            height = (int)(height * dpiY);
        }
        
        return new System.Drawing.Size(Math.Max(100, width), Math.Max(50, height));
    }

    private void ViewModel_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (sender is not AppPortalViewModel vm) return;

        Dispatcher.Invoke(() =>
        {
            switch (vm.App.State)
            {
                case AppPortalState.Stopped:
                    ControlPanel.Visibility = Visibility.Visible;
                    FloatingControls.Visibility = Visibility.Collapsed;
                    PoppedOutPanel.Visibility = Visibility.Collapsed;
                    // Restore background when stopped
                    ContentGrid.Background = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromArgb(0xCC, 0x1E, 0x1E, 0x2E));
                    break;

                case AppPortalState.Running:
                    ControlPanel.Visibility = Visibility.Collapsed;
                    FloatingControls.Visibility = Visibility.Visible;
                    PoppedOutPanel.Visibility = Visibility.Collapsed;
                    // Make content area transparent so embedded app shows through
                    ContentGrid.Background = System.Windows.Media.Brushes.Transparent;
                    break;

                case AppPortalState.PoppedOut:
                    ControlPanel.Visibility = Visibility.Collapsed;
                    FloatingControls.Visibility = Visibility.Collapsed;
                    PoppedOutPanel.Visibility = Visibility.Visible;
                    // Restore background when popped out
                    ContentGrid.Background = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromArgb(0xCC, 0x1E, 0x1E, 0x2E));
                    break;
            }
        });
    }

    #region Blur Effect

    private void ApplyBlurEffect()
    {
        try
        {
            var hwnd = new WindowInteropHelper(this).Handle;
            if (hwnd == IntPtr.Zero) return;

            var accent = new AccentPolicy
            {
                AccentState = AccentState.ACCENT_ENABLE_BLURBEHIND,
                AccentFlags = 2,
                GradientColor = unchecked((int)0x99000000)
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
        catch { }
    }

    [DllImport("user32.dll")]
    private static extern int SetWindowCompositionAttribute(IntPtr hwnd, ref WindowCompositionAttributeData data);

    private enum AccentState { ACCENT_ENABLE_BLURBEHIND = 3 }

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

    private enum WindowCompositionAttribute { WCA_ACCENT_POLICY = 19 }

    #endregion

    #region Drag and Move

    private void Header_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ClickCount == 2)
        {
            ToggleRollUp();
            return;
        }

        _isDragging = true;
        _dragStartPoint = PointToScreen(e.GetPosition(this));

        var hwnd = new WindowInteropHelper(this).Handle;
        if (hwnd != IntPtr.Zero && GetWindowRect(hwnd, out RECT rect))
        {
            _windowStartPosition = new Point(rect.Left, rect.Top);
        }
        else
        {
            _windowStartPosition = new Point(Left, Top);
        }

        HeaderBorder.CaptureMouse();
    }

    private void Header_MouseMove(object sender, System.Windows.Input.MouseEventArgs e)
    {
        if (!_isDragging) return;

        var currentScreenPoint = PointToScreen(e.GetPosition(this));
        var newX = (int)(_windowStartPosition.X + (currentScreenPoint.X - _dragStartPoint.X));
        var newY = (int)(_windowStartPosition.Y + (currentScreenPoint.Y - _dragStartPoint.Y));

        var hwnd = new WindowInteropHelper(this).Handle;
        if (hwnd != IntPtr.Zero)
        {
            SetWindowPos(hwnd, IntPtr.Zero, newX, newY, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
            
            // Update embedded app position
            if (DataContext is AppPortalViewModel vm)
            {
                vm.OnPortalMoved(newX, newY);
            }
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
        if (Application.Current is App app)
        {
            app.DesktopManager?.SaveLayout();
        }
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

        if (DataContext is AppPortalViewModel vm)
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

        // Hide the embedded app window so it doesn't render outside Portal bounds
        if (DataContext is AppPortalViewModel vm)
        {
            vm.SetEmbeddedWindowVisibility(false);
        }

        var heightAnim = new DoubleAnimation
        {
            To = 32,
            Duration = TimeSpan.FromMilliseconds(200),
            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
        };
        BeginAnimation(HeightProperty, heightAnim);

        var rotateAnim = new DoubleAnimation
        {
            To = 180,
            Duration = TimeSpan.FromMilliseconds(200),
            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
        };
        RollUpArrow.BeginAnimation(RotateTransform.AngleProperty, rotateAnim);
    }

    private void Expand()
    {
        var heightAnim = new DoubleAnimation
        {
            To = _expandedHeight,
            Duration = TimeSpan.FromMilliseconds(200),
            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
        };
        BeginAnimation(HeightProperty, heightAnim);

        var rotateAnim = new DoubleAnimation
        {
            To = 0,
            Duration = TimeSpan.FromMilliseconds(200),
            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
        };
        RollUpArrow.BeginAnimation(RotateTransform.AngleProperty, rotateAnim);

        // Show the embedded app window after expanding
        if (DataContext is AppPortalViewModel vm)
        {
            vm.SetEmbeddedWindowVisibility(true);
        }
    }

    #endregion

    #region Preview Mouse Handler for Desktop Embedding

    /// <summary>
    /// Find a Button at the given position using visual tree hit testing
    /// </summary>
    private System.Windows.Controls.Button? GetButtonAtPosition(Point position)
    {
        var hitTestResult = VisualTreeHelper.HitTest(this, position);
        if (hitTestResult?.VisualHit == null)
        {
            System.Diagnostics.Debug.WriteLine($"[AppPortal] HitTest returned null");
            return null;
        }

        System.Diagnostics.Debug.WriteLine($"[AppPortal] HitTest found: {hitTestResult.VisualHit.GetType().Name}");

        // Walk up the visual tree to find a Button
        DependencyObject? current = hitTestResult.VisualHit;
        int depth = 0;
        while (current != null && current is not System.Windows.Controls.Button && depth < 20)
        {
            current = VisualTreeHelper.GetParent(current);
            depth++;
            if (current != null)
            {
                System.Diagnostics.Debug.WriteLine($"[AppPortal] Walking up: {current.GetType().Name}");
            }
        }

        if (current is System.Windows.Controls.Button btn)
        {
            System.Diagnostics.Debug.WriteLine($"[AppPortal] Found Button: {btn.Name}");
        }
        else
        {
            System.Diagnostics.Debug.WriteLine($"[AppPortal] No Button found after walking tree");
        }

        return current as System.Windows.Controls.Button;
    }

    private void AppPortalWindow_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        var position = e.GetPosition(this);
        var button = GetButtonAtPosition(position);

        if (button != null)
        {
            System.Diagnostics.Debug.WriteLine($"[AppPortal] PreviewMouseLeftButtonDown hit button: {button.Name}");

            // Execute the appropriate action based on which button was hit
            if (DataContext is AppPortalViewModel vm)
            {
                switch (button.Name)
                {
                    case "StartButton":
                        System.Diagnostics.Debug.WriteLine("[AppPortal] Executing Start");
                        vm.StartCommand.Execute(null);
                        e.Handled = true;
                        break;

                    case "StopButton":
                        System.Diagnostics.Debug.WriteLine("[AppPortal] Executing Stop");
                        vm.StopCommand.Execute(null);
                        e.Handled = true;
                        break;

                    case "PopOutButton":
                        System.Diagnostics.Debug.WriteLine("[AppPortal] Executing PopOut");
                        vm.PopOutCommand.Execute(null);
                        e.Handled = true;
                        break;

                    case "PopInButton":
                        System.Diagnostics.Debug.WriteLine("[AppPortal] Executing PopIn");
                        vm.PopInCommand.Execute(null);
                        e.Handled = true;
                        break;

                    case "RollUpButton":
                        System.Diagnostics.Debug.WriteLine("[AppPortal] Executing RollUp");
                        ToggleRollUp();
                        e.Handled = true;
                        break;

                    case "CloseButton":
                        System.Diagnostics.Debug.WriteLine("[AppPortal] Executing Close");
                        CloseButton_Click(sender, e);
                        e.Handled = true;
                        break;
                }
            }
        }
    }

    private void ContentGrid_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        var position = e.GetPosition(this);
        System.Diagnostics.Debug.WriteLine($"[AppPortal] ContentGrid_PreviewMouseLeftButtonDown at {position}");
        
        if (DataContext is not AppPortalViewModel vm) return;

        // Check each button's bounds directly instead of hit-testing
        if (IsClickOnButton(StartButton, position))
        {
            System.Diagnostics.Debug.WriteLine("[AppPortal] Click is on StartButton - executing Start");
            vm.StartCommand.Execute(null);
            e.Handled = true;
            return;
        }

        if (IsClickOnButton(StopButton, position))
        {
            System.Diagnostics.Debug.WriteLine("[AppPortal] Click is on StopButton - executing Stop");
            vm.StopCommand.Execute(null);
            e.Handled = true;
            return;
        }

        if (IsClickOnButton(PopOutButton, position))
        {
            System.Diagnostics.Debug.WriteLine("[AppPortal] Click is on PopOutButton - executing PopOut");
            vm.PopOutCommand.Execute(null);
            e.Handled = true;
            return;
        }

        if (IsClickOnButton(PopInButton, position))
        {
            System.Diagnostics.Debug.WriteLine("[AppPortal] Click is on PopInButton - executing PopIn");
            vm.PopInCommand.Execute(null);
            e.Handled = true;
            return;
        }

        if (IsClickOnButton(RollUpButton, position))
        {
            System.Diagnostics.Debug.WriteLine("[AppPortal] Click is on RollUpButton - executing RollUp");
            ToggleRollUp();
            e.Handled = true;
            return;
        }

        if (IsClickOnButton(CloseButton, position))
        {
            System.Diagnostics.Debug.WriteLine("[AppPortal] Click is on CloseButton - executing Close");
            CloseButton_Click(sender, e);
            e.Handled = true;
            return;
        }
    }

    private bool IsClickOnButton(System.Windows.Controls.Button? button, Point clickPosition)
    {
        if (button == null || !button.IsVisible || button.ActualWidth <= 0 || button.ActualHeight <= 0)
            return false;

        try
        {
            // Transform button bounds to window coordinates
            var transform = button.TransformToAncestor(this);
            var buttonBounds = transform.TransformBounds(new Rect(0, 0, button.ActualWidth, button.ActualHeight));
            
            System.Diagnostics.Debug.WriteLine($"[AppPortal] Button '{button.Name}' bounds: {buttonBounds}, click: {clickPosition}");
            
            return buttonBounds.Contains(clickPosition);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[AppPortal] Error checking button bounds: {ex.Message}");
            return false;
        }
    }

    #endregion

    #region Action Button Handlers

    private void StartButton_Click(object sender, RoutedEventArgs e)
    {
        System.Diagnostics.Debug.WriteLine("[AppPortal] Start button clicked");
        if (DataContext is AppPortalViewModel vm)
        {
            vm.StartCommand.Execute(null);
        }
        e.Handled = true;
    }

    private void StopButton_Click(object sender, RoutedEventArgs e)
    {
        System.Diagnostics.Debug.WriteLine("[AppPortal] Stop button clicked");
        if (DataContext is AppPortalViewModel vm)
        {
            vm.StopCommand.Execute(null);
        }
        e.Handled = true;
    }

    private void PopOutButton_Click(object sender, RoutedEventArgs e)
    {
        System.Diagnostics.Debug.WriteLine("[AppPortal] PopOut button clicked");
        if (DataContext is AppPortalViewModel vm)
        {
            vm.PopOutCommand.Execute(null);
        }
        e.Handled = true;
    }

    private void PopInButton_Click(object sender, RoutedEventArgs e)
    {
        System.Diagnostics.Debug.WriteLine("[AppPortal] PopIn button clicked");
        if (DataContext is AppPortalViewModel vm)
        {
            vm.PopInCommand.Execute(null);
        }
        e.Handled = true;
    }

    #endregion

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is AppPortalViewModel vm)
        {
            if (vm.App.State == AppPortalState.Running)
            {
                var result = MessageBox.Show(
                    "The embedded app is still running. Stop it and close the Portal?",
                    "Confirm Close",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                {
                    vm.StopCommand.Execute(null);
                }
                else
                {
                    return;
                }
            }
        }

        Close();
    }

    protected override void OnClosed(EventArgs e)
    {
        if (DataContext is AppPortalViewModel vm)
        {
            vm.PropertyChanged -= ViewModel_PropertyChanged;
            vm.Dispose();
        }

        if (Application.Current is App app)
        {
            app.DesktopManager?.DetachWindow(this);
        }

        base.OnClosed(e);
    }

    #region Win32 Declarations

    private const uint SWP_NOSIZE = 0x0001;
    private const uint SWP_NOZORDER = 0x0004;
    private const uint SWP_NOACTIVATE = 0x0010;

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left, Top, Right, Bottom;
    }

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr FindWindow(string? lpClassName, string? lpWindowName);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string? lpszWindow);

    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    /// <summary>
    /// Finds the SHELLDLL_DefView window (desktop icons container)
    /// </summary>
    private IntPtr FindShellDefView()
    {
        IntPtr result = IntPtr.Zero;
        
        // First try to find in Progman
        var progman = FindWindow("Progman", null);
        if (progman != IntPtr.Zero)
        {
            result = FindWindowEx(progman, IntPtr.Zero, "SHELLDLL_DefView", null);
            if (result != IntPtr.Zero)
                return result;
        }
        
        // If not in Progman, search in WorkerW windows
        EnumWindows((hwnd, lParam) =>
        {
            var shellView = FindWindowEx(hwnd, IntPtr.Zero, "SHELLDLL_DefView", null);
            if (shellView != IntPtr.Zero)
            {
                result = shellView;
                return false; // Stop enumeration
            }
            return true; // Continue enumeration
        }, IntPtr.Zero);
        
        return result;
    }

    #endregion
}

/// <summary>
/// Converts bool to Visibility
/// </summary>
public class BoolToVisibilityConverter : System.Windows.Data.IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
    {
        return value is bool b && b ? Visibility.Visible : Visibility.Collapsed;
    }

    public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
    {
        return value is Visibility v && v == Visibility.Visible;
    }
}
