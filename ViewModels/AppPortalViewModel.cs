using System.Diagnostics;
using System.Runtime.InteropServices;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using DeskGrid.Models;

namespace DeskGrid.ViewModels;

/// <summary>
/// ViewModel for AppPortal - manages an embedded application
/// </summary>
public partial class AppPortalViewModel : ObservableObject, IDisposable
{
    [ObservableProperty]
    private EmbeddedApp _app;

    [ObservableProperty]
    private string _title = "App portal";

    [ObservableProperty]
    private bool _isRolledUp;

    [ObservableProperty]
    private string _statusText = "Click Start to launch";

    /// <summary>
    /// Current clipping region handle (must be deleted when replaced)
    /// </summary>
    private IntPtr _currentRegion = IntPtr.Zero;

    /// <summary>
    /// Handle to the content area where app will be embedded
    /// </summary>
    public IntPtr ContentHandle { get; set; }

    /// <summary>
    /// Handle to the desktop (SHELLDLL_DefView) for parenting embedded app
    /// </summary>
    public IntPtr DesktopHandle { get; set; }

    /// <summary>
    /// Current screen position of the portal window
    /// </summary>
    public System.Drawing.Point PortalScreenPosition { get; set; }

    /// <summary>
    /// Size of the content area for embedding
    /// </summary>
    public System.Drawing.Size ContentSize { get; set; }

    /// <summary>
    /// Event raised when the portal needs to be resized to fit the embedded app
    /// </summary>
    public event Action<int, int>? PortalResizeRequested;

    /// <summary>
    /// Header height offset for positioning embedded windows
    /// </summary>
    private const int HeaderHeight = 32;

    public AppPortalViewModel(string executablePath, string? name = null)
    {
        _app = new EmbeddedApp
        {
            ExecutablePath = executablePath,
            Name = name ?? System.IO.Path.GetFileNameWithoutExtension(executablePath)
        };
        _title = _app.DisplayName;
        _app.AppExited += OnAppExited;
    }

    private void OnAppExited(object? sender, EventArgs e)
    {
        // Use BeginInvoke to avoid deadlock during shutdown
        System.Windows.Application.Current?.Dispatcher.BeginInvoke(() =>
        {
            StatusText = "App closed. Click Start to relaunch.";
            OnPropertyChanged(nameof(CanStart));
            OnPropertyChanged(nameof(CanStop));
            OnPropertyChanged(nameof(CanPopOut));
        });
    }

    public bool CanStart => App.State == AppPortalState.Stopped;
    public bool CanStop => App.State == AppPortalState.Running;
    public bool CanPopOut => App.State == AppPortalState.Running;
    public bool CanPopIn => App.State == AppPortalState.PoppedOut;

    [RelayCommand]
    private async Task StartAsync()
    {
        if (!CanStart) return;
        
        try
        {
            StatusText = "Starting...";

            // Start the process
            var startInfo = new ProcessStartInfo
            {
                FileName = App.ExecutablePath,
                Arguments = App.Arguments ?? "",
                UseShellExecute = true
            };

            var process = Process.Start(startInfo);
            if (process == null)
            {
                StatusText = "Failed to start process";
                return;
            }

            App.Process = process;
            App.ProcessId = process.Id;

            // Wait for the window to appear with improved detection
            StatusText = "Waiting for window...";
            IntPtr hwnd = IntPtr.Zero;
            int attempts = 0;
            const int maxAttempts = 100; // 10 seconds max for slow-starting apps

            while (attempts < maxAttempts)
            {
                await Task.Delay(100);
                
                // Check if process has exited
                process.Refresh();
                if (process.HasExited)
                {
                    StatusText = "Process exited unexpectedly";
                    return;
                }
                
                hwnd = process.MainWindowHandle;
                
                // Validate the window is actually visible and has a title
                // This helps avoid catching splash screens or hidden windows
                if (hwnd != IntPtr.Zero && IsWindowVisible(hwnd))
                {
                    // Additional check: ensure window has valid dimensions
                    if (GetWindowRect(hwnd, out RECT checkRect))
                    {
                        int checkWidth = checkRect.Right - checkRect.Left;
                        int checkHeight = checkRect.Bottom - checkRect.Top;
                        if (checkWidth > 50 && checkHeight > 50)
                        {
                            break; // Found a valid window
                        }
                    }
                }
                
                hwnd = IntPtr.Zero; // Reset and try again
                attempts++;
            }

            if (hwnd == IntPtr.Zero)
            {
                StatusText = "Could not find app window (timeout)";
                return;
            }

            App.WindowHandle = hwnd;

            // Store original styles for restore on pop-out
            App.OriginalStyle = GetWindowLong(hwnd, GWL_STYLE);
            App.OriginalExStyle = GetWindowLong(hwnd, GWL_EXSTYLE);

            // Get original bounds using RECT
            if (GetWindowRect(hwnd, out RECT rect))
            {
                App.OriginalBounds = new System.Drawing.Rectangle(
                    rect.Left, rect.Top, 
                    rect.Right - rect.Left, 
                    rect.Bottom - rect.Top);
            }

            // Embed the window
            EmbedWindow(hwnd);

            // Monitor for process exit
            process.EnableRaisingEvents = true;
            process.Exited += (s, e) => App.OnAppExited();

            App.State = AppPortalState.Running;
            StatusText = "Running";

            OnPropertyChanged(nameof(CanStart));
            OnPropertyChanged(nameof(CanStop));
            OnPropertyChanged(nameof(CanPopOut));
        }
        catch (Exception ex)
        {
            StatusText = $"Error: {ex.Message}";
            Debug.WriteLine($"[AppPortal] Start failed: {ex}");
        }
    }

    private void EmbedWindow(IntPtr hwnd)
    {
        if (ContentHandle == IntPtr.Zero)
        {
            Debug.WriteLine("[AppPortal] ContentHandle not set, cannot embed");
            return;
        }

        // Remove title bar and borders
        int style = GetWindowLong(hwnd, GWL_STYLE);
        style &= ~(WS_CAPTION | WS_THICKFRAME | WS_MINIMIZEBOX | WS_MAXIMIZEBOX | WS_SYSMENU);
        style |= WS_CHILD;
        SetWindowLong(hwnd, GWL_STYLE, style);

        // Hide from taskbar
        int exStyle = GetWindowLong(hwnd, GWL_EXSTYLE);
        exStyle |= WS_EX_TOOLWINDOW;
        exStyle &= ~WS_EX_APPWINDOW;
        SetWindowLong(hwnd, GWL_EXSTYLE, exStyle);

        // Parent to desktop (SHELLDLL_DefView) - same as the portal window
        // This makes the app a sibling of the portal, not a child
        var parentHandle = DesktopHandle != IntPtr.Zero ? DesktopHandle : ContentHandle;
        var prevParent = SetParent(hwnd, parentHandle);
        Debug.WriteLine($"[AppPortal] SetParent to {parentHandle}, prev parent: {prevParent}");

        // Calculate position: portal position + header offset
        int x = PortalScreenPosition.X;
        int y = PortalScreenPosition.Y + HeaderHeight;
        int width = ContentSize.Width;
        int height = ContentSize.Height;
        
        Debug.WriteLine($"[AppPortal] Positioning at ({x}, {y}) with size ({width}, {height})");

        // Position and size the window to match content area exactly
        bool posResult = SetWindowPos(hwnd, HWND_TOP, x, y, width, height, 
            SWP_SHOWWINDOW | SWP_FRAMECHANGED);
        Debug.WriteLine($"[AppPortal] SetWindowPos result: {posResult}");

        // Force show first
        ShowWindow(hwnd, SW_SHOW);

        // Check if the app has a minimum size larger than our content area
        // by looking at the actual window size after setting
        if (GetWindowRect(hwnd, out RECT actualRect))
        {
            int actualWidth = actualRect.Right - actualRect.Left;
            int actualHeight = actualRect.Bottom - actualRect.Top;
            
            Debug.WriteLine($"[AppPortal] Requested size: ({width}, {height}), Actual size: ({actualWidth}, {actualHeight})");
            
            // If app is larger than requested, we need to expand the portal
            if (actualWidth > width || actualHeight > height)
            {
                int neededWidth = Math.Max(actualWidth, width);
                int neededHeight = Math.Max(actualHeight, height);
                
                Debug.WriteLine($"[AppPortal] App needs more space: ({neededWidth}, {neededHeight}), requesting portal resize");
                
                // Request portal resize (add header height to get total portal height needed)
                PortalResizeRequested?.Invoke(neededWidth, neededHeight + HeaderHeight);
                
                // Update our content size
                ContentSize = new System.Drawing.Size(neededWidth, neededHeight);
                width = neededWidth;
                height = neededHeight;
                
                // Reposition and resize with new dimensions
                SetWindowPos(hwnd, HWND_TOP, x, y, width, height, SWP_SHOWWINDOW | SWP_FRAMECHANGED);
            }
        }

        // Create a clipping region to prevent the app from extending beyond the content area
        // This ensures the app is visually contained within the portal
        // Delete the old region first to prevent GDI resource leak
        if (_currentRegion != IntPtr.Zero)
        {
            DeleteObject(_currentRegion);
        }
        _currentRegion = CreateRectRgn(0, 0, width, height);
        int rgnResult = SetWindowRgn(hwnd, _currentRegion, true);
        Debug.WriteLine($"[AppPortal] SetWindowRgn result: {rgnResult}");

        // Force redraw
        RedrawWindow(hwnd, IntPtr.Zero, IntPtr.Zero, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);

        // Verify actual window position
        if (GetWindowRect(hwnd, out RECT verifyRect))
        {
            Debug.WriteLine($"[AppPortal] Final window position: ({verifyRect.Left}, {verifyRect.Top}) size: ({verifyRect.Right - verifyRect.Left}, {verifyRect.Bottom - verifyRect.Top})");
        }

        Debug.WriteLine($"[AppPortal] Embedded window {hwnd} positioned at ({x}, {y}) with clip region ({width}, {height})");
    }

    [RelayCommand]
    private void Stop()
    {
        if (!CanStop) return;

        try
        {
            App.Process?.Kill();
            App.Process?.WaitForExit(1000);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[AppPortal] Stop failed: {ex.Message}");
        }

        App.OnAppExited();
        StatusText = "Stopped. Click Start to relaunch.";
    }

    [RelayCommand]
    private void PopOut()
    {
        if (!CanPopOut || App.WindowHandle == IntPtr.Zero) return;

        try
        {
            var hwnd = App.WindowHandle;

            // Clear clipping region first
            SetWindowRgn(hwnd, IntPtr.Zero, true);

            // Unparent
            SetParent(hwnd, IntPtr.Zero);

            // Restore original styles
            SetWindowLong(hwnd, GWL_STYLE, App.OriginalStyle);
            SetWindowLong(hwnd, GWL_EXSTYLE, App.OriginalExStyle);

            // Restore original position/size
            MoveWindow(hwnd, 
                App.OriginalBounds.X, App.OriginalBounds.Y,
                App.OriginalBounds.Width, App.OriginalBounds.Height, 
                true);

            ShowWindow(hwnd, SW_SHOW);
            
            // Force a redraw to ensure proper rendering
            RedrawWindow(hwnd, IntPtr.Zero, IntPtr.Zero, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);

            App.State = AppPortalState.PoppedOut;
            StatusText = "Running outside portal";

            OnPropertyChanged(nameof(CanStart));
            OnPropertyChanged(nameof(CanStop));
            OnPropertyChanged(nameof(CanPopOut));
            OnPropertyChanged(nameof(CanPopIn));

            Debug.WriteLine($"[AppPortal] Popped out window {hwnd}");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[AppPortal] PopOut failed: {ex.Message}");
            StatusText = $"PopOut failed: {ex.Message}";
        }
    }

    [RelayCommand]
    private void PopIn()
    {
        if (!CanPopIn || App.WindowHandle == IntPtr.Zero) return;

        try
        {
            EmbedWindow(App.WindowHandle);
            App.State = AppPortalState.Running;
            StatusText = "Running";

            OnPropertyChanged(nameof(CanStart));
            OnPropertyChanged(nameof(CanStop));
            OnPropertyChanged(nameof(CanPopOut));
            OnPropertyChanged(nameof(CanPopIn));
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[AppPortal] PopIn failed: {ex.Message}");
        }
    }

    /// <summary>
    /// Called when the portal content area is resized or moved
    /// </summary>
    public void OnContentResized(int width, int height)
    {
        ContentSize = new System.Drawing.Size(width, height);

        if (App.State == AppPortalState.Running && App.WindowHandle != IntPtr.Zero)
        {
            int x = PortalScreenPosition.X;
            int y = PortalScreenPosition.Y + HeaderHeight;
            
            // Update position and size with FRAMECHANGED to ensure proper redraw
            SetWindowPos(App.WindowHandle, HWND_TOP, x, y, width, height, SWP_SHOWWINDOW | SWP_FRAMECHANGED);
            
            // Update clipping region to match new size
            // Delete the old region first to prevent GDI resource leak
            if (_currentRegion != IntPtr.Zero)
            {
                DeleteObject(_currentRegion);
            }
            _currentRegion = CreateRectRgn(0, 0, width, height);
            SetWindowRgn(App.WindowHandle, _currentRegion, true);
        }
    }

    /// <summary>
    /// Called when the portal window is moved
    /// </summary>
    public void OnPortalMoved(int screenX, int screenY)
    {
        PortalScreenPosition = new System.Drawing.Point(screenX, screenY);

        if (App.State == AppPortalState.Running && App.WindowHandle != IntPtr.Zero)
        {
            int x = screenX;
            int y = screenY + HeaderHeight;
            SetWindowPos(App.WindowHandle, HWND_TOP, x, y, ContentSize.Width, ContentSize.Height, SWP_SHOWWINDOW);
        }
    }

    /// <summary>
    /// Shows or hides the embedded app window (for roll-up functionality)
    /// </summary>
    public void SetEmbeddedWindowVisibility(bool visible)
    {
        if (App.State == AppPortalState.Running && App.WindowHandle != IntPtr.Zero)
        {
            ShowWindow(App.WindowHandle, visible ? SW_SHOW : SW_HIDE);
        }
    }

    public void Dispose()
    {
        // Clean up GDI region handle
        if (_currentRegion != IntPtr.Zero)
        {
            DeleteObject(_currentRegion);
            _currentRegion = IntPtr.Zero;
        }

        if (App.State == AppPortalState.Running)
        {
            try { App.Process?.Kill(); } catch { }
        }
        App.AppExited -= OnAppExited;
    }

    #region Win32 APIs

    private const int GWL_STYLE = -16;
    private const int GWL_EXSTYLE = -20;
    
    private const int WS_CAPTION = 0x00C00000;
    private const int WS_THICKFRAME = 0x00040000;
    private const int WS_MINIMIZEBOX = 0x00020000;
    private const int WS_MAXIMIZEBOX = 0x00010000;
    private const int WS_SYSMENU = 0x00080000;
    private const int WS_CHILD = 0x40000000;
    
    private const int WS_EX_TOOLWINDOW = 0x00000080;
    private const int WS_EX_APPWINDOW = 0x00040000;

    private const int SW_SHOW = 5;
    private const int SW_HIDE = 0;

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left, Top, Right, Bottom;
    }

    [DllImport("user32.dll")]
    private static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

    [DllImport("user32.dll")]
    private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll")]
    private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    // SetWindowPos constants
    private static readonly IntPtr HWND_TOP = IntPtr.Zero;
    private const uint SWP_SHOWWINDOW = 0x0040;
    private const uint SWP_FRAMECHANGED = 0x0020;

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

    // RedrawWindow constants
    private const uint RDW_INVALIDATE = 0x0001;
    private const uint RDW_UPDATENOW = 0x0100;
    private const uint RDW_ALLCHILDREN = 0x0080;

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool RedrawWindow(IntPtr hWnd, IntPtr lprcUpdate, IntPtr hrgnUpdate, uint flags);

    // Region functions for clipping
    [DllImport("gdi32.dll")]
    private static extern IntPtr CreateRectRgn(int nLeftRect, int nTopRect, int nRightRect, int nBottomRect);

    [DllImport("user32.dll")]
    private static extern int SetWindowRgn(IntPtr hWnd, IntPtr hRgn, bool bRedraw);

    [DllImport("gdi32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DeleteObject(IntPtr hObject);

    #endregion
}
