using System.Runtime.InteropServices;
using System.Text;
using H.Hooks;

namespace DeskGrid.Services;

/// <summary>
/// Event args for Portal drawing operations
/// </summary>
public class PortalDrawEventArgs : EventArgs
{
    public System.Drawing.Rectangle Bounds { get; set; }
    public System.Drawing.Point StartPoint { get; set; }
    public System.Drawing.Point CurrentPoint { get; set; }
}

/// <summary>
/// Low-level mouse hook to detect double-click on empty desktop space
/// and drag-to-create Portal functionality
/// </summary>
public class InputListener : IDisposable
{
    private LowLevelMouseHook? _mouseHook;
    private DateTime _lastClickTime = DateTime.MinValue;
    private System.Drawing.Point _lastClickPoint;
    private bool _disposed;
    
    // Drag-to-create state
    private bool _isDrawingPortal;
    private System.Drawing.Point _drawStartPoint;
    private DateTime _doubleClickTime = DateTime.MinValue;
    private bool _waitingForDragAfterDoubleClick;
    private bool _ctrlClickMode; // True if using Ctrl+click to draw (not double-click)

    /// <summary>
    /// Fired when user double-clicks on empty desktop space (quick double-click, no drag)
    /// </summary>
    public event EventHandler? DesktopDoubleClick;
    
    /// <summary>
    /// Fired when user starts drawing a Portal (double-click + hold + drag)
    /// </summary>
    public event EventHandler<PortalDrawEventArgs>? PortalDrawStart;
    
    /// <summary>
    /// Fired during Portal drawing with updated bounds
    /// </summary>
    public event EventHandler<PortalDrawEventArgs>? PortalDrawing;
    
    /// <summary>
    /// Fired when Portal drawing is complete
    /// </summary>
    public event EventHandler<PortalDrawEventArgs>? PortalDrawEnd;

    /// <summary>
    /// Starts listening for mouse events
    /// </summary>
    public void Start()
    {
        try
        {
            System.Diagnostics.Debug.WriteLine("[InputListener] Starting mouse hook...");
            _mouseHook = new LowLevelMouseHook();
            _mouseHook.Down += OnMouseDown;
            _mouseHook.Up += OnMouseUp;
            _mouseHook.Move += OnMouseMove;
            _mouseHook.Start();
            System.Diagnostics.Debug.WriteLine("[InputListener] Mouse hook started successfully");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[InputListener] ERROR starting mouse hook: {ex.Message}");
        }
    }

    /// <summary>
    /// Stops listening for mouse events
    /// </summary>
    public void Stop()
    {
        if (_mouseHook != null)
        {
            _mouseHook.Down -= OnMouseDown;
            _mouseHook.Up -= OnMouseUp;
            _mouseHook.Move -= OnMouseMove;
            _mouseHook.Stop();
            _mouseHook.Dispose();
            _mouseHook = null;
        }
    }

    private void OnMouseDown(object? sender, H.Hooks.MouseEventArgs e)
    {
        // Only process left clicks
        if (!e.Keys.Values.Contains(Key.MouseLeft))
            return;

        System.Diagnostics.Debug.WriteLine($"[InputListener] MouseDown at {e.Position}");

        // Check if click is on desktop
        var isOnDesktop = IsClickOnDesktop(e.Position);
        System.Diagnostics.Debug.WriteLine($"[InputListener] IsOnDesktop: {isOnDesktop}");
        
        if (!isOnDesktop)
        {
            _waitingForDragAfterDoubleClick = false;
            return;
        }

        var now = DateTime.Now;
        var doubleClickTimeMs = TimeSpan.FromMilliseconds(GetDoubleClickTime());
        var timeDelta = now - _lastClickTime;
        
        // Check if Ctrl is held - Ctrl+click+drag creates a portal
        bool ctrlHeld = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
        if (ctrlHeld)
        {
            System.Diagnostics.Debug.WriteLine("[InputListener] Ctrl+click detected, starting portal draw mode");
            _waitingForDragAfterDoubleClick = true;
            _ctrlClickMode = true; // Mark as Ctrl mode (not double-click)
            _drawStartPoint = e.Position;
            _lastClickTime = DateTime.MinValue;
            return;
        }

        // Check if this is a double-click (original behavior)
        if (timeDelta <= doubleClickTimeMs && IsWithinDoubleClickDistance(e.Position, _lastClickPoint))
        {
            System.Diagnostics.Debug.WriteLine("[InputListener] Double-click detected, waiting for drag");
            // This is a double-click - wait to see if user drags or releases
            _waitingForDragAfterDoubleClick = true;
            _doubleClickTime = now;
            _drawStartPoint = e.Position;
            _lastClickTime = DateTime.MinValue; // Reset
        }
        else
        {
            _lastClickTime = now;
            _lastClickPoint = e.Position;
            _waitingForDragAfterDoubleClick = false;
        }
    }

    private void OnMouseMove(object? sender, H.Hooks.MouseEventArgs e)
    {
        // Check if we're waiting for drag after double-click
        if (_waitingForDragAfterDoubleClick && !_isDrawingPortal)
        {
            // Check if mouse moved enough to start drawing
            var distance = Math.Sqrt(
                Math.Pow(e.Position.X - _drawStartPoint.X, 2) + 
                Math.Pow(e.Position.Y - _drawStartPoint.Y, 2));
            
            if (distance > 10) // Minimum drag distance to start drawing
            {
                System.Diagnostics.Debug.WriteLine($"[InputListener] PortalDrawStart at {_drawStartPoint}");
                _isDrawingPortal = true;
                _waitingForDragAfterDoubleClick = false;
                
                var args = CreateDrawEventArgs(_drawStartPoint, e.Position);
                PortalDrawStart?.Invoke(this, args);
            }
        }
        
        // If actively drawing, fire the drawing event
        if (_isDrawingPortal)
        {
            var args = CreateDrawEventArgs(_drawStartPoint, e.Position);
            PortalDrawing?.Invoke(this, args);
        }
    }

    private void OnMouseUp(object? sender, H.Hooks.MouseEventArgs e)
    {
        // Only process left clicks
        if (!e.Keys.Values.Contains(Key.MouseLeft))
            return;

        if (_isDrawingPortal)
        {
            // Finish drawing the Portal
            var args = CreateDrawEventArgs(_drawStartPoint, e.Position);
            System.Diagnostics.Debug.WriteLine($"[InputListener] PortalDrawEnd at {args.Bounds}");
            PortalDrawEnd?.Invoke(this, args);
            
            _isDrawingPortal = false;
            _waitingForDragAfterDoubleClick = false;
            _ctrlClickMode = false;
        }
        else if (_waitingForDragAfterDoubleClick)
        {
            // Ctrl+click without drag = do nothing (just cancel)
            // Double-click without drag = toggle visibility
            if (!_ctrlClickMode)
            {
                System.Diagnostics.Debug.WriteLine("[InputListener] DesktopDoubleClick - toggling visibility");
                DesktopDoubleClick?.Invoke(this, EventArgs.Empty);
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("[InputListener] Ctrl+click released without drag - cancelled");
            }
            _waitingForDragAfterDoubleClick = false;
            _ctrlClickMode = false;
        }
    }

    private PortalDrawEventArgs CreateDrawEventArgs(System.Drawing.Point start, System.Drawing.Point current)
    {
        int x = Math.Min(start.X, current.X);
        int y = Math.Min(start.Y, current.Y);
        int width = Math.Abs(current.X - start.X);
        int height = Math.Abs(current.Y - start.Y);

        return new PortalDrawEventArgs
        {
            StartPoint = start,
            CurrentPoint = current,
            Bounds = new System.Drawing.Rectangle(x, y, width, height)
        };
    }

    private bool IsClickOnDesktop(System.Drawing.Point point)
    {
        try
        {
            // Get window under cursor
            var hwnd = WindowFromPoint(point);
            if (hwnd == IntPtr.Zero)
                return false;

            // Get window class name
            var className = new StringBuilder(256);
            GetClassName(hwnd, className, 256);
            var windowClass = className.ToString();

            // Check if it's the desktop (SysListView32 is the desktop icon list)
            if (windowClass == "SysListView32" || windowClass == "SHELLDLL_DefView" || 
                windowClass == "WorkerW" || windowClass == "Progman")
            {
                // Additional check: make sure we're not clicking on an icon
                if (windowClass == "SysListView32")
                {
                    return !IsClickOnDesktopIcon(hwnd, point);
                }
                return true;
            }

            return false;
        }
        catch
        {
            return false;
        }
    }

    private bool IsClickOnDesktopIcon(IntPtr listViewHwnd, System.Drawing.Point screenPoint)
    {
        try
        {
            var clientPoint = new POINT { X = screenPoint.X, Y = screenPoint.Y };
            ScreenToClient(listViewHwnd, ref clientPoint);
            return false;
        }
        catch
        {
            return false;
        }
    }

    private bool IsWithinDoubleClickDistance(System.Drawing.Point p1, System.Drawing.Point p2)
    {
        int cxDouble = GetSystemMetrics(SM_CXDOUBLECLK);
        int cyDouble = GetSystemMetrics(SM_CYDOUBLECLK);
        return Math.Abs(p1.X - p2.X) <= cxDouble && Math.Abs(p1.Y - p2.Y) <= cyDouble;
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        Stop();
        GC.SuppressFinalize(this);
    }

    #region P/Invoke

    private const int SM_CXDOUBLECLK = 36;
    private const int SM_CYDOUBLECLK = 37;

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT
    {
        public int X;
        public int Y;
    }

    [DllImport("user32.dll")]
    private static extern uint GetDoubleClickTime();

    [DllImport("user32.dll")]
    private static extern int GetSystemMetrics(int nIndex);

    [DllImport("user32.dll")]
    private static extern IntPtr WindowFromPoint(System.Drawing.Point Point);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    private static extern bool ScreenToClient(IntPtr hWnd, ref POINT lpPoint);
    
    [DllImport("user32.dll")]
    private static extern short GetKeyState(int nVirtKey);
    
    private const int VK_CONTROL = 0x11;

    #endregion
}
