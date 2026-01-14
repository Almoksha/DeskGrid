using System.Windows;

namespace DeskGrid.Services;

/// <summary>
/// Service for snapping portal windows to screen edges and other portals
/// </summary>
public class SnapService
{
    /// <summary>
    /// Snap threshold in pixels - windows snap when within this distance
    /// </summary>
    public int SnapThreshold { get; set; } = 15;

    /// <summary>
    /// Calculates snap-adjusted position for a portal being dragged
    /// </summary>
    /// <param name="currentBounds">Current bounds of the portal being dragged</param>
    /// <param name="otherPortals">Bounds of other portal windows</param>
    /// <returns>Adjusted position with snapping applied</returns>
    public System.Drawing.Point GetSnappedPosition(
        System.Drawing.Rectangle currentBounds,
        IEnumerable<System.Drawing.Rectangle> otherPortals)
    {
        int snapX = currentBounds.X;
        int snapY = currentBounds.Y;

        // Get all screen bounds
        var screens = System.Windows.Forms.Screen.AllScreens;

        // Check screen edge snapping for each monitor
        foreach (var screen in screens)
        {
            var workArea = screen.WorkingArea;

            // Snap to left edge
            if (Math.Abs(currentBounds.Left - workArea.Left) < SnapThreshold)
                snapX = workArea.Left;

            // Snap to right edge
            if (Math.Abs(currentBounds.Right - workArea.Right) < SnapThreshold)
                snapX = workArea.Right - currentBounds.Width;

            // Snap to top edge
            if (Math.Abs(currentBounds.Top - workArea.Top) < SnapThreshold)
                snapY = workArea.Top;

            // Snap to bottom edge
            if (Math.Abs(currentBounds.Bottom - workArea.Bottom) < SnapThreshold)
                snapY = workArea.Bottom - currentBounds.Height;
        }

        // Check snapping to other portals
        foreach (var other in otherPortals)
        {
            // Skip if same portal (bounds match exactly)
            if (other == currentBounds) continue;

            // Snap to left edge of other portal
            if (Math.Abs(currentBounds.Left - other.Right) < SnapThreshold)
                snapX = other.Right;

            // Snap to right edge of other portal
            if (Math.Abs(currentBounds.Right - other.Left) < SnapThreshold)
                snapX = other.Left - currentBounds.Width;

            // Snap to top edge of other portal
            if (Math.Abs(currentBounds.Top - other.Bottom) < SnapThreshold)
                snapY = other.Bottom;

            // Snap to bottom edge of other portal
            if (Math.Abs(currentBounds.Bottom - other.Top) < SnapThreshold)
                snapY = other.Top - currentBounds.Height;

            // Align horizontally with other portal (left edges)
            if (Math.Abs(currentBounds.Left - other.Left) < SnapThreshold)
                snapX = other.Left;

            // Align horizontally with other portal (right edges)
            if (Math.Abs(currentBounds.Right - other.Right) < SnapThreshold)
                snapX = other.Right - currentBounds.Width;

            // Align vertically with other portal (top edges)
            if (Math.Abs(currentBounds.Top - other.Top) < SnapThreshold)
                snapY = other.Top;

            // Align vertically with other portal (bottom edges)
            if (Math.Abs(currentBounds.Bottom - other.Bottom) < SnapThreshold)
                snapY = other.Bottom - currentBounds.Height;
        }

        return new System.Drawing.Point(snapX, snapY);
    }

    /// <summary>
    /// Gets the bounds of a WPF window in physical pixels
    /// </summary>
    public System.Drawing.Rectangle GetWindowBounds(Window window)
    {
        var hwnd = new System.Windows.Interop.WindowInteropHelper(window).Handle;
        if (hwnd != IntPtr.Zero && GetWindowRect(hwnd, out RECT rect))
        {
            return new System.Drawing.Rectangle(
                rect.Left, rect.Top,
                rect.Right - rect.Left,
                rect.Bottom - rect.Top);
        }
        return new System.Drawing.Rectangle(
            (int)window.Left, (int)window.Top,
            (int)window.Width, (int)window.Height);
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
