using System.Diagnostics;
using CommunityToolkit.Mvvm.ComponentModel;

namespace DeskGrid.Models;

/// <summary>
/// Represents an application that can be embedded in a portal
/// </summary>
public partial class EmbeddedApp : ObservableObject
{
    [ObservableProperty]
    private string _name = string.Empty;

    [ObservableProperty]
    private string _executablePath = string.Empty;

    [ObservableProperty]
    private string? _arguments;

    [ObservableProperty]
    private AppPortalState _state = AppPortalState.Stopped;

    [ObservableProperty]
    private int _processId;

    [ObservableProperty]
    private IntPtr _windowHandle = IntPtr.Zero;

    /// <summary>
    /// Original window style before embedding (for restore on pop-out)
    /// </summary>
    public int OriginalStyle { get; set; }

    /// <summary>
    /// Original extended window style before embedding
    /// </summary>
    public int OriginalExStyle { get; set; }

    /// <summary>
    /// Original window bounds before embedding (for restore on pop-out)
    /// </summary>
    public System.Drawing.Rectangle OriginalBounds { get; set; }

    /// <summary>
    /// The running process instance
    /// </summary>
    public Process? Process { get; set; }

    /// <summary>
    /// Event raised when the embedded app exits
    /// </summary>
    public event EventHandler? AppExited;

    /// <summary>
    /// Raises the AppExited event
    /// </summary>
    public void OnAppExited()
    {
        State = AppPortalState.Stopped;
        ProcessId = 0;
        WindowHandle = IntPtr.Zero;
        Process = null;
        AppExited?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>
    /// Gets the display name for the app
    /// </summary>
    public string DisplayName => string.IsNullOrEmpty(Name) 
        ? System.IO.Path.GetFileNameWithoutExtension(ExecutablePath) 
        : Name;
}
