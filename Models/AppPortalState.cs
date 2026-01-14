namespace DeskGrid.Models;

/// <summary>
/// Represents the state of an app portal
/// </summary>
public enum AppPortalState
{
    /// <summary>App is not running, Start button visible</summary>
    Stopped,
    
    /// <summary>App is running and embedded in portal</summary>
    Running,
    
    /// <summary>App is running outside the portal (popped out)</summary>
    PoppedOut
}
