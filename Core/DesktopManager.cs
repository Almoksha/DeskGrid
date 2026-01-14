using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Threading;
using DeskGrid.Models;
using DeskGrid.Services;

namespace DeskGrid.Core;

/// <summary>
/// Manages desktop integration - keeping Portal windows at desktop level
/// </summary>
public class DesktopManager : IDisposable
{
    private IntPtr _workerW = IntPtr.Zero;
    private readonly List<Window> _PortalWindows = new();
    private readonly Dictionary<Window, string> _portalIds = new();
    private readonly PortalLayoutService _layoutService = new();
    private bool _PortalsVisible = true;
    private bool _disposed;
    private bool _isShuttingDown = false; // Prevents save during app shutdown
    private bool _isLoadingPortals = false; // Prevents save during initial load
    private bool _attachedSuccessfully = false; // Prevents save if we never attached
    private int _loadedPortalCount = 0; // Tracks how many portals were loaded from disk (data loss protection)
    private DispatcherTimer? _zOrderTimer;
    private DispatcherTimer? _autoSaveTimer;

    // Undocumented message to spawn WorkerW
    private const uint SPAWN_WORKER_W = 0x052C;

    /// <summary>
    /// Attaches to the desktop by finding/spawning the WorkerW layer.
    /// Retries multiple times to handle slow Explorer startup after reboot.
    /// </summary>
    public bool AttachToDesktop()
    {
        const int maxRetries = 5;
        const int retryDelayMs = 2000; // 2 seconds between retries
        
        for (int attempt = 1; attempt <= maxRetries; attempt++)
        {
            System.Diagnostics.Debug.WriteLine($"[DesktopManager] AttachToDesktop attempt {attempt}/{maxRetries}");
            
            if (TryAttachToDesktopOnce())
            {
                System.Diagnostics.Debug.WriteLine($"[DesktopManager] Successfully attached on attempt {attempt}");
                _attachedSuccessfully = true;
                return true;
            }
            
            if (attempt < maxRetries)
            {
                System.Diagnostics.Debug.WriteLine($"[DesktopManager] Attempt {attempt} failed, waiting {retryDelayMs}ms before retry...");
                System.Threading.Thread.Sleep(retryDelayMs);
            }
        }
        
        System.Diagnostics.Debug.WriteLine($"[DesktopManager] Failed to attach after {maxRetries} attempts");
        return false;
    }
    
    /// <summary>
    /// Single attempt to attach to desktop
    /// </summary>
    private bool TryAttachToDesktopOnce()
    {
        try
        {
            // Find the Progman window
            var progman = FindWindow("Progman", null);
            if (progman == IntPtr.Zero)
            {
                System.Diagnostics.Debug.WriteLine("Failed to find Progman");
                return false;
            }

            // First, check if SHELLDLL_DefView already exists (this is what we actually need for parenting portals)
            var existingShellView = FindShellDefView();
            System.Diagnostics.Debug.WriteLine($"Existing SHELLDLL_DefView: {existingShellView}");
            
            // Try to find existing WorkerW first
            _workerW = FindDesktopWorkerW();
            System.Diagnostics.Debug.WriteLine($"Existing WorkerW before spawn: {_workerW}");
            
            // Only send spawn message if WorkerW doesn't exist yet
            if (_workerW == IntPtr.Zero)
            {
                System.Diagnostics.Debug.WriteLine("No existing WorkerW found, sending spawn message...");
                SendMessage(progman, SPAWN_WORKER_W, (IntPtr)0x0000000D, IntPtr.Zero);
                
                // Give the system a moment to create the window
                System.Threading.Thread.Sleep(100);
                
                // Try finding WorkerW again
                _workerW = FindDesktopWorkerW();
                System.Diagnostics.Debug.WriteLine($"WorkerW after spawn: {_workerW}");
            }

            // Re-check SHELLDLL_DefView after potential spawn
            var shellView = FindShellDefView();
            System.Diagnostics.Debug.WriteLine($"SHELLDLL_DefView after spawn: {shellView}");
            
            // The critical requirement is that SHELLDLL_DefView exists - that's where we parent our portals
            // WorkerW is nice to have for z-order but not essential since we use SetParent
            if (shellView == IntPtr.Zero)
            {
                System.Diagnostics.Debug.WriteLine("Failed to find SHELLDLL_DefView - cannot attach to desktop");
                return false;
            }

            // If we still don't have WorkerW but we have SHELLDLL_DefView, we can still work
            if (_workerW == IntPtr.Zero)
            {
                System.Diagnostics.Debug.WriteLine("WorkerW not found but SHELLDLL_DefView exists - proceeding anyway");
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"Found WorkerW: {_workerW}");
            }
            
            // Z-order timer is not needed when using SetParent for true embedding
            // StartZOrderTimer();
            
            return true;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"AttachToDesktop failed: {ex.Message}");
            return false;
        }
    }

    private void StartZOrderTimer()
    {
        _zOrderTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(500) // Slower interval to reduce conflicts
        };
        _zOrderTimer.Tick += (s, e) => SendPortalsToBottom();
        _zOrderTimer.Start();
    }

    /// <summary>
    /// Temporarily pauses z-order enforcement (call when user is interacting with a Portal)
    /// </summary>
    public void PauseZOrderEnforcement()
    {
        _zOrderTimer?.Stop();
    }

    /// <summary>
    /// Resumes z-order enforcement after user interaction ends
    /// </summary>
    public void ResumeZOrderEnforcement()
    {
        // Delay resume to prevent immediate re-ordering during drag
        if (_zOrderTimer != null && !_zOrderTimer.IsEnabled)
        {
            // Use a small delay before resuming
            var resumeTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
            resumeTimer.Tick += (s, e) =>
            {
                resumeTimer.Stop();
                _zOrderTimer?.Start();
            };
            resumeTimer.Start();
        }
    }

    private void SendPortalsToBottom()
    {
        if (!_PortalsVisible) return;

        // Check if any Portal is being interacted with (has mouse over it)
        foreach (var window in _PortalWindows)
        {
            if (window.IsMouseOver || window.IsActive)
            {
                return; // Skip z-order enforcement when user is interacting
            }
        }

        foreach (var window in _PortalWindows)
        {
            try
            {
                var hwnd = new WindowInteropHelper(window).Handle;
                if (hwnd != IntPtr.Zero)
                {
                    // Position window just above the desktop (WorkerW)
                    SetWindowPos(hwnd, _workerW != IntPtr.Zero ? _workerW : HWND_BOTTOM,
                        0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
                }
            }
            catch { }
        }
    }

    /// <summary>
    /// Finds the correct WorkerW window (the one without SHELLDLL_DefView as child)
    /// </summary>
    private IntPtr FindDesktopWorkerW()
    {
        IntPtr result = IntPtr.Zero;

        EnumWindows((hwnd, lParam) =>
        {
            // Find WorkerW windows
            var shellView = FindWindowEx(hwnd, IntPtr.Zero, "SHELLDLL_DefView", null);
            
            if (shellView != IntPtr.Zero)
            {
                // Found the WorkerW with the shell view, now find the one AFTER this
                result = FindWindowEx(IntPtr.Zero, hwnd, "WorkerW", null);
            }

            return true; // Continue enumeration
        }, IntPtr.Zero);

        // If we didn't find it that way, try to find any WorkerW that's a child of Progman
        if (result == IntPtr.Zero)
        {
            var progman = FindWindow("Progman", null);
            if (progman != IntPtr.Zero)
            {
                result = FindWindowEx(progman, IntPtr.Zero, "WorkerW", null);
            }
        }

        return result;
    }

    /// <summary>
    /// Attaches a WPF window to the desktop (parented to SHELLDLL_DefView like DeskFrame)
    /// </summary>
    public bool AttachWindow(Window window)
    {
        System.Diagnostics.Debug.WriteLine($"[DesktopManager] AttachWindow called");

        try
        {
            // Find SHELLDLL_DefView (the desktop icons container)
            IntPtr shellView = FindShellDefView();
            if (shellView == IntPtr.Zero)
            {
                System.Diagnostics.Debug.WriteLine("[DesktopManager] SHELLDLL_DefView not found!");
                return false;
            }
            System.Diagnostics.Debug.WriteLine($"[DesktopManager] Found SHELLDLL_DefView: {shellView}");

            var hwnd = new WindowInteropHelper(window).EnsureHandle();
            
            // Store the desired screen position and size
            var screenLeft = (int)window.Left;
            var screenTop = (int)window.Top;
            var width = (int)window.Width;
            var height = (int)window.Height;
            
            System.Diagnostics.Debug.WriteLine($"[DesktopManager] Window hwnd: {hwnd}, Screen Position: ({screenLeft}, {screenTop}), Size: {width}x{height}");
            
            // Set parent to SHELLDLL_DefView (desktop icons container)
            var prevParent = SetParent(hwnd, shellView);
            System.Diagnostics.Debug.WriteLine($"[DesktopManager] SetParent result (prev parent): {prevParent}");

            // Change window style: remove WS_POPUP, add WS_CHILD
            int style = GetWindowLong(hwnd, GWL_STYLE);
            style &= ~WS_POPUP;  // Remove popup
            style |= WS_CHILD;   // Add child
            SetWindowLong(hwnd, GWL_STYLE, style);
            System.Diagnostics.Debug.WriteLine("[DesktopManager] Window style changed to WS_CHILD");
            
            // Set extended styles - tool window (no taskbar), no activate
            var exStyle = GetWindowLong(hwnd, GWL_EXSTYLE);
            exStyle |= WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE;
            SetWindowLong(hwnd, GWL_EXSTYLE, exStyle);
            System.Diagnostics.Debug.WriteLine("[DesktopManager] Extended styles set");

            // Convert screen coordinates to parent-relative (client) coordinates
            POINT pt = new POINT { X = screenLeft, Y = screenTop };
            ScreenToClient(shellView, ref pt);
            System.Diagnostics.Debug.WriteLine($"[DesktopManager] Converted to client coords: ({pt.X}, {pt.Y})");

            // Position the window using Win32 API
            SetWindowPos(hwnd, IntPtr.Zero, pt.X, pt.Y, width, height, 
                SWP_NOZORDER | SWP_NOACTIVATE | SWP_SHOWWINDOW);
            System.Diagnostics.Debug.WriteLine("[DesktopManager] SetWindowPos called with SWP_SHOWWINDOW");
            
            _PortalWindows.Add(window);
            
            // Assign an ID if this window doesn't have one yet (new portal, not loaded)
            if (!_portalIds.ContainsKey(window))
            {
                _portalIds[window] = Guid.NewGuid().ToString();
                System.Diagnostics.Debug.WriteLine($"[DesktopManager] Assigned new ID: {_portalIds[window]}");
            }
            
            System.Diagnostics.Debug.WriteLine("[DesktopManager] Window attached to desktop successfully");
            
            // Auto-save layout after attaching a new portal
            SaveLayout();
            
            return true;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[DesktopManager] AttachWindow failed: {ex.Message}");
            return false;
        }
    }

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

    /// <summary>
    /// Detaches a window from the desktop
    /// </summary>
    public void DetachWindow(Window window)
    {
        System.Diagnostics.Debug.WriteLine($"[DesktopManager] DetachWindow called for '{window.Title}', isShuttingDown: {_isShuttingDown}");
        _PortalWindows.Remove(window);
        _portalIds.Remove(window);
        
        // Don't save during shutdown - the main save happens before windows close
        if (!_isShuttingDown)
        {
            SaveLayout();
        }
    }
    
    /// <summary>
    /// Marks the manager as shutting down (prevents saves during window close)
    /// </summary>
    public void BeginShutdown()
    {
        System.Diagnostics.Debug.WriteLine("[DesktopManager] BeginShutdown - disabling saves");
        _isShuttingDown = true;
    }

    /// <summary>
    /// Toggles visibility of all Portal windows (double-click to hide feature)
    /// </summary>
    public void TogglePortalsVisibility()
    {
        _PortalsVisible = !_PortalsVisible;
        System.Diagnostics.Debug.WriteLine($"[DesktopManager] TogglePortalsVisibility: {_PortalsVisible}");

        foreach (var window in _PortalWindows)
        {
            if (_PortalsVisible)
            {
                window.Visibility = Visibility.Visible;
            }
            else
            {
                window.Visibility = Visibility.Collapsed;
            }
        }
    }

    /// <summary>
    /// Gets or sets whether Portals are currently visible
    /// </summary>
    public bool PortalsVisible
    {
        get => _PortalsVisible;
        set
        {
            if (_PortalsVisible != value)
            {
                _PortalsVisible = value;
                foreach (var window in _PortalWindows)
                {
                    window.Visibility = value ? Visibility.Visible : Visibility.Collapsed;
                }
            }
        }
    }

    /// <summary>
    /// Gets the WorkerW handle
    /// </summary>
    public IntPtr WorkerWHandle => _workerW;

    /// <summary>
    /// Creates and attaches a new Portal window
    /// </summary>
    public Views.PortalWindow CreatePortal(string title, string folderPath)
    {
        var Portal = new Views.PortalWindow
        {
            Title = title
        };
        Portal.SetFolder(folderPath);
        
        // Sync the title to ViewModel so it persists correctly
        if (Portal.DataContext is ViewModels.PortalViewModel vm)
        {
            vm.Title = title;
        }
        
        Portal.Show();
        AttachWindow(Portal);
        
        return Portal;
    }

    /// <summary>
    /// Creates an empty Portal at the specified bounds (from drag-to-create)
    /// </summary>
    public Views.PortalWindow CreateEmptyPortal(System.Drawing.Rectangle bounds, string title = "New Portal")
    {
        System.Diagnostics.Debug.WriteLine($"[DesktopManager] CreateEmptyPortal: {bounds}, title: {title}");
        
        // Ensure minimum size
        var width = Math.Max(bounds.Width, 150);
        var height = Math.Max(bounds.Height, 100);

        System.Diagnostics.Debug.WriteLine($"[DesktopManager] Creating PortalWindow with size {width}x{height}");
        var Portal = new Views.PortalWindow
        {
            Title = title,
            Left = bounds.X,
            Top = bounds.Y,
            Width = width,
            Height = height
        };
        
        System.Diagnostics.Debug.WriteLine("[DesktopManager] Showing Portal window...");
        Portal.Show();
        
        System.Diagnostics.Debug.WriteLine("[DesktopManager] Attaching Portal to desktop...");
        var attached = AttachWindow(Portal);
        System.Diagnostics.Debug.WriteLine($"[DesktopManager] Attached: {attached}");
        
        if (!attached)
        {
            // Still track the portal even if attachment failed, so it gets saved
            System.Diagnostics.Debug.WriteLine($"[DesktopManager] AttachWindow failed for empty portal, but still tracking for persistence");
            _PortalWindows.Add(Portal);
            SaveLayout();
        }
        
        // Register portal ID for tracking
        var portalId = Guid.NewGuid().ToString();
        _portalIds[Portal] = portalId;
        
        return Portal;
    }

    /// <summary>
    /// Creates an app Portal for embedding an application
    /// </summary>
    public Views.AppPortalWindow? CreateAppPortal(string executablePath, string? name = null)
    {
        System.Diagnostics.Debug.WriteLine($"[DesktopManager] CreateAppPortal: {executablePath}, name: {name}");
        
        if (!System.IO.File.Exists(executablePath))
        {
            System.Diagnostics.Debug.WriteLine($"[DesktopManager] Executable not found: {executablePath}");
            return null;
        }

        var Portal = new Views.AppPortalWindow(executablePath, name)
        {
            Left = 100,
            Top = 100,
            Width = 400,
            Height = 300
        };
        
        System.Diagnostics.Debug.WriteLine("[DesktopManager] Showing app Portal window...");
        Portal.Show();
        
        System.Diagnostics.Debug.WriteLine("[DesktopManager] Attaching app Portal to desktop...");
        var attached = AttachWindow(Portal);
        System.Diagnostics.Debug.WriteLine($"[DesktopManager] App Portal attached: {attached}");
        
        // IMPORTANT: Update PortalScreenPosition AFTER attachment to get the actual physical position
        if (attached && Portal.DataContext is ViewModels.AppPortalViewModel vm)
        {
            var hwnd = new System.Windows.Interop.WindowInteropHelper(Portal).Handle;
            if (GetWindowRect(hwnd, out RECT rect))
            {
                vm.PortalScreenPosition = new System.Drawing.Point(rect.Left, rect.Top);
                System.Diagnostics.Debug.WriteLine($"[DesktopManager] Updated PortalScreenPosition after attachment: ({rect.Left}, {rect.Top})");
            }
        }
        
        // Track portal ID
        var portalId = Guid.NewGuid().ToString();
        _portalIds[Portal] = portalId;
        
        return Portal;
    }

    /// <summary>
    /// Creates a URL Portal for bookmarks
    /// </summary>
    public Views.UrlPortalWindow? CreateUrlPortal(string title = "Bookmarks")
    {
        System.Diagnostics.Debug.WriteLine($"[DesktopManager] CreateUrlPortal: {title}");
        
        var portal = new Views.UrlPortalWindow
        {
            Title = title,
            Left = 100,
            Top = 100
        };
        
        portal.Show();
        AttachWindow(portal);
        
        var portalId = Guid.NewGuid().ToString();
        _portalIds[portal] = portalId;
        
        return portal;
    }

    /// <summary>
    /// Gets the folder path for a portal by its ID (for auto-sort rules)
    /// </summary>
    public string? GetPortalFolderById(string portalId)
    {
        foreach (var kvp in _portalIds)
        {
            if (kvp.Value == portalId && kvp.Key is Views.PortalWindow portalWindow)
            {
                if (portalWindow.DataContext is ViewModels.PortalViewModel vm)
                {
                    return vm.FolderPath;
                }
            }
        }
        return null;
    }

    /// <summary>
    /// Gets info about all folder portals (for rules configuration)
    /// </summary>
    public List<(string Id, string Title)> GetAllPortalInfo()
    {
        var result = new List<(string Id, string Title)>();
        
        foreach (var kvp in _portalIds)
        {
            if (kvp.Key is Views.PortalWindow portalWindow && 
                portalWindow.DataContext is ViewModels.PortalViewModel vm)
            {
                result.Add((kvp.Value, vm.Title));
            }
        }
        
        return result;
    }

    /// <summary>
    /// Saves the current portal layout to disk
    /// </summary>
    public void SaveLayout()
    {
        // Don't save during initial load - we'd overwrite with incomplete data
        if (_isLoadingPortals)
        {
            System.Diagnostics.Debug.WriteLine("[DesktopManager] Skipping save - currently loading portals");
            return;
        }
        
        // Don't save if we never successfully attached - we'd wipe the config
        if (!_attachedSuccessfully)
        {
            System.Diagnostics.Debug.WriteLine("[DesktopManager] Skipping save - never attached to desktop successfully");
            return;
        }
        
        // CRITICAL: Don't save if we have fewer portals than we loaded (data loss protection)
        // This can happen during startup race conditions where some portals fail to load
        // Exception: during shutdown we want to save the final state
        if (!_isShuttingDown && _loadedPortalCount > 0 && _PortalWindows.Count < _loadedPortalCount)
        {
            System.Diagnostics.Debug.WriteLine($"[DesktopManager] BLOCKING SAVE - would lose data! Current: {_PortalWindows.Count}, Loaded: {_loadedPortalCount}");
            return;
        }
        
        System.Diagnostics.Debug.WriteLine("[DesktopManager] === SAVING LAYOUT ===");
        System.Diagnostics.Debug.WriteLine($"[DesktopManager] _PortalWindows.Count = {_PortalWindows.Count}");
        System.Diagnostics.Debug.WriteLine($"[DesktopManager] _portalIds.Count = {_portalIds.Count}");
        System.Diagnostics.Debug.WriteLine($"[DesktopManager] _PortalsVisible = {_PortalsVisible}");
        
        // Log each window in _PortalWindows
        for (int i = 0; i < _PortalWindows.Count; i++)
        {
            var w = _PortalWindows[i];
            System.Diagnostics.Debug.WriteLine($"[DesktopManager] _PortalWindows[{i}]: {w.GetType().Name}, Title='{w.Title}', IsLoaded={w.IsLoaded}");
        }
        
        try
        {
            var config = new DeskGridConfig
            {
                PortalsVisible = _PortalsVisible,
                Portals = new List<PortalConfig>()
            };

            foreach (var window in _PortalWindows)
            {
                string id;
                if (_portalIds.TryGetValue(window, out var existingId))
                {
                    id = existingId;
                }
                else
                {
                    // Generate new ID and store it so subsequent saves use the same ID
                    id = Guid.NewGuid().ToString();
                    _portalIds[window] = id;
                    System.Diagnostics.Debug.WriteLine($"[DesktopManager] Generated and stored new ID for window: {id}");
                }
                
                if (window is Views.PortalWindow portalWindow)
                {
                    var portalConfig = _layoutService.CreateConfigFromWindow(portalWindow, id);
                    config.Portals.Add(portalConfig);
                    System.Diagnostics.Debug.WriteLine($"[DesktopManager] Added to config: '{portalConfig.Title}', ID: {id}, Type: {portalConfig.PortalType}, pos=({portalConfig.X},{portalConfig.Y})");
                }
                else if (window is Views.AppPortalWindow appPortalWindow)
                {
                    var portalConfig = _layoutService.CreateConfigFromWindow(appPortalWindow, id);
                    config.Portals.Add(portalConfig);
                    System.Diagnostics.Debug.WriteLine($"[DesktopManager] Added app portal to config: '{portalConfig.Title}', ID: {id}");
                }
                else if (window is Views.UrlPortalWindow urlPortalWindow)
                {
                    var portalConfig = _layoutService.CreateConfigFromWindow(urlPortalWindow, id);
                    config.Portals.Add(portalConfig);
                    System.Diagnostics.Debug.WriteLine($"[DesktopManager] Added URL portal to config: '{portalConfig.Title}', ID: {id}, bookmarks: {portalConfig.Bookmarks?.Count ?? 0}");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[DesktopManager] Unknown window type: {window.GetType().Name}");
                }
            }

            System.Diagnostics.Debug.WriteLine($"[DesktopManager] Calling _layoutService.SaveConfig with {config.Portals.Count} portals...");
            _layoutService.SaveConfig(config);
            System.Diagnostics.Debug.WriteLine($"[DesktopManager] === SAVE COMPLETE ===");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[DesktopManager] SaveLayout EXCEPTION: {ex.Message}");
            System.Diagnostics.Debug.WriteLine($"[DesktopManager] Stack trace: {ex.StackTrace}");
        }
    }

    /// <summary>
    /// Gets the current portal configuration (for saving profiles)
    /// </summary>
    public DeskGridConfig GetCurrentConfig()
    {
        var config = new DeskGridConfig
        {
            PortalsVisible = _PortalsVisible
        };

        foreach (var window in _PortalWindows)
        {
            string id;
            if (_portalIds.TryGetValue(window, out var existingId))
            {
                id = existingId;
            }
            else
            {
                id = Guid.NewGuid().ToString();
                _portalIds[window] = id;
            }
            
            if (window is Views.PortalWindow portalWindow)
            {
                config.Portals.Add(_layoutService.CreateConfigFromWindow(portalWindow, id));
            }
            else if (window is Views.AppPortalWindow appPortalWindow)
            {
                config.Portals.Add(_layoutService.CreateConfigFromWindow(appPortalWindow, id));
            }
            else if (window is Views.UrlPortalWindow urlPortalWindow)
            {
                config.Portals.Add(_layoutService.CreateConfigFromWindow(urlPortalWindow, id));
            }
        }

        return config;
    }

    /// <summary>
    /// Loads a configuration (for loading profiles)
    /// </summary>
    public void LoadConfig(DeskGridConfig config)
    {
        // Close all existing portals
        foreach (var window in _PortalWindows.ToList())
        {
            DetachWindow(window);
            window.Close();
        }

        // Load new portals from config
        _PortalsVisible = config.PortalsVisible;
        foreach (var portalConfig in config.Portals)
        {
            // Create portal and preserve its ID from config
            var window = CreatePortalFromConfig(portalConfig, preserveId: true);
            if (window != null)
            {
                // ID was already set in CreatePortalFromConfig before AttachWindow
            }
        }

        System.Diagnostics.Debug.WriteLine($"[DesktopManager] Loaded profile with {config.Portals.Count} portals");
    }

    /// <summary>
    /// Loads and restores saved portals from disk
    /// </summary>
    public void LoadSavedPortals()
    {
        System.Diagnostics.Debug.WriteLine("[DesktopManager] === LOADING SAVED PORTALS ===");
        _isLoadingPortals = true; // Prevent saves during load
        
        try
        {
            var config = _layoutService.LoadConfig();
            _loadedPortalCount = config.Portals.Count; // Track for data loss protection
            System.Diagnostics.Debug.WriteLine($"[DesktopManager] Config loaded: {config.Portals.Count} portals, portalsVisible: {config.PortalsVisible}");
            _PortalsVisible = config.PortalsVisible;

            int successCount = 0;
            int failCount = 0;
            foreach (var portalConfig in config.Portals)
            {
                System.Diagnostics.Debug.WriteLine($"[DesktopManager] Creating portal '{portalConfig.Title}' (type={portalConfig.PortalType})...");
                // Pass preserveId:true to set the ID from config BEFORE AttachWindow
                var window = CreatePortalFromConfig(portalConfig, preserveId: true);
                if (window != null)
                {
                    // ID already set in CreatePortalFromConfig
                    successCount++;
                    System.Diagnostics.Debug.WriteLine($"[DesktopManager] Portal '{portalConfig.Title}' created successfully");
                }
                else
                {
                    failCount++;
                    System.Diagnostics.Debug.WriteLine($"[DesktopManager] Portal '{portalConfig.Title}' FAILED to create!");
                }
            }

            System.Diagnostics.Debug.WriteLine($"[DesktopManager] Portal creation complete: {successCount} succeeded, {failCount} failed");
            System.Diagnostics.Debug.WriteLine($"[DesktopManager] _PortalWindows count: {_PortalWindows.Count}");
            System.Diagnostics.Debug.WriteLine($"[DesktopManager] _portalIds count: {_portalIds.Count}");

            // Apply visibility state to all loaded portals
            // (portals are created visible, so hide them if needed)
            if (!_PortalsVisible)
            {
                System.Diagnostics.Debug.WriteLine("[DesktopManager] Hiding portals (portalsVisible was false)...");
                foreach (var window in _PortalWindows)
                {
                    window.Visibility = Visibility.Collapsed;
                }
                System.Diagnostics.Debug.WriteLine("[DesktopManager] Portals hidden");
            }

            System.Diagnostics.Debug.WriteLine($"[DesktopManager] === LOAD COMPLETE: {_PortalWindows.Count} portals in memory, visible: {_PortalsVisible} ===");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[DesktopManager] LoadSavedPortals EXCEPTION: {ex.Message}");
            System.Diagnostics.Debug.WriteLine($"[DesktopManager] Stack trace: {ex.StackTrace}");
        }
        finally
        {
            _isLoadingPortals = false; // Allow saves again
        }
    }

    /// <summary>
    /// Creates a portal from a saved configuration
    /// </summary>
    /// <param name="config">The portal configuration</param>
    /// <param name="preserveId">If true, set the portal ID from config BEFORE AttachWindow to preserve it</param>
    private Window? CreatePortalFromConfig(PortalConfig config, bool preserveId = false)
    {
        try
        {
            if (config.PortalType == "folder")
            {
                var portal = new Views.PortalWindow
                {
                    Title = config.Title,
                    Left = config.X,
                    Top = config.Y,
                    Width = config.Width,
                    Height = config.Height
                };

                if (!string.IsNullOrEmpty(config.FolderPath))
                {
                    portal.SetFolder(config.FolderPath);
                }

                if (portal.DataContext is ViewModels.PortalViewModel vm)
                {
                    vm.Title = config.Title; // CRITICAL: Set title in ViewModel for persistence
                    vm.IsRolledUp = config.IsRolledUp;
                    vm.HeaderColor = config.HeaderColor;
                    vm.BackgroundColor = config.BackgroundColor;
                    vm.TitleAlignment = config.TitleAlignment ?? "Left";
                    if (Enum.TryParse<ViewModels.SortMode>(config.SortMode, out var sortMode))
                    {
                        vm.CurrentSortMode = sortMode;
                    }
                }

                // Apply colors and alignment to the UI after portal is loaded
                portal.Loaded += (s, e) =>
                {
                    try
                    {
                        // Apply header color
                        if (!string.IsNullOrEmpty(config.HeaderColor))
                        {
                            var color = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(config.HeaderColor);
                            portal.HeaderBorder.Background = new System.Windows.Media.SolidColorBrush(color);
                        }
                        
                        // Apply background color
                        if (!string.IsNullOrEmpty(config.BackgroundColor))
                        {
                            var bgColor = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(config.BackgroundColor);
                            portal.MainBorder.Background = new System.Windows.Media.SolidColorBrush(bgColor);
                        }
                        
                        // Apply title alignment
                        if (!string.IsNullOrEmpty(config.TitleAlignment))
                        {
                            var alignment = config.TitleAlignment switch
                            {
                                "Center" => System.Windows.HorizontalAlignment.Center,
                                "Right" => System.Windows.HorizontalAlignment.Right,
                                _ => System.Windows.HorizontalAlignment.Left
                            };
                            portal.TitleText.HorizontalAlignment = alignment;
                            portal.TitleText.Margin = alignment switch
                            {
                                System.Windows.HorizontalAlignment.Center => new Thickness(0),
                                System.Windows.HorizontalAlignment.Right => new Thickness(0, 0, 8, 0),
                                _ => new Thickness(8, 0, 0, 0)
                            };
                        }
                    }
                    catch { }
                };

                portal.Show();
                
                // CRITICAL: Set portal ID BEFORE AttachWindow to preserve the ID from config
                // Otherwise AttachWindow will assign a new GUID, breaking persistence
                if (preserveId && !string.IsNullOrEmpty(config.Id))
                {
                    _portalIds[portal] = config.Id;
                    System.Diagnostics.Debug.WriteLine($"[DesktopManager] Pre-assigned ID from config: {config.Id}");
                }
                
                var attached = AttachWindow(portal);
                if (!attached)
                {
                    // Still track the portal even if attachment failed, so it gets saved
                    System.Diagnostics.Debug.WriteLine($"[DesktopManager] AttachWindow failed for portal '{config.Title}', but still tracking for persistence");
                    _PortalWindows.Add(portal);
                }
                return portal;
            }
            else if (config.PortalType == "app" && !string.IsNullOrEmpty(config.ExecutablePath))
            {
                var portal = new Views.AppPortalWindow(config.ExecutablePath, config.Title)
                {
                    Left = config.X,
                    Top = config.Y,
                    Width = config.Width,
                    Height = config.Height
                };

                portal.Show();
                
                // CRITICAL: Set portal ID BEFORE AttachWindow to preserve the ID from config
                if (preserveId && !string.IsNullOrEmpty(config.Id))
                {
                    _portalIds[portal] = config.Id;
                    System.Diagnostics.Debug.WriteLine($"[DesktopManager] Pre-assigned app portal ID from config: {config.Id}");
                }
                
                var attached = AttachWindow(portal);
                if (!attached)
                {
                    // Still track the portal even if attachment failed, so it gets saved
                    System.Diagnostics.Debug.WriteLine($"[DesktopManager] AttachWindow failed for app portal '{config.Title}', but still tracking for persistence");
                    _PortalWindows.Add(portal);
                }
                return portal;
            }
            else if (config.PortalType == "url")
            {
                var portal = new Views.UrlPortalWindow
                {
                    Title = config.Title,
                    Left = config.X,
                    Top = config.Y,
                    Width = config.Width,
                    Height = config.Height
                };

                // Restore bookmarks
                if (config.Bookmarks != null)
                {
                    foreach (var bookmarkConfig in config.Bookmarks)
                    {
                        portal.Bookmarks.Add(new Models.BookmarkItem
                        {
                            Title = bookmarkConfig.Title,
                            Url = bookmarkConfig.Url
                        });
                    }
                }

                portal.Show();
                
                if (preserveId && !string.IsNullOrEmpty(config.Id))
                {
                    _portalIds[portal] = config.Id;
                    System.Diagnostics.Debug.WriteLine($"[DesktopManager] Pre-assigned URL portal ID from config: {config.Id}");
                }
                
                var attached = AttachWindow(portal);
                if (!attached)
                {
                    System.Diagnostics.Debug.WriteLine($"[DesktopManager] AttachWindow failed for URL portal '{config.Title}', but still tracking for persistence");
                    _PortalWindows.Add(portal);
                }
                return portal;
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[DesktopManager] CreatePortalFromConfig failed: {ex.Message}");
        }

        return null;
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        _zOrderTimer?.Stop();
        _zOrderTimer = null;

        foreach (var window in _PortalWindows.ToList())
        {
            window.Close();
        }
        _PortalWindows.Clear();

        GC.SuppressFinalize(this);
    }

    #region P/Invoke

    private const int GWL_STYLE = -16;
    private const int GWL_EXSTYLE = -20;
    private const int WS_POPUP = unchecked((int)0x80000000);
    private const int WS_CHILD = 0x40000000;
    private const int WS_EX_TOOLWINDOW = 0x00000080;
    private const int WS_EX_NOACTIVATE = 0x08000000;
    
    private static readonly IntPtr HWND_BOTTOM = new IntPtr(1);
    private const uint SWP_NOMOVE = 0x0002;
    private const uint SWP_NOSIZE = 0x0001;
    private const uint SWP_NOACTIVATE = 0x0010;
    private const uint SWP_NOZORDER = 0x0004;
    private const uint SWP_SHOWWINDOW = 0x0040;

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT
    {
        public int X;
        public int Y;
    }

    private delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    private static extern IntPtr FindWindow(string lpClassName, string? lpWindowName);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    private static extern IntPtr FindWindowEx(IntPtr hWndParent, IntPtr hWndChildAfter, string lpszClass, string? lpszWindow);

    [DllImport("user32.dll")]
    private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool ScreenToClient(IntPtr hWnd, ref POINT lpPoint);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    private const int SW_SHOW = 5;

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

    [DllImport("user32.dll", EntryPoint = "GetWindowLong")]
    private static extern int GetWindowLong32(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtr")]
    private static extern IntPtr GetWindowLong64(IntPtr hWnd, int nIndex);

    private static int GetWindowLong(IntPtr hWnd, int nIndex)
    {
        return IntPtr.Size == 8 
            ? (int)GetWindowLong64(hWnd, nIndex) 
            : GetWindowLong32(hWnd, nIndex);
    }

    [DllImport("user32.dll", EntryPoint = "SetWindowLong")]
    private static extern int SetWindowLong32(IntPtr hWnd, int nIndex, int dwNewLong);

    [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr")]
    private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

    private static void SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong)
    {
        if (IntPtr.Size == 8)
            SetWindowLongPtr64(hWnd, nIndex, (IntPtr)dwNewLong);
        else
            SetWindowLong32(hWnd, nIndex, dwNewLong);
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left, Top, Right, Bottom;
    }

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    #endregion
}

