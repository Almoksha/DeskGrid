using System.Drawing;
using System.Windows;
using DeskGrid.Core;
using DeskGrid.Services;
using DeskGrid.Views;

namespace DeskGrid;

/// <summary>
/// DeskGrid Application - Desktop Fencing for Windows
/// </summary>
public partial class App : Application
{
    private static Mutex? _mutex;
    private DesktopManager? _desktopManager;
    private InputListener? _inputListener;
    private System.Windows.Forms.NotifyIcon? _trayIcon;
    private SelectionRectangle? _selectionRectangle;
    private int _PortalCounter = 0;

    protected override void OnStartup(StartupEventArgs e)
    {
        System.Diagnostics.Debug.WriteLine("===== DESKGRID STARTUP =====");
        System.Diagnostics.Debug.WriteLine($"[App] OnStartup started at {DateTime.Now}");
        
        // Single instance check
        const string mutexName = "DeskGrid_SingleInstance_Mutex";
        _mutex = new Mutex(true, mutexName, out bool createdNew);

        if (!createdNew)
        {
            System.Diagnostics.Debug.WriteLine("[App] Another instance is running, exiting");
            MessageBox.Show("DeskGrid is already running.", "DeskGrid", 
                MessageBoxButton.OK, MessageBoxImage.Information);
            Shutdown();
            return;
        }

        base.OnStartup(e);

        // Initialize system tray icon
        System.Diagnostics.Debug.WriteLine("[App] Initializing tray icon...");
        InitializeTrayIcon();

        // Initialize desktop integration
        System.Diagnostics.Debug.WriteLine("[App] Creating DesktopManager...");
        _desktopManager = new DesktopManager();
        _inputListener = new InputListener();

        // Attach to desktop on startup
        System.Diagnostics.Debug.WriteLine("[App] Attaching to desktop...");
        if (!_desktopManager.AttachToDesktop())
        {
            System.Diagnostics.Debug.WriteLine("[App] FAILED to attach to desktop!");
            MessageBox.Show("Failed to attach to desktop. Please restart Explorer and try again.\n\nThe application will now exit.", 
                "DeskGrid Error", MessageBoxButton.OK, MessageBoxImage.Error);
            // Exit the app to avoid running in a broken state that could wipe the config
            Shutdown();
            return;
        }
        
        // Load saved portal layout from previous session
        System.Diagnostics.Debug.WriteLine("[App] Desktop attached, loading saved portals...");
        _desktopManager.LoadSavedPortals();
        System.Diagnostics.Debug.WriteLine("[App] LoadSavedPortals completed");

        // Start input listener for double-click hide and drag-to-create
        System.Diagnostics.Debug.WriteLine("[App] Starting input listener...");
        _inputListener.Start();
        _inputListener.DesktopDoubleClick += OnDesktopDoubleClick;
        _inputListener.PortalDrawStart += OnPortalDrawStart;
        _inputListener.PortalDrawing += OnPortalDrawing;
        _inputListener.PortalDrawEnd += OnPortalDrawEnd;
        
        System.Diagnostics.Debug.WriteLine("===== DESKGRID STARTUP COMPLETE =====");
    }

    private void InitializeTrayIcon()
    {
        _trayIcon = new System.Windows.Forms.NotifyIcon
        {
            Text = "DeskGrid",
            Visible = true
        };

        // Load icon from file
        var iconPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "src", "icons", "app_icon.ico");
        if (System.IO.File.Exists(iconPath))
        {
            _trayIcon.Icon = new Icon(iconPath);
        }
        else
        {
            // Fallback: Create a simple icon programmatically
            using (var bitmap = new Bitmap(32, 32))
            using (var g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.FromArgb(124, 58, 237)); // Purple accent color
                using var font = new Font("Segoe UI", 16, System.Drawing.FontStyle.Bold);
                using var brush = new SolidBrush(Color.White);
                g.DrawString("D", font, brush, 6, 4);
                var iconHandle = bitmap.GetHicon();
                _trayIcon.Icon = Icon.FromHandle(iconHandle);
            }
        }

        // Create context menu
        var contextMenu = new System.Windows.Forms.ContextMenuStrip();
        
        var showItem = new System.Windows.Forms.ToolStripMenuItem("Show DeskGrid");
        showItem.Click += (s, e) => 
        {
            if (MainWindow != null)
            {
                MainWindow.Show();
                MainWindow.WindowState = WindowState.Normal;
                MainWindow.Activate();
            }
        };
        contextMenu.Items.Add(showItem);

        var toggleItem = new System.Windows.Forms.ToolStripMenuItem("Toggle Portals");
        toggleItem.Click += (s, e) => _desktopManager?.TogglePortalsVisibility();
        contextMenu.Items.Add(toggleItem);

        contextMenu.Items.Add(new System.Windows.Forms.ToolStripSeparator());

        var exitItem = new System.Windows.Forms.ToolStripMenuItem("Exit");
        exitItem.Click += (s, e) => 
        {
            // IMPORTANT: Save layout and mark shutdown BEFORE Application.Shutdown
            // This prevents window close handlers from wiping the saved config
            System.Diagnostics.Debug.WriteLine("[App] Exit clicked - saving before shutdown...");
            _desktopManager?.SaveLayout();
            _desktopManager?.BeginShutdown();
            Shutdown();
        };
        contextMenu.Items.Add(exitItem);

        _trayIcon.ContextMenuStrip = contextMenu;
        _trayIcon.DoubleClick += (s, e) =>
        {
            if (MainWindow != null)
            {
                MainWindow.Show();
                MainWindow.WindowState = WindowState.Normal;
                MainWindow.Activate();
            }
        };
    }

    protected override void OnExit(ExitEventArgs e)
    {
        System.Diagnostics.Debug.WriteLine("===== DESKGRID EXIT (OnExit) =====");
        System.Diagnostics.Debug.WriteLine($"[App] OnExit started at {DateTime.Now}");
        System.Diagnostics.Debug.WriteLine("[App] Note: Layout was already saved before Shutdown() was called");
        
        try
        {
            System.Diagnostics.Debug.WriteLine("[App] OnExit - Starting cleanup...");
            _selectionRectangle?.Close();
            _trayIcon?.Dispose();
            _inputListener?.Stop();
            _inputListener?.Dispose();
            System.Diagnostics.Debug.WriteLine("[App] OnExit - About to dispose DesktopManager...");
            _desktopManager?.Dispose();
            System.Diagnostics.Debug.WriteLine("[App] OnExit - DesktopManager disposed");
            _mutex?.ReleaseMutex();
            _mutex?.Dispose();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[App] OnExit - Cleanup failed: {ex.Message}");
        }

        System.Diagnostics.Debug.WriteLine("===== DESKGRID EXIT COMPLETE =====");
        base.OnExit(e);
    }

    private void OnDesktopDoubleClick(object? sender, EventArgs e)
    {
        // Toggle visibility of all Portal windows
        // Must dispatch to UI thread since mouse hook runs on background thread
        System.Diagnostics.Debug.WriteLine("[App] OnDesktopDoubleClick - toggling visibility");
        Dispatcher.Invoke(() => _desktopManager?.TogglePortalsVisibility());
    }

    private void OnPortalDrawStart(object? sender, PortalDrawEventArgs e)
    {
        System.Diagnostics.Debug.WriteLine($"[App] OnPortalDrawStart: {e.Bounds}");
        Dispatcher.Invoke(() =>
        {
            // Create and show the selection rectangle
            _selectionRectangle = new SelectionRectangle();
            _selectionRectangle.UpdateBounds(e.Bounds);
            _selectionRectangle.Show();
        });
    }

    private void OnPortalDrawing(object? sender, PortalDrawEventArgs e)
    {
        Dispatcher.Invoke(() =>
        {
            // Update the selection rectangle as user drags
            _selectionRectangle?.UpdateBounds(e.Bounds);
        });
    }

    private void OnPortalDrawEnd(object? sender, PortalDrawEventArgs e)
    {
        System.Diagnostics.Debug.WriteLine($"[App] OnPortalDrawEnd: {e.Bounds}");
        Dispatcher.Invoke(() =>
        {
            // Hide and dispose the selection rectangle
            _selectionRectangle?.Close();
            _selectionRectangle = null;

            // Create the Portal if the bounds are large enough
            if (e.Bounds.Width >= 50 && e.Bounds.Height >= 50)
            {
                _PortalCounter++;
                System.Diagnostics.Debug.WriteLine($"[App] Creating Portal {_PortalCounter} at {e.Bounds}");
                var Portal = _desktopManager?.CreateEmptyPortal(e.Bounds, $"Portal {_PortalCounter}");
                System.Diagnostics.Debug.WriteLine($"[App] Portal created: {Portal != null}");
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"[App] Bounds too small: {e.Bounds.Width}x{e.Bounds.Height}");
            }
        });
    }

    public DesktopManager? DesktopManager => _desktopManager;
}

