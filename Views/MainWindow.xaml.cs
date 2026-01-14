using System.Windows;
using System.Windows.Input;

namespace DeskGrid.Views;

/// <summary>
/// Main application window - control panel for DeskGrid
/// </summary>
public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        UpdatePortalsStatus();
    }

    private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ClickCount == 2)
        {
            // Double-click to toggle maximize (if we supported it)
            return;
        }
        DragMove();
    }

    private void MinimizeButton_Click(object sender, RoutedEventArgs e)
    {
        WindowState = WindowState.Minimized;
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        // Hide instead of close - app lives in system tray
        Hide();
    }

    private void CreatePortal_Click(object sender, RoutedEventArgs e)
    {
        System.Diagnostics.Debug.WriteLine("[MainWindow] CreatePortal_Click called");
        
        using var dialog = new System.Windows.Forms.FolderBrowserDialog
        {
            Description = "Select a folder for the new Portal",
            ShowNewFolderButton = true
        };

        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            System.Diagnostics.Debug.WriteLine($"[MainWindow] Folder selected: {dialog.SelectedPath}");
            
            if (System.Windows.Application.Current is App app && app.DesktopManager != null)
            {
                // Get folder name as Portal title
                var folderName = System.IO.Path.GetFileName(dialog.SelectedPath);
                if (string.IsNullOrEmpty(folderName))
                    folderName = dialog.SelectedPath;

                System.Diagnostics.Debug.WriteLine($"[MainWindow] Creating Portal '{folderName}'");
                var Portal = app.DesktopManager.CreatePortal(folderName, dialog.SelectedPath);
                System.Diagnostics.Debug.WriteLine($"[MainWindow] Portal created: {Portal != null}");
                UpdatePortalsStatus();
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("[MainWindow] ERROR: App or DesktopManager is null");
            }
        }
        else
        {
            System.Diagnostics.Debug.WriteLine("[MainWindow] Folder dialog cancelled");
        }
    }

    private void CreateEmptyPortal_Click(object sender, RoutedEventArgs e)
    {
        System.Diagnostics.Debug.WriteLine("[MainWindow] CreateEmptyPortal_Click called");
        
        var name = Microsoft.VisualBasic.Interaction.InputBox(
            "Enter a name for the empty portal:", "Create Empty Portal", "New Portal");
        
        if (!string.IsNullOrWhiteSpace(name))
        {
            if (System.Windows.Application.Current is App app && app.DesktopManager != null)
            {
                var portal = app.DesktopManager.CreateEmptyPortal(
                    new System.Drawing.Rectangle(100, 100, 280, 350), name);
                System.Diagnostics.Debug.WriteLine($"[MainWindow] Empty Portal created: {portal != null}");
                UpdatePortalsStatus();
            }
        }
    }

    private void ToggleVisibility_Click(object sender, RoutedEventArgs e)
    {
        if (System.Windows.Application.Current is App app)
        {
            app.DesktopManager?.TogglePortalsVisibility();
        }
    }

    private void EmbedApp_Click(object sender, RoutedEventArgs e)
    {
        System.Diagnostics.Debug.WriteLine("[MainWindow] EmbedApp_Click called");
        
        using var dialog = new System.Windows.Forms.OpenFileDialog
        {
            Title = "Select an application to embed",
            Filter = "Executable files (*.exe)|*.exe|All files (*.*)|*.*",
            CheckFileExists = true
        };

        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            System.Diagnostics.Debug.WriteLine($"[MainWindow] Executable selected: {dialog.FileName}");
            
            if (System.Windows.Application.Current is App app && app.DesktopManager != null)
            {
                var appName = System.IO.Path.GetFileNameWithoutExtension(dialog.FileName);
                System.Diagnostics.Debug.WriteLine($"[MainWindow] Creating app Portal '{appName}'");
                
                var Portal = app.DesktopManager.CreateAppPortal(dialog.FileName, appName);
                System.Diagnostics.Debug.WriteLine($"[MainWindow] App Portal created: {Portal != null}");
                UpdatePortalsStatus();
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("[MainWindow] ERROR: App or DesktopManager is null");
            }
        }
        else
        {
            System.Diagnostics.Debug.WriteLine("[MainWindow] Executable dialog cancelled");
        }
    }

    private void CreateUrlPortal_Click(object sender, RoutedEventArgs e)
    {
        System.Diagnostics.Debug.WriteLine("[MainWindow] CreateUrlPortal_Click called");
        
        if (System.Windows.Application.Current is App app && app.DesktopManager != null)
        {
            var portal = app.DesktopManager.CreateUrlPortal("Bookmarks");
            System.Diagnostics.Debug.WriteLine($"[MainWindow] URL Portal created: {portal != null}");
            UpdatePortalsStatus();
        }
    }

    private void UpdatePortalsStatus()
    {
        // This would be better with proper data binding, but for simplicity:
        if (System.Windows.Application.Current is App app && app.DesktopManager != null)
        {
            // For now, just show a simple status
            var visible = app.DesktopManager.PortalsVisible;
            PortalsStatus.Text = visible 
                ? "Portals are visible. Double-click desktop to hide them."
                : "Portals are hidden. Double-click desktop to show them.";
        }
    }

    private void SaveProfile_Click(object sender, RoutedEventArgs e)
    {
        // Simple input dialog using Microsoft.VisualBasic (available in .NET)
        var profileName = Microsoft.VisualBasic.Interaction.InputBox(
            "Enter a name for this profile:",
            "Save Profile",
            "My Profile");

        if (!string.IsNullOrWhiteSpace(profileName))
        {
            if (System.Windows.Application.Current is App app && app.DesktopManager != null)
            {
                var profileService = new Services.ProfileService();
                var config = app.DesktopManager.GetCurrentConfig();
                profileService.SaveProfile(profileName, config);
                MessageBox.Show($"Profile '{profileName}' saved!", "Profile Saved", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
    }

    private void LoadProfile_Click(object sender, RoutedEventArgs e)
    {
        var profileService = new Services.ProfileService();
        var profiles = profileService.GetProfiles().ToList();

        if (profiles.Count == 0)
        {
            MessageBox.Show("No saved profiles found. Save a profile first!", "No Profiles", 
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        // Simple selection using numbered list
        var message = "Available profiles:\n\n";
        for (int i = 0; i < profiles.Count; i++)
        {
            message += $"{i + 1}. {profiles[i]}\n";
        }
        message += "\nEnter the number of the profile to load:";

        var selection = Microsoft.VisualBasic.Interaction.InputBox(message, "Load Profile", "1");
        
        if (int.TryParse(selection, out int index) && index >= 1 && index <= profiles.Count)
        {
            var profileName = profiles[index - 1];
            var config = profileService.LoadProfile(profileName);
            
            if (config != null && System.Windows.Application.Current is App app && app.DesktopManager != null)
            {
                app.DesktopManager.LoadConfig(config);
                MessageBox.Show($"Profile '{profileName}' loaded!", "Profile Loaded", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
                UpdatePortalsStatus();
            }
        }
    }

    private void ManageRules_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new RulesDialog();
        dialog.ShowDialog();
    }

    private void BetaFeatures_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new BetaFeaturesDialog();
        dialog.ShowDialog();
    }

    protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
    {
        // Prevent closing, just hide
        e.Cancel = true;
        Hide();
    }

    private void Hyperlink_RequestNavigate(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
    {
        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
        {
            FileName = e.Uri.AbsoluteUri,
            UseShellExecute = true
        });
        e.Handled = true;
    }
}
