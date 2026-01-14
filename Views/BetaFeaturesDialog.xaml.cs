using System.Windows;
using System.Windows.Input;

namespace DeskGrid.Views;

public partial class BetaFeaturesDialog : Window
{
    public BetaFeaturesDialog()
    {
        InitializeComponent();
    }

    private void Header_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        DragMove();
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void EmbedApp_Click(object sender, RoutedEventArgs e)
    {
        System.Diagnostics.Debug.WriteLine("[BetaFeatures] EmbedApp_Click called");
        
        using var dialog = new System.Windows.Forms.OpenFileDialog
        {
            Title = "Select an application to embed",
            Filter = "Executable files (*.exe)|*.exe|All files (*.*)|*.*",
            CheckFileExists = true
        };

        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            System.Diagnostics.Debug.WriteLine($"[BetaFeatures] Executable selected: {dialog.FileName}");
            
            if (System.Windows.Application.Current is App app && app.DesktopManager != null)
            {
                var appName = System.IO.Path.GetFileNameWithoutExtension(dialog.FileName);
                System.Diagnostics.Debug.WriteLine($"[BetaFeatures] Creating app Portal '{appName}'");
                
                var portal = app.DesktopManager.CreateAppPortal(dialog.FileName, appName);
                System.Diagnostics.Debug.WriteLine($"[BetaFeatures] App Portal created: {portal != null}");
                
                // Close dialog after creating
                Close();
            }
        }
    }
}
