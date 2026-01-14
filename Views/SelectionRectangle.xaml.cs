using System.Windows;

namespace DeskGrid.Views;

/// <summary>
/// Selection rectangle overlay for drag-to-create Portal
/// </summary>
public partial class SelectionRectangle : Window
{
    public SelectionRectangle()
    {
        InitializeComponent();
    }

    /// <summary>
    /// Updates the position and size of the selection rectangle
    /// </summary>
    public void UpdateBounds(System.Drawing.Rectangle bounds)
    {
        Left = bounds.X;
        Top = bounds.Y;
        Width = Math.Max(bounds.Width, 1);
        Height = Math.Max(bounds.Height, 1);
    }
}
