using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Input;
using DeskGrid.Models;
using DeskGrid.Services;

namespace DeskGrid.Views;

public partial class RulesDialog : Window
{
    private readonly RuleEngine _ruleEngine;

    public ObservableCollection<PortalRule> Rules { get; } = new();
    public bool HasNoRules => Rules.Count == 0;

    public RulesDialog()
    {
        InitializeComponent();
        DataContext = this;
        
        _ruleEngine = new RuleEngine();
        LoadRules();
        
        Rules.CollectionChanged += (s, e) => 
        {
            // Notify UI of empty state change
        };
    }

    private void LoadRules()
    {
        Rules.Clear();
        foreach (var rule in _ruleEngine.Rules)
        {
            Rules.Add(rule);
        }
    }

    private void Header_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        DragMove();
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        // Save rules before closing
        SaveAllRules();
        Close();
    }

    private void AddRule_Click(object sender, RoutedEventArgs e)
    {
        var name = Microsoft.VisualBasic.Interaction.InputBox(
            "Enter a name for this rule:", "Add Rule", "My Rule");
        
        if (string.IsNullOrWhiteSpace(name)) return;

        var pattern = Microsoft.VisualBasic.Interaction.InputBox(
            "Enter the file pattern to match:\n\nExamples:\n  *.pdf - All PDF files\n  *.jpg, *.png - Images\n  Report* - Files starting with 'Report'",
            "File Pattern", "*.pdf");
        
        if (string.IsNullOrWhiteSpace(pattern)) return;

        // Get list of available portals
        string? targetPortalId = null;
        if (System.Windows.Application.Current is App app && app.DesktopManager != null)
        {
            var portals = app.DesktopManager.GetAllPortalInfo();
            if (portals.Count > 0)
            {
                var message = "Select target portal:\n\n";
                for (int i = 0; i < portals.Count; i++)
                {
                    message += $"{i + 1}. {portals[i].Title}\n";
                }
                message += "\nEnter number (or leave empty to skip):";

                var selection = Microsoft.VisualBasic.Interaction.InputBox(message, "Target Portal", "1");
                
                if (int.TryParse(selection, out int index) && index >= 1 && index <= portals.Count)
                {
                    targetPortalId = portals[index - 1].Id;
                }
            }
        }

        var rule = new PortalRule
        {
            Name = name,
            Pattern = pattern,
            IsEnabled = true,
            TargetPortalId = targetPortalId
        };

        Rules.Add(rule);
        _ruleEngine.AddRule(rule);
    }

    private void RemoveRule_Click(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement element && element.Tag is PortalRule rule)
        {
            Rules.Remove(rule);
            _ruleEngine.RemoveRule(rule.Id);
        }
    }

    private void SaveAllRules()
    {
        _ruleEngine.SaveRules();
    }
}
