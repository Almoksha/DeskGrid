namespace DeskGrid.Models;

/// <summary>
/// Represents a rule for automatically sorting files into portals
/// </summary>
public class PortalRule
{
    /// <summary>
    /// Unique identifier for this rule
    /// </summary>
    public string Id { get; set; } = Guid.NewGuid().ToString();

    /// <summary>
    /// Human-readable name for the rule
    /// </summary>
    public string Name { get; set; } = "New Rule";

    /// <summary>
    /// Pattern to match (glob format, e.g., "*.pdf", "Report*")
    /// </summary>
    public string Pattern { get; set; } = "*";

    /// <summary>
    /// Whether the rule is active
    /// </summary>
    public bool IsEnabled { get; set; } = true;

    /// <summary>
    /// Target portal ID where matching files should go
    /// </summary>
    public string? TargetPortalId { get; set; }

    /// <summary>
    /// Action to perform: Move or Copy
    /// </summary>
    public RuleAction Action { get; set; } = RuleAction.Move;

    /// <summary>
    /// Check if a filename matches this rule
    /// </summary>
    public bool Matches(string filename)
    {
        if (!IsEnabled || string.IsNullOrEmpty(Pattern))
            return false;

        // Simple glob matching
        var pattern = Pattern.Replace(".", "\\.").Replace("*", ".*").Replace("?", ".");
        try
        {
            return System.Text.RegularExpressions.Regex.IsMatch(
                filename, 
                $"^{pattern}$", 
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        }
        catch
        {
            return false;
        }
    }
}

public enum RuleAction
{
    Move,
    Copy
}
