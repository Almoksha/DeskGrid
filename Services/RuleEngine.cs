using System.IO;
using System.Text.Json;
using DeskGrid.Models;

namespace DeskGrid.Services;

/// <summary>
/// Processes rules to automatically sort files into portals
/// </summary>
public class RuleEngine
{
    private readonly string _rulesPath;
    private List<PortalRule> _rules = new();

    private static readonly JsonSerializerOptions _jsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    public RuleEngine()
    {
        var appDir = AppDomain.CurrentDomain.BaseDirectory;
        var configDir = Path.Combine(appDir, "config");
        Directory.CreateDirectory(configDir);
        _rulesPath = Path.Combine(configDir, "rules.json");
        LoadRules();
    }

    /// <summary>
    /// All configured rules
    /// </summary>
    public IReadOnlyList<PortalRule> Rules => _rules.AsReadOnly();

    /// <summary>
    /// Loads rules from disk
    /// </summary>
    public void LoadRules()
    {
        try
        {
            if (File.Exists(_rulesPath))
            {
                var json = File.ReadAllText(_rulesPath);
                _rules = JsonSerializer.Deserialize<List<PortalRule>>(json, _jsonOptions) ?? new();
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[RuleEngine] Load failed: {ex.Message}");
            _rules = new();
        }
    }

    /// <summary>
    /// Saves rules to disk
    /// </summary>
    public void SaveRules()
    {
        try
        {
            var json = JsonSerializer.Serialize(_rules, _jsonOptions);
            File.WriteAllText(_rulesPath, json);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[RuleEngine] Save failed: {ex.Message}");
        }
    }

    /// <summary>
    /// Adds a new rule
    /// </summary>
    public void AddRule(PortalRule rule)
    {
        _rules.Add(rule);
        SaveRules();
    }

    /// <summary>
    /// Removes a rule
    /// </summary>
    public bool RemoveRule(string ruleId)
    {
        var removed = _rules.RemoveAll(r => r.Id == ruleId) > 0;
        if (removed) SaveRules();
        return removed;
    }

    /// <summary>
    /// Finds the first matching rule for a filename
    /// </summary>
    public PortalRule? FindMatchingRule(string filename)
    {
        return _rules.FirstOrDefault(r => r.Matches(filename));
    }

    /// <summary>
    /// Applies rules to a file, moving/copying to target portal folder
    /// </summary>
    public bool ApplyRules(string filePath, Func<string, string?> getPortalFolderById)
    {
        var filename = Path.GetFileName(filePath);
        var rule = FindMatchingRule(filename);
        
        if (rule == null || string.IsNullOrEmpty(rule.TargetPortalId))
            return false;

        var targetFolder = getPortalFolderById(rule.TargetPortalId);
        if (string.IsNullOrEmpty(targetFolder))
            return false;

        try
        {
            var destPath = Path.Combine(targetFolder, filename);
            
            if (rule.Action == RuleAction.Copy)
                File.Copy(filePath, destPath, overwrite: false);
            else
                File.Move(filePath, destPath);

            System.Diagnostics.Debug.WriteLine($"[RuleEngine] Applied rule '{rule.Name}': {filePath} -> {destPath}");
            return true;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[RuleEngine] Apply failed: {ex.Message}");
            return false;
        }
    }
}
