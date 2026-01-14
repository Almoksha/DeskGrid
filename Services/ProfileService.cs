using System.IO;
using System.Text.Json;
using DeskGrid.Models;

namespace DeskGrid.Services;

/// <summary>
/// Manages multiple named profile layouts
/// </summary>
public class ProfileService
{
    private readonly string _profilesDir;
    private readonly PortalLayoutService _layoutService;

    private static readonly JsonSerializerOptions _jsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    public ProfileService()
    {
        var appDir = AppDomain.CurrentDomain.BaseDirectory;
        _profilesDir = Path.Combine(appDir, "config", "profiles");
        Directory.CreateDirectory(_profilesDir);
        _layoutService = new PortalLayoutService();
    }

    /// <summary>
    /// Gets all available profile names
    /// </summary>
    public IEnumerable<string> GetProfiles()
    {
        if (!Directory.Exists(_profilesDir))
            return Enumerable.Empty<string>();

        return Directory.GetFiles(_profilesDir, "*.json")
            .Select(Path.GetFileNameWithoutExtension)
            .Where(n => n != null)
            .Cast<string>()
            .OrderBy(n => n);
    }

    /// <summary>
    /// Saves current layout as a named profile
    /// </summary>
    public void SaveProfile(string name, DeskGridConfig config)
    {
        try
        {
            var path = GetProfilePath(name);
            var json = JsonSerializer.Serialize(config, _jsonOptions);
            File.WriteAllText(path, json);
            System.Diagnostics.Debug.WriteLine($"[ProfileService] Saved profile: {name}");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[ProfileService] Save failed: {ex.Message}");
        }
    }

    /// <summary>
    /// Loads a named profile
    /// </summary>
    public DeskGridConfig? LoadProfile(string name)
    {
        try
        {
            var path = GetProfilePath(name);
            if (!File.Exists(path))
                return null;

            var json = File.ReadAllText(path);
            return JsonSerializer.Deserialize<DeskGridConfig>(json, _jsonOptions);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[ProfileService] Load failed: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Deletes a named profile
    /// </summary>
    public bool DeleteProfile(string name)
    {
        try
        {
            var path = GetProfilePath(name);
            if (File.Exists(path))
            {
                File.Delete(path);
                return true;
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[ProfileService] Delete failed: {ex.Message}");
        }
        return false;
    }

    private string GetProfilePath(string name)
    {
        // Sanitize name
        var safeName = string.Join("_", name.Split(Path.GetInvalidFileNameChars()));
        return Path.Combine(_profilesDir, $"{safeName}.json");
    }
}
