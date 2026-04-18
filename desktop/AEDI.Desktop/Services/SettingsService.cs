using System.Text.Json;
using AEDI.Sdk.Models;

namespace AEDI.Desktop.Services;

/// <summary>
/// Persists user settings to app data directory.
/// </summary>
public sealed class SettingsService
{
    private static readonly string SettingsPath = Path.Combine(
        FileSystem.AppDataDirectory, "aedi-settings.json");

    private static readonly JsonSerializerOptions Json = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    public ServerSettings Current { get; private set; } = new();

    public async Task LoadAsync()
    {
        if (!File.Exists(SettingsPath))
            return;

        try
        {
            await using var fs = File.OpenRead(SettingsPath);
            Current = await JsonSerializer.DeserializeAsync<ServerSettings>(fs, Json) ?? new();
        }
        catch
        {
            Current = new();
        }
    }

    public async Task SaveAsync()
    {
        var dir = Path.GetDirectoryName(SettingsPath);
        if (dir is not null)
            Directory.CreateDirectory(dir);

        await using var fs = File.Create(SettingsPath);
        await JsonSerializer.SerializeAsync(fs, Current, Json);
    }
}
