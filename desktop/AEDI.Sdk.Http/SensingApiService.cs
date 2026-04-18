using System.Text.Json;
using AEDI.Sdk.Models;

namespace AEDI.Sdk.Http;

/// <summary>
/// HTTP client for the sensing server REST API at :3000.
/// </summary>
public sealed class SensingApiService
{
    private readonly HttpClient _http;
    private readonly JsonSerializerOptions _json = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower,
        PropertyNameCaseInsensitive = true
    };

    public SensingApiService(HttpClient http)
    {
        _http = http;
    }

    public void SetBaseUrl(string host, int port)
    {
        _http.BaseAddress = new Uri($"http://{host}:{port}");
    }

    public Task<HealthStatus?> GetHealthAsync(CancellationToken ct = default)
        => GetAsync<HealthStatus>("/health", ct);

    public Task<SensingUpdate?> GetLatestAsync(CancellationToken ct = default)
        => GetAsync<SensingUpdate>("/api/v1/sensing/latest", ct);

    public Task<VitalSigns?> GetVitalSignsAsync(CancellationToken ct = default)
        => GetAsync<VitalSigns>("/api/v1/vital-signs", ct);

    public Task<List<NodeInfo>?> GetNodesAsync(CancellationToken ct = default)
        => GetAsync<List<NodeInfo>>("/api/v1/nodes", ct);

    public Task<List<PersonDetection>?> GetPoseCurrentAsync(CancellationToken ct = default)
        => GetAsync<List<PersonDetection>>("/api/v1/pose/current", ct);

    public Task<List<ModelInfo>?> GetModelsAsync(CancellationToken ct = default)
        => GetAsync<List<ModelInfo>>("/api/v1/models", ct);

    public Task<ModelInfo?> GetActiveModelAsync(CancellationToken ct = default)
        => GetAsync<ModelInfo>("/api/v1/models/active", ct);

    public async Task<bool> LoadModelAsync(string modelId, CancellationToken ct = default)
    {
        var content = new StringContent(
            JsonSerializer.Serialize(new { id = modelId }),
            System.Text.Encoding.UTF8, "application/json");
        var resp = await _http.PostAsync("/api/v1/models/load", content, ct);
        return resp.IsSuccessStatusCode;
    }

    public async Task<bool> UnloadModelAsync(CancellationToken ct = default)
    {
        var resp = await _http.PostAsync("/api/v1/models/unload", null, ct);
        return resp.IsSuccessStatusCode;
    }

    public Task<List<RecordingInfo>?> GetRecordingsAsync(CancellationToken ct = default)
        => GetAsync<List<RecordingInfo>>("/api/v1/recording/list", ct);

    public async Task<bool> StartRecordingAsync(CancellationToken ct = default)
    {
        var resp = await _http.PostAsync("/api/v1/recording/start", null, ct);
        return resp.IsSuccessStatusCode;
    }

    public async Task<bool> StopRecordingAsync(CancellationToken ct = default)
    {
        var resp = await _http.PostAsync("/api/v1/recording/stop", null, ct);
        return resp.IsSuccessStatusCode;
    }

    public Task<TrainingStatus?> GetTrainingStatusAsync(CancellationToken ct = default)
        => GetAsync<TrainingStatus>("/api/v1/train/status", ct);

    public async Task<bool> StartTrainingAsync(CancellationToken ct = default)
    {
        var resp = await _http.PostAsync("/api/v1/train/start", null, ct);
        return resp.IsSuccessStatusCode;
    }

    public async Task<bool> StopTrainingAsync(CancellationToken ct = default)
    {
        var resp = await _http.PostAsync("/api/v1/train/stop", null, ct);
        return resp.IsSuccessStatusCode;
    }

    public async Task<bool> StartCalibrationAsync(CancellationToken ct = default)
    {
        var resp = await _http.PostAsync("/api/v1/calibration/start", null, ct);
        return resp.IsSuccessStatusCode;
    }

    public async Task<bool> StopCalibrationAsync(CancellationToken ct = default)
    {
        var resp = await _http.PostAsync("/api/v1/calibration/stop", null, ct);
        return resp.IsSuccessStatusCode;
    }

    private async Task<T?> GetAsync<T>(string path, CancellationToken ct)
    {
        try
        {
            var resp = await _http.GetAsync(path, ct);
            if (!resp.IsSuccessStatusCode) return default;
            var stream = await resp.Content.ReadAsStreamAsync(ct);
            return await JsonSerializer.DeserializeAsync<T>(stream, _json, ct);
        }
        catch
        {
            return default;
        }
    }
}