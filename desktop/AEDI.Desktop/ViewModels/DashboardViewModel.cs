using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using AEDI.Desktop.Services;
using AEDI.Sdk.Http;
using AEDI.Sdk.Models;
using AEDI.Sdk.WebSockets;

namespace AEDI.Desktop.ViewModels;

/// <summary>
/// Main dashboard: connection status, presence, motion, vitals summary, signal features, node table.
/// </summary>
public partial class DashboardViewModel : ObservableObject, IDisposable
{
    private readonly SensingWebSocketService _ws;
    private readonly SensingApiService _api;
    private readonly SettingsService _settings;

    public DashboardViewModel(SensingWebSocketService ws, SensingApiService api, SettingsService settings)
    {
        _ws = ws;
        _api = api;
        _settings = settings;

        _ws.OnUpdate += HandleUpdate;
        _ws.OnConnectionChanged += connected =>
            MainThread.BeginInvokeOnMainThread(() => IsConnected = connected);
        _ws.OnError += msg =>
            MainThread.BeginInvokeOnMainThread(() => StatusMessage = msg);
    }

    // ── Connection ──────────────────────────────────────────────
    [ObservableProperty] private bool _isConnected;
    [ObservableProperty] private string _statusMessage = "Disconnected";

    // ── Presence / Motion ───────────────────────────────────────
    [ObservableProperty] private bool _presenceDetected;
    [ObservableProperty] private string _motionLevel = "absent";
    [ObservableProperty] private double _confidence;
    [ObservableProperty] private int _estimatedPersons;

    // ── Vitals summary ──────────────────────────────────────────
    [ObservableProperty] private double _heartRate;
    [ObservableProperty] private double _breathingRate;
    [ObservableProperty] private double _signalQuality;

    // ── Signal features ─────────────────────────────────────────
    [ObservableProperty] private double _meanRssi;
    [ObservableProperty] private double _variance;
    [ObservableProperty] private double _spectralPower;
    [ObservableProperty] private double _dominantFreq;
    [ObservableProperty] private double _motionBandPower;
    [ObservableProperty] private double _breathingBandPower;

    // ── Tick counter ────────────────────────────────────────────
    [ObservableProperty] private long _tick;
    [ObservableProperty] private string _source = "";

    // ── Signal field heatmap (flattened 20×20) ──────────────────
    [ObservableProperty] private double[] _signalFieldValues = [];

    // ── Active nodes ────────────────────────────────────────────
    [ObservableProperty] private List<NodeInfo> _nodes = [];

    [RelayCommand]
    private async Task ConnectAsync()
    {
        var s = _settings.Current;
        _api.SetBaseUrl(s.ServerHost, s.HttpPort);
        StatusMessage = $"Connecting to {s.ServerHost}…";
        await _ws.ConnectAsync(s.ServerHost, s.WsPort);
    }

    [RelayCommand]
    private async Task DisconnectAsync()
    {
        await _ws.DisconnectAsync();
        StatusMessage = "Disconnected";
    }

    private void HandleUpdate(SensingUpdate u)
    {
        MainThread.BeginInvokeOnMainThread(() =>
        {
            Tick = u.Tick;
            Source = u.Source;
            EstimatedPersons = u.EstimatedPersons;

            if (u.Classification is not null)
            {
                PresenceDetected = u.Classification.Presence;
                MotionLevel = u.Classification.MotionLevel;
                Confidence = u.Classification.Confidence;
            }

            if (u.VitalSigns is not null)
            {
                HeartRate = u.VitalSigns.HeartRateBpm;
                BreathingRate = u.VitalSigns.BreathingRateBpm;
                SignalQuality = u.VitalSigns.SignalQuality;
            }

            if (u.Features is not null)
            {
                MeanRssi = u.Features.MeanRssi;
                Variance = u.Features.Variance;
                SpectralPower = u.Features.SpectralPower;
                DominantFreq = u.Features.DominantFreqHz;
                MotionBandPower = u.Features.MotionBandPower;
                BreathingBandPower = u.Features.BreathingBandPower;
            }

            if (u.SignalField is not null)
                SignalFieldValues = u.SignalField.Values;

            Nodes = u.Nodes;

            StatusMessage = $"Live — tick {u.Tick}";
        });
    }

    public void Dispose()
    {
        _ws.OnUpdate -= HandleUpdate;
    }
}
