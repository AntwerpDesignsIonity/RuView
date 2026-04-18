using CommunityToolkit.Mvvm.ComponentModel;
using AEDI.Sdk.Models;
using AEDI.Sdk.WebSockets;

namespace AEDI.Desktop.ViewModels;

/// <summary>
/// Vital signs monitoring: heart rate, breathing rate, confidence, signal quality with history.
/// </summary>
public partial class VitalsViewModel : ObservableObject, IDisposable
{
    private readonly SensingWebSocketService _ws;
    private const int HistorySize = 120;

    public VitalsViewModel(SensingWebSocketService ws)
    {
        _ws = ws;
        _ws.OnUpdate += HandleUpdate;
    }

    [ObservableProperty] private double _heartRate;
    [ObservableProperty] private double _breathingRate;
    [ObservableProperty] private double _heartbeatConfidence;
    [ObservableProperty] private double _breathingConfidence;
    [ObservableProperty] private double _signalQuality;

    // Ring buffers for chart history
    [ObservableProperty] private List<double> _heartRateHistory = [];
    [ObservableProperty] private List<double> _breathingRateHistory = [];
    [ObservableProperty] private List<double> _qualityHistory = [];

    [ObservableProperty] private bool _presenceDetected;
    [ObservableProperty] private string _motionLevel = "absent";

    private readonly List<double> _hrBuf = [];
    private readonly List<double> _brBuf = [];
    private readonly List<double> _sqBuf = [];

    private void HandleUpdate(SensingUpdate u)
    {
        MainThread.BeginInvokeOnMainThread(() =>
        {
            if (u.Classification is not null)
            {
                PresenceDetected = u.Classification.Presence;
                MotionLevel = u.Classification.MotionLevel;
            }

            if (u.VitalSigns is null) return;

            HeartRate = u.VitalSigns.HeartRateBpm;
            BreathingRate = u.VitalSigns.BreathingRateBpm;
            HeartbeatConfidence = u.VitalSigns.HeartbeatConfidence;
            BreathingConfidence = u.VitalSigns.BreathingConfidence;
            SignalQuality = u.VitalSigns.SignalQuality;

            Append(_hrBuf, u.VitalSigns.HeartRateBpm);
            Append(_brBuf, u.VitalSigns.BreathingRateBpm);
            Append(_sqBuf, u.VitalSigns.SignalQuality);

            HeartRateHistory = [.. _hrBuf];
            BreathingRateHistory = [.. _brBuf];
            QualityHistory = [.. _sqBuf];
        });
    }

    private static void Append(List<double> buf, double val)
    {
        buf.Add(val);
        if (buf.Count > HistorySize)
            buf.RemoveAt(0);
    }

    public void Dispose() => _ws.OnUpdate -= HandleUpdate;
}
