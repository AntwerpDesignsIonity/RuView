using CommunityToolkit.Mvvm.ComponentModel;
using AEDI.Sdk.Models;
using AEDI.Sdk.WebSockets;

namespace AEDI.Desktop.ViewModels;

/// <summary>
/// Presence detection page: person count, motion classification, confidence, zone map.
/// </summary>
public partial class PresenceViewModel : ObservableObject, IDisposable
{
    private readonly SensingWebSocketService _ws;

    public PresenceViewModel(SensingWebSocketService ws)
    {
        _ws = ws;
        _ws.OnUpdate += HandleUpdate;
    }

    [ObservableProperty] private bool _presenceDetected;
    [ObservableProperty] private string _motionLevel = "absent";
    [ObservableProperty] private double _confidence;
    [ObservableProperty] private int _estimatedPersons;
    [ObservableProperty] private double _meanRssi;
    [ObservableProperty] private double[] _signalFieldValues = [];
    [ObservableProperty] private int _signalFieldWidth = 20;
    [ObservableProperty] private int _signalFieldHeight = 20;
    [ObservableProperty] private List<NodeInfo> _nodes = [];

    // History for trend sparkline
    [ObservableProperty] private List<double> _confidenceHistory = [];
    private readonly List<double> _confBuf = [];

    private void HandleUpdate(SensingUpdate u)
    {
        MainThread.BeginInvokeOnMainThread(() =>
        {
            EstimatedPersons = u.EstimatedPersons;
            Nodes = u.Nodes;

            if (u.Classification is not null)
            {
                PresenceDetected = u.Classification.Presence;
                MotionLevel = u.Classification.MotionLevel;
                Confidence = u.Classification.Confidence;

                _confBuf.Add(u.Classification.Confidence);
                if (_confBuf.Count > 60) _confBuf.RemoveAt(0);
                ConfidenceHistory = [.. _confBuf];
            }

            if (u.Features is not null)
                MeanRssi = u.Features.MeanRssi;

            if (u.SignalField is not null)
            {
                SignalFieldValues = u.SignalField.Values;
                if (u.SignalField.GridSize.Length >= 3)
                {
                    SignalFieldWidth = u.SignalField.GridSize[0];
                    SignalFieldHeight = u.SignalField.GridSize[2];
                }
            }
        });
    }

    public void Dispose() => _ws.OnUpdate -= HandleUpdate;
}
