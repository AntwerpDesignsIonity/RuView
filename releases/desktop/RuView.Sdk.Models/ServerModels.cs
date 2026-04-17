namespace RuView.Sdk.Models;

public sealed class HealthStatus
{
    public string Status { get; set; } = "unknown";
    public string Version { get; set; } = "";
    public double Uptime { get; set; }
    public string Source { get; set; } = "";
    public int NodeCount { get; set; }
}

public sealed class ModelInfo
{
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
    public long SizeBytes { get; set; }
    public string Status { get; set; } = "unloaded";
}

public sealed class RecordingInfo
{
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
    public int FrameCount { get; set; }
    public double DurationSecs { get; set; }
    public string Status { get; set; } = "stopped";
}

public sealed class TrainingStatus
{
    public bool Active { get; set; }
    public int Epoch { get; set; }
    public int TotalEpochs { get; set; }
    public double Loss { get; set; }
    public double Accuracy { get; set; }
}

public sealed class ServerSettings
{
    public string ServerHost { get; set; } = "192.168.1.1";
    public int HttpPort { get; set; } = 3000;
    public int WsPort { get; set; } = 3001;
    public bool AutoReconnect { get; set; } = true;
    public int ReconnectIntervalMs { get; set; } = 3000;
    public string Theme { get; set; } = "dark";
}