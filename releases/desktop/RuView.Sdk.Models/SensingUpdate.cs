namespace RuView.Sdk.Models;

/// <summary>
/// Root message from the sensing server WebSocket (/ws/sensing).
/// </summary>
public sealed class SensingUpdate
{
    public string Type { get; set; } = "sensing_update";
    public double Timestamp { get; set; }
    public string Source { get; set; } = "simulated";
    public long Tick { get; set; }

    public List<NodeInfo> Nodes { get; set; } = [];
    public SignalFeatures? Features { get; set; }
    public Classification? Classification { get; set; }
    public SignalField? SignalField { get; set; }
    public VitalSigns? VitalSigns { get; set; }
    public List<PersonDetection> Persons { get; set; } = [];
    public int EstimatedPersons { get; set; }
}

public sealed class NodeInfo
{
    public int NodeId { get; set; }
    public double RssiDbm { get; set; }
    public double[] Position { get; set; } = [0, 0, 0];
    public double[] Amplitude { get; set; } = [];
    public int SubcarrierCount { get; set; }
}

public sealed class SignalFeatures
{
    public double MeanRssi { get; set; }
    public double Variance { get; set; }
    public double MotionBandPower { get; set; }
    public double BreathingBandPower { get; set; }
    public double DominantFreqHz { get; set; }
    public int ChangePoints { get; set; }
    public double SpectralPower { get; set; }
}

public sealed class Classification
{
    public string MotionLevel { get; set; } = "absent";
    public bool Presence { get; set; }
    public double Confidence { get; set; }
}

public sealed class SignalField
{
    public int[] GridSize { get; set; } = [20, 1, 20];
    public double[] Values { get; set; } = [];
}

public sealed class VitalSigns
{
    public double BreathingRateBpm { get; set; }
    public double HeartRateBpm { get; set; }
    public double BreathingConfidence { get; set; }
    public double HeartbeatConfidence { get; set; }
    public double SignalQuality { get; set; }
}

public sealed class PersonDetection
{
    public int Id { get; set; }
    public double Confidence { get; set; }
    public BoundingBox? Bbox { get; set; }
    public List<Keypoint> Keypoints { get; set; } = [];
    public string? Zone { get; set; }
}

public sealed class BoundingBox
{
    public double X { get; set; }
    public double Y { get; set; }
    public double Width { get; set; }
    public double Height { get; set; }
}

public sealed class Keypoint
{
    public string Name { get; set; } = "";
    public double X { get; set; }
    public double Y { get; set; }
    public double Z { get; set; }
    public double Confidence { get; set; }
}