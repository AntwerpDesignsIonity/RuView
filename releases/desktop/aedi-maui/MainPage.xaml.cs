using System.Net.Http;
using System.Net.WebSockets;
using System.Text;
using System.Text.Json;

namespace Ionity.AediMaui;

public partial class MainPage : ContentPage
{
    private ClientWebSocket? _ws;
    private CancellationTokenSource? _cts;
    private Task? _pump;
    private readonly HttpClient _http = new() { Timeout = TimeSpan.FromSeconds(5) };

    public MainPage()
    {
        InitializeComponent();
        HubEntry.Text = Preferences.Get("hub_url", "ws://192.168.124.7:3001/ws/sensing");
    }

    protected override void OnDisappearing()
    {
        base.OnDisappearing();
        _ = DisconnectAsync();
    }

    private async void OnConnectClicked(object? sender, EventArgs e)
    {
        if (_ws?.State == WebSocketState.Open)
        {
            await DisconnectAsync();
            return;
        }
        var url = (HubEntry.Text ?? "").Trim();
        if (string.IsNullOrWhiteSpace(url))
        {
            Log("hub URL empty");
            return;
        }
        Preferences.Set("hub_url", url);
        await ConnectAsync(url);
    }

    private async Task ConnectAsync(string url)
    {
        try
        {
            _cts = new CancellationTokenSource();
            _ws = new ClientWebSocket();
            Log($"connecting {url} …");
            await _ws.ConnectAsync(new Uri(url), _cts.Token);
            SetStatus("ONLINE", true);
            ConnectBtn.Text = "Disconnect";
            _pump = Task.Run(() => PumpAsync(_cts.Token));
        }
        catch (Exception ex)
        {
            Log($"connect failed: {ex.Message}");
            SetStatus("OFFLINE", false);
        }
    }

    private async Task DisconnectAsync()
    {
        try
        {
            _cts?.Cancel();
            if (_ws?.State == WebSocketState.Open)
            {
                await _ws.CloseAsync(WebSocketCloseStatus.NormalClosure, "bye", CancellationToken.None);
            }
        }
        catch { /* ignore */ }
        finally
        {
            _ws?.Dispose();
            _ws = null;
            SetStatus("OFFLINE", false);
            ConnectBtn.Text = "Connect";
        }
    }

    private async Task PumpAsync(CancellationToken ct)
    {
        if (_ws is null) return;
        var buf = new byte[1 << 16];
        var sb = new StringBuilder();
        try
        {
            while (!ct.IsCancellationRequested && _ws.State == WebSocketState.Open)
            {
                sb.Clear();
                WebSocketReceiveResult res;
                do
                {
                    res = await _ws.ReceiveAsync(new ArraySegment<byte>(buf), ct);
                    if (res.MessageType == WebSocketMessageType.Close)
                    {
                        MainThread.BeginInvokeOnMainThread(() => SetStatus("OFFLINE", false));
                        return;
                    }
                    sb.Append(Encoding.UTF8.GetString(buf, 0, res.Count));
                } while (!res.EndOfMessage);

                var msg = sb.ToString();
                try
                {
                    using var doc = JsonDocument.Parse(msg);
                    var root = doc.RootElement;
                    MainThread.BeginInvokeOnMainThread(() => ApplyUpdate(root));
                }
                catch (Exception ex)
                {
                    MainThread.BeginInvokeOnMainThread(() => Log($"parse: {ex.Message}"));
                }
            }
        }
        catch (OperationCanceledException) { }
        catch (Exception ex)
        {
            MainThread.BeginInvokeOnMainThread(() => { Log($"pump: {ex.Message}"); SetStatus("OFFLINE", false); });
        }
    }

    private void ApplyUpdate(JsonElement root)
    {
        if (root.TryGetProperty("tick", out var tick))
            TickLbl.Text = $"tick={tick.GetRawText()}";

        if (root.TryGetProperty("classification", out var cls))
        {
            var presence = cls.TryGetProperty("presence", out var p) && p.ValueKind == JsonValueKind.True;
            var motion = cls.TryGetProperty("motion_level", out var m) ? m.GetString() ?? "—" : "—";
            PresenceLbl.Text = presence ? "PRESENT" : "absent";
            PresenceLbl.TextColor = presence ? Color.FromArgb("#4FD1C5") : Color.FromArgb("#9AA6BF");
            MotionLbl.Text = motion;
        }
        if (root.TryGetProperty("estimated_persons", out var ep) && ep.ValueKind == JsonValueKind.Number)
            PersonsLbl.Text = ep.GetInt32().ToString();

        if (root.TryGetProperty("nodes", out var nodes) && nodes.ValueKind == JsonValueKind.Array)
            NodesLbl.Text = nodes.GetArrayLength().ToString();

        if (root.TryGetProperty("vital_signs", out var vs))
        {
            if (vs.TryGetProperty("heart_rate_bpm", out var hr) && hr.ValueKind == JsonValueKind.Number)
                HrLbl.Text = hr.GetDouble().ToString("0");
            if (vs.TryGetProperty("breathing_rate_bpm", out var br) && br.ValueKind == JsonValueKind.Number)
                BrLbl.Text = br.GetDouble().ToString("0");
        }

        // Fire off a PSO localization probe every ~10 frames.
        if (root.TryGetProperty("tick", out var tk) && tk.ValueKind == JsonValueKind.Number
            && tk.GetInt64() % 10 == 0)
        {
            _ = TryFetchLocalizationAsync();
        }
    }

    private async Task TryFetchLocalizationAsync()
    {
        try
        {
            var wsUrl = HubEntry.Text ?? "";
            if (string.IsNullOrWhiteSpace(wsUrl)) return;
            // ws://host:3001/ws/sensing → http://host:3000/api/v1/localization/person
            var uri = new Uri(wsUrl);
            var httpUri = new UriBuilder("http", uri.Host, 3000, "/api/v1/localization/person").Uri;
            var json = await _http.GetStringAsync(httpUri);
            using var doc = JsonDocument.Parse(json);
            var r = doc.RootElement;
            var status = r.GetProperty("status").GetString();
            if (status == "ok")
            {
                var pos = r.GetProperty("position");
                var x = pos.GetProperty("x").GetDouble();
                var y = pos.GetProperty("y").GetDouble();
                var z = pos.GetProperty("z").GetDouble();
                var res = r.GetProperty("residual_db").GetDouble();
                var conf = r.GetProperty("confidence").GetDouble();
                var used = r.GetProperty("nodes_used").GetInt32();
                var iters = r.GetProperty("iters").GetInt32();
                MainThread.BeginInvokeOnMainThread(() =>
                {
                    LocationLbl.Text = $"x={x:F2} m   y={y:F2} m   z={z:F2} m";
                    LocationMetaLbl.Text = $"residual={res:F2} dB   confidence={conf:P0}   nodes={used}   iters={iters}";
                });
            }
            else
            {
                MainThread.BeginInvokeOnMainThread(() =>
                {
                    LocationLbl.Text = status ?? "—";
                    LocationMetaLbl.Text = "";
                });
            }
        }
        catch (Exception ex)
        {
            MainThread.BeginInvokeOnMainThread(() => Log($"pso: {ex.Message}"));
        }
    }

    private void SetStatus(string text, bool ok)
    {
        StatusBadge.Text = text;
        StatusBadge.TextColor = ok ? Color.FromArgb("#4FD1C5") : Color.FromArgb("#FF5C7A");
    }

    private void Log(string line)
    {
        var ts = DateTime.Now.ToString("HH:mm:ss");
        LogBox.Text = $"[{ts}] {line}\n" + LogBox.Text;
        if (LogBox.Text.Length > 4096) LogBox.Text = LogBox.Text[..4096];
    }
}
