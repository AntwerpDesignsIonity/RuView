using System.Net.WebSockets;
using System.Text;
using System.Text.Json;
using RuView.Sdk.Models;

namespace RuView.Sdk.WebSockets;

/// <summary>
/// Manages the persistent WebSocket connection to the sensing server.
/// Deserializes SensingUpdate frames and raises events.
/// </summary>
public sealed class SensingWebSocketService : IAsyncDisposable
{
    private ClientWebSocket? _ws;
    private CancellationTokenSource? _cts;
    private readonly JsonSerializerOptions _json = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower,
        PropertyNameCaseInsensitive = true
    };

    public event Action<SensingUpdate>? OnUpdate;
    public event Action<string>? OnError;
    public event Action<bool>? OnConnectionChanged;

    public bool IsConnected => _ws?.State == WebSocketState.Open;

    public async Task ConnectAsync(string host, int port, CancellationToken ct = default)
    {
        await DisconnectAsync();

        _cts = CancellationTokenSource.CreateLinkedTokenSource(ct);
        _ws = new ClientWebSocket();

        var uri = new Uri($"ws://{host}:{port}/ws/sensing");

        try
        {
            await _ws.ConnectAsync(uri, _cts.Token);
            OnConnectionChanged?.Invoke(true);
            _ = Task.Run(() => ReceiveLoop(_cts.Token), _cts.Token);
        }
        catch (Exception ex)
        {
            OnError?.Invoke($"Connect failed: {ex.Message}");
            OnConnectionChanged?.Invoke(false);
        }
    }

    private async Task ReceiveLoop(CancellationToken ct)
    {
        var buffer = new byte[64 * 1024];

        try
        {
            while (!ct.IsCancellationRequested && _ws?.State == WebSocketState.Open)
            {
                using var ms = new MemoryStream();
                WebSocketReceiveResult result;

                do
                {
                    result = await _ws.ReceiveAsync(buffer, ct);
                    ms.Write(buffer, 0, result.Count);
                }
                while (!result.EndOfMessage);

                if (result.MessageType == WebSocketMessageType.Close)
                {
                    OnConnectionChanged?.Invoke(false);
                    return;
                }

                if (result.MessageType == WebSocketMessageType.Text)
                {
                    var json = Encoding.UTF8.GetString(ms.ToArray());
                    var update = JsonSerializer.Deserialize<SensingUpdate>(json, _json);
                    if (update is not null)
                        OnUpdate?.Invoke(update);
                }
            }
        }
        catch (OperationCanceledException) { }
        catch (Exception ex)
        {
            OnError?.Invoke($"WebSocket error: {ex.Message}");
        }
        finally
        {
            OnConnectionChanged?.Invoke(false);
        }
    }

    public async Task DisconnectAsync()
    {
        if (_cts is not null)
        {
            await _cts.CancelAsync();
            _cts.Dispose();
            _cts = null;
        }

        if (_ws is not null)
        {
            if (_ws.State == WebSocketState.Open)
            {
                try
                {
                    await _ws.CloseAsync(WebSocketCloseStatus.NormalClosure, "bye", CancellationToken.None);
                }
                catch { }
            }
            _ws.Dispose();
            _ws = null;
        }
    }

    public async ValueTask DisposeAsync() => await DisconnectAsync();
}