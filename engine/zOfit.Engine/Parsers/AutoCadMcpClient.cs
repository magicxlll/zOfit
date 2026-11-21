using System.Net.Sockets;
using System.Text;
using System.Text.Json;
using ZOfit.Engine.Models;

namespace ZOfit.Engine.Parsers;

/// <summary>
/// Direct TCP client to communicate with AutoCAD MCP Bridge on port 9877 (127.0.0.1)
/// </summary>
public sealed class AutoCadMcpClient
{
    private readonly string _host;
    private readonly int _port;
    private readonly int _timeoutMs;

    public AutoCadMcpClient(string host = "127.0.0.1", int port = 9877, int timeoutMs = 8000)
    {
        _host = host;
        _port = port;
        _timeoutMs = timeoutMs;
    }

    public async Task<string> SendRpcAsync(string method, object? parameters = null)
    {
        using var client = new TcpClient();
        using var cts = new CancellationTokenSource(_timeoutMs);

        await client.ConnectAsync(_host, _port, cts.Token);
        using var stream = client.GetStream();
        stream.ReadTimeout = _timeoutMs;
        stream.WriteTimeout = _timeoutMs;

        var payload = new
        {
            id = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds(),
            method = method,
            @params = parameters ?? new { }
        };

        string jsonRequest = JsonSerializer.Serialize(payload) + "\n";
        byte[] reqBytes = Encoding.UTF8.GetBytes(jsonRequest);
        await stream.WriteAsync(reqBytes, cts.Token);

        var sb = new StringBuilder();
        byte[] buf = new byte[8192];
        while (!cts.IsCancellationRequested)
        {
            int read = await stream.ReadAsync(buf, cts.Token);
            if (read == 0) break;
            string chunk = Encoding.UTF8.GetString(buf, 0, read);
            sb.Append(chunk);
            if (chunk.Contains('\n')) break;
        }

        string rawResponse = sb.ToString().Trim();
        using var doc = JsonDocument.Parse(rawResponse);
        if (doc.RootElement.TryGetProperty("error", out var errProp) && errProp.ValueKind == JsonValueKind.String)
        {
            throw new InvalidOperationException($"AutoCAD MCP Error: {errProp.GetString()}");
        }

        if (doc.RootElement.TryGetProperty("result", out var resProp))
        {
            return resProp.GetRawText();
        }

        return rawResponse;
    }

    public async Task<List<CadEntityItemDto>> GetSelectionAsync()
    {
        string raw = await SendRpcAsync("get_selection");
        return ParseEntityItems(raw);
    }

    public async Task<List<CadEntityItemDto>> QueryEntitiesAsync(double[]? windowMin = null, double[]? windowMax = null, string? textContains = null, int maxCount = 300)
    {
        var p = new Dictionary<string, object>();
        if (windowMin != null) p["windowMin"] = windowMin;
        if (windowMax != null) p["windowMax"] = windowMax;
        if (!string.IsNullOrEmpty(textContains)) p["textContains"] = textContains;
        p["maxCount"] = maxCount;

        string raw = await SendRpcAsync("query_entities", p);
        return ParseEntityItems(raw);
    }

    private static List<CadEntityItemDto> ParseEntityItems(string rawResultJson)
    {
        var items = new List<CadEntityItemDto>();
        try
        {
            using var doc = JsonDocument.Parse(rawResultJson);
            JsonElement root = doc.RootElement;
            JsonElement arr = root;

            if (root.ValueKind == JsonValueKind.Object)
            {
                if (root.TryGetProperty("items", out var itProp) && itProp.ValueKind == JsonValueKind.Array)
                {
                    arr = itProp;
                }
                else if (root.TryGetProperty("selection", out var selProp) && selProp.ValueKind == JsonValueKind.Array)
                {
                    arr = selProp;
                }
            }

            if (arr.ValueKind == JsonValueKind.Array)
            {
                foreach (var el in arr.EnumerateArray())
                {
                    var dto = new CadEntityItemDto();
                    if (el.TryGetProperty("handle", out var h)) dto.Handle = h.GetString() ?? "";
                    if (el.TryGetProperty("type", out var tp)) dto.Type = tp.GetString() ?? "";
                    if (el.TryGetProperty("layer", out var l)) dto.Layer = l.GetString() ?? "";
                    if (el.TryGetProperty("text", out var tx)) dto.Text = tx.GetString() ?? "";
                    if (el.TryGetProperty("blockName", out var bn)) dto.BlockName = bn.GetString() ?? "";

                    if (el.TryGetProperty("position", out var pos) && pos.ValueKind == JsonValueKind.Array)
                    {
                        var coords = new List<double>();
                        foreach (var c in pos.EnumerateArray())
                        {
                            if (c.TryGetDouble(out double d)) coords.Add(d);
                        }
                        dto.Position = coords.ToArray();
                    }

                    if (el.TryGetProperty("bounds", out var bnd) && bnd.ValueKind == JsonValueKind.Object)
                    {
                        var b = new CadBoundsDto();
                        if (bnd.TryGetProperty("width", out var w)) b.Width = w.GetDouble();
                        if (bnd.TryGetProperty("height", out var ht)) b.Height = ht.GetDouble();
                        dto.Bounds = b;
                    }

                    if (el.TryGetProperty("attributes", out var attr) && attr.ValueKind == JsonValueKind.Object)
                    {
                        dto.Attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        foreach (var prop in attr.EnumerateObject())
                        {
                            dto.Attributes[prop.Name] = prop.Value.GetString() ?? "";
                        }
                    }

                    items.Add(dto);
                }
            }
        }
        catch
        {
            // Fallback empty if cannot parse
        }

        return items;
    }
}
