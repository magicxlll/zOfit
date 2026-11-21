using System.Text.Json;
using Xunit;
using ZOfit.Engine;
using ZOfit.Engine.Models;

namespace ZOfit.Engine.Tests;

public class PingTests
{
    [Fact]
    public void Ping_PreservesVietnameseUtf8Echo()
    {
        string input = JsonSerializer.Serialize(new PingRequest { Echo = "Đài cọc PC1 - Bê tông M400 Ø25" });
        string raw = Program.Dispatch("ping", input);
        var resp = JsonSerializer.Deserialize<IpcResponse<PingResponse>>(raw);

        Assert.NotNull(resp);
        Assert.True(resp!.Ok);
        Assert.Equal("ping", resp.Cmd);
        Assert.Contains("sẵn sàng", resp.Message);
        Assert.Equal("Đài cọc PC1 - Bê tông M400 Ø25", resp.Data!.Echo);
    }
}
