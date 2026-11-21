using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Windows.Forms;
using ZOfit.Engine.Generators;
using ZOfit.Engine.Models;
using ZOfit.Engine.Parsers;
using ZOfit.Engine.Sync;
using ZOfit.Engine.UI;

namespace ZOfit.Engine;

public static class Program
{
    private sealed class HwndOwner : IWin32Window
    {
        public IntPtr Handle { get; }
        public HwndOwner(long hwnd) => Handle = new IntPtr(hwnd);
    }

    [DllImport("kernel32.dll")] private static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNameCaseInsensitive = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = false,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    [STAThread]
    public static int Main(string[] args)
    {
        CultureInfo.DefaultThreadCurrentCulture = CultureInfo.InvariantCulture;
        CultureInfo.DefaultThreadCurrentUICulture = CultureInfo.InvariantCulture;
        try { Console.OutputEncoding = Encoding.UTF8; } catch { /* WinExe may have no console handle */ }

        string? cmd = null;
        string? inPath = null;
        string? outPath = null;

        for (int i = 0; i < args.Length; i++)
        {
            switch (args[i].ToLowerInvariant())
            {
                case "--cmd" when i + 1 < args.Length:
                    cmd = args[++i];
                    break;
                case "--in" when i + 1 < args.Length:
                    inPath = args[++i];
                    break;
                case "--out" when i + 1 < args.Length:
                    outPath = args[++i];
                    break;
            }
        }

        if (string.IsNullOrWhiteSpace(cmd))
        {
            Console.Error.WriteLine("Usage: zofit-engine --cmd <op> --in <req.json> --out <res.json>");
            return 2;
        }

        try
        {
            string inputJson = !string.IsNullOrWhiteSpace(inPath) && File.Exists(inPath)
                ? File.ReadAllText(inPath, Encoding.UTF8)
                : "{}";

            string outputJson = Dispatch(cmd.Trim().ToLowerInvariant(), inputJson);

            if (!string.IsNullOrWhiteSpace(outPath))
            {
                string? dir = Path.GetDirectoryName(Path.GetFullPath(outPath));
                if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
                File.WriteAllText(outPath, outputJson, new UTF8Encoding(false));
            }
            else
            {
                Console.WriteLine(outputJson);
            }
            return 0;
        }
        catch (Exception ex)
        {
            var errResp = new IpcResponse<object>
            {
                Ok = false,
                Cmd = cmd,
                Message = ex.Message
            };
            string errJson = JsonSerializer.Serialize(errResp, JsonOpts);
            if (!string.IsNullOrWhiteSpace(outPath))
            {
                File.WriteAllText(outPath, errJson, new UTF8Encoding(false));
            }
            else
            {
                Console.Error.WriteLine(errJson);
            }
            return 1;
        }
    }

    public static string Dispatch(string cmd, string inputJson)
    {
        switch (cmd)
        {
            case "ping":
            {
                var req = JsonSerializer.Deserialize<PingRequest>(inputJson, JsonOpts) ?? new PingRequest();
                return Ok("ping", "zOfit Engine sẵn sàng / Ready", new PingResponse
                {
                    Version = "2.1.0",
                    Echo = req.Echo,
                    Timestamp = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture)
                });
            }

            case "expand_rebar_presets":
            {
                var req = JsonSerializer.Deserialize<ExpandRebarPresetsRequest>(inputJson, JsonOpts)
                          ?? throw new InvalidOperationException("Invalid ExpandRebarPresetsRequest JSON.");
                var gen = new RebarBbsGenerator();
                var data = gen.ExpandPresetsToConfigs(req);
                return Ok("expand_rebar_presets", $"Expanded {data.AddedRowsCount} smart rebar config rows.", data);
            }

            case "open_rebar_editor":
            {
                var req = JsonSerializer.Deserialize<EditRebarUiRequest>(inputJson, JsonOpts)
                          ?? throw new InvalidOperationException("Invalid EditRebarUiRequest JSON.");
                var hCon = GetConsoleWindow();
                if (hCon != IntPtr.Zero) ShowWindow(hCon, 0);
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                using var dlg = new RebarEditorForm(req);
                if (req.ParentHwnd != 0)
                    dlg.ShowDialog(new HwndOwner(req.ParentHwnd));
                else
                    dlg.ShowDialog();
                return Ok("open_rebar_editor", dlg.Result.Accepted ? "Updated rebar configuration." : "Cancelled.", dlg.Result);
            }

            case "generate_structure":
            {
                var req = JsonSerializer.Deserialize<GenerateStructureRequest>(inputJson, JsonOpts)
                          ?? throw new InvalidOperationException("Invalid GenerateStructureRequest JSON.");
                var engine = new SyncEngine();
                var data = engine.GenerateAndSync(req);
                return Ok("generate_structure", $"Generated {data.DetailSheets.Count} Detail & {data.TktSheets.Count} TKT sheets.", data);
            }

            case "parse_cad":
            {
                var req = JsonSerializer.Deserialize<CadParseRequest>(inputJson, JsonOpts)
                          ?? throw new InvalidOperationException("Invalid CadParseRequest JSON.");
                var parser = new CadTextParser();
                var data = parser.Parse(req);
                return Ok("parse_cad", $"Parsed {data.ParsedCount}/{data.TotalInput} CAD items ({data.UnknownCount} manual).", data);
            }

            case "generate_finishing":
            {
                var req = JsonSerializer.Deserialize<GenerateFinishingRequest>(inputJson, JsonOpts)
                          ?? throw new InvalidOperationException("Invalid GenerateFinishingRequest JSON.");
                var engine = new FinishingEngine();
                var sheet = engine.Generate(req);
                return Ok("generate_finishing", $"Generated Finishing_Detail ({sheet.Rows.Count} rows).", new GenerateFinishingResponse { Sheet = sheet });
            }

            case "generate_infra":
            {
                var req = JsonSerializer.Deserialize<GenerateInfraRequest>(inputJson, JsonOpts)
                          ?? throw new InvalidOperationException("Invalid GenerateInfraRequest JSON.");
                var engine = new InfraEngine();
                var sheet = engine.Generate(req);
                return Ok("generate_infra", $"Generated Civil_Detail ({sheet.Rows.Count} rows).", new GenerateInfraResponse { Sheet = sheet });
            }

            case "parse_beam_section":
            {
                var req = JsonSerializer.Deserialize<CadBeamSectionParseRequest>(inputJson, JsonOpts)
                          ?? throw new InvalidOperationException("Invalid CadBeamSectionParseRequest JSON.");
                var parser = new CadBeamSectionParser();
                var data = parser.Parse(req);
                return Ok("parse_beam_section", $"Parsed {data.BeamCount} beams from CAD section table.", data);
            }

            case "scan_cad_beam":
            {
                var req = JsonSerializer.Deserialize<ScanCadBeamRequest>(inputJson, JsonOpts)
                          ?? throw new InvalidOperationException("Invalid ScanCadBeamRequest JSON.");

                var mcpClient = new AutoCadMcpClient(req.Host, req.Port);
                List<CadEntityItemDto> entities;
                try
                {
                    entities = mcpClient.GetSelectionAsync().GetAwaiter().GetResult();
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException($"Không thể kết nối AutoCAD tại {req.Host}:{req.Port}. Hãy mở AutoCAD và kiểm tra cổng 9877: {ex.Message}");
                }

                if (entities == null || entities.Count == 0)
                {
                    return Ok("scan_cad_beam", "Chưa có đối tượng nào được chọn trên AutoCAD.", new ScanCadBeamResponse
                    {
                        Accepted = false,
                        BeamCount = 0,
                        Message = "Chưa có đối tượng nào được chọn trên AutoCAD!\n\nVui lòng làm theo các bước:\n1. Mở màn hình AutoCAD đang mở bản vẽ.\n2. Dùng chuột quét chọn (bôi đen) trọn vẹn bảng mặt cắt dầm cần bóc tách.\n3. Bấm lại nút 'Quét Dầm CAD (MCP)' trên Excel để nạp số liệu."
                    });
                }

                var parser = new CadBeamSectionParser();
                var parsed = parser.Parse(new CadBeamSectionParseRequest
                {
                    Building = req.Building,
                    Floor = req.Floor,
                    DefaultLengthM = req.DefaultLengthM,
                    Entities = entities
                });

                if (parsed.BeamCount == 0)
                {
                    return Ok("scan_cad_beam", "Không nhận diện được cấu trúc bảng dầm.", new ScanCadBeamResponse
                    {
                        Accepted = false,
                        BeamCount = 0,
                        Message = "Không tìm thấy cấu trúc bảng dầm trong vùng chọn CAD."
                    });
                }

                bool accepted = true;
                if (req.AutoOpenPreview && parsed.Beams.Count > 0)
                {
                    var first = parsed.Beams[0];
                    var editReq = new EditRebarUiRequest
                    {
                        ParentHwnd = req.ParentHwnd,
                        Spec = req.Spec,
                        Element = first.Element,
                        ExistingRows = first.RebarRows
                    };

                    var hCon = GetConsoleWindow();
                    if (hCon != IntPtr.Zero) ShowWindow(hCon, 0);
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                    using var dlg = new RebarEditorForm(editReq);
                    dlg.Text = $"zOfit QS Pro - Quét CAD Dầm Thành Công: {first.Element.Mark} ({first.Element.WidthM * 1000:0}x{first.Element.HeightM * 1000:0})";
                    if (req.ParentHwnd != 0)
                        dlg.ShowDialog(new HwndOwner(req.ParentHwnd));
                    else
                        dlg.ShowDialog();

                    accepted = dlg.Result.Accepted;
                    if (accepted)
                    {
                        first.RebarRows = dlg.Result.Rows;
                    }
                }

                return Ok("scan_cad_beam", accepted ? $"Đã nhận diện {parsed.BeamCount} dầm và xác nhận ghi vào Excel." : "Người dùng đã hủy.", new ScanCadBeamResponse
                {
                    Accepted = accepted,
                    BeamCount = parsed.BeamCount,
                    Beams = parsed.Beams,
                    Message = accepted ? $"Đã quét thành công {parsed.BeamCount} dầm từ AutoCAD." : "Đã hủy."
                });
            }

            default:
                throw new InvalidOperationException($"Unknown command: {cmd}");
        }
    }

    private static string Ok<T>(string cmd, string msg, T data)
    {
        return JsonSerializer.Serialize(new IpcResponse<T>
        {
            Ok = true,
            Cmd = cmd,
            Message = msg,
            Data = data
        }, JsonOpts);
    }
}
