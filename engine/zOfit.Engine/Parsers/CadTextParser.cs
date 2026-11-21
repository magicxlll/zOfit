using System.Globalization;
using System.Text.RegularExpressions;
using ZOfit.Engine.Models;

namespace ZOfit.Engine.Parsers;

/// <summary>
/// Parses raw AutoCAD Text dumps and Block Attribute dumps from sheet 'Data'
/// e.g.:
///   "A-W300- i0.2%-L=13m"
///   "B-W400- L=6m"
///   "Box:600x600-L11m-i0.2%"
///   "D200-L44m- i0.5%"
///   "TN-TENGA2 | 19.50 | 18.10 | 1"
/// </summary>
public sealed partial class CadTextParser
{
    [GeneratedRegex(@"^(?<prefix>[A-Za-z0-9]+)-W(?<w>\d+)(?:\s*-\s*i(?<slope>[\d.]+)%)?\s*-\s*L\s*=?\s*(?<len>[\d.]+)\s*m?", RegexOptions.IgnoreCase)]
    private static partial Regex TrenchRegex();

    [GeneratedRegex(@"^Box[:\s]*(?<w>\d+)\s*[xX*]\s*(?<h>\d+)(?:\s*-\s*L\s*=?\s*(?<len>[\d.]+)\s*m?)?(?:\s*-\s*i(?<slope>[\d.]+)%)?(?:\s*-\s*L\s*=?\s*(?<len2>[\d.]+)\s*m?)?", RegexOptions.IgnoreCase)]
    private static partial Regex BoxCulvertRegex();

    [GeneratedRegex(@"^(?:[A-Za-z]+\s+)?(?<mat>HDPE|PVC|uPVC|BTCT|D)?\s*D(?<dia>\d+)(?:\s*[- ]\s*L\s*=?\s*(?<len>[\d.]+)\s*m?)(?:\s*[- ]\s*i(?<slope>[\d.]+)%)?", RegexOptions.IgnoreCase)]
    private static partial Regex PipeRegex();

    public CadParseResponse Parse(CadParseRequest req)
    {
        var items = new List<CadParsedItemDto>();
        var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var raw in req.RawRows)
        {
            string colA = (raw.ColA ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(colA)) continue;

            // Skip header / total rows
            string up = colA.ToUpperInvariant();
            if (up.StartsWith("BANG THONG KE") || up == "STT" || up == "GIA TRI" || up == "TEN BLOCK" || up == "TONG CONG")
                continue;

            var parsed = ParseSingleRow(raw.SourceRow, colA, raw.ColB, raw.ColC, raw.ColD);
            items.Add(parsed);
            if (!parsed.NeedsManualCheck && !string.IsNullOrWhiteSpace(parsed.CategoryKey))
            {
                keys.Add(parsed.CategoryKey);
            }
        }

        int unknownCount = items.Count(i => i.NeedsManualCheck);
        return new CadParseResponse
        {
            TotalInput = items.Count,
            ParsedCount = items.Count - unknownCount,
            UnknownCount = unknownCount,
            Items = items,
            DistinctKeys = keys.OrderBy(k => k, StringComparer.OrdinalIgnoreCase).ToList()
        };
    }

    public CadParsedItemDto ParseSingleRow(int sourceRow, string colA, string? colB, string? colC, string? colD)
    {
        // Normalize multiple spaces around hyphens
        string normalized = Regex.Replace(colA.Trim(), @"\s+", " ");

        // 1. Open Trench / Gutter: "A-W300- i0.2%-L=13m", "B-W400- L=6m"
        var mTrench = TrenchRegex().Match(normalized);
        if (mTrench.Success)
        {
            double w = ParseDouble(mTrench.Groups["w"].Value);
            double slope = ParseDouble(mTrench.Groups["slope"].Value);
            double len = ParseDouble(mTrench.Groups["len"].Value);
            string key = $"{mTrench.Groups["prefix"].Value.ToUpperInvariant()}-W{(int)w}";
            return new CadParsedItemDto
            {
                SourceRow = sourceRow,
                RawText = colA,
                ItemType = "TRENCH",
                CategoryKey = key,
                WidthMm = w,
                SlopePct = slope,
                LengthM = len,
                Qty = ParseQtyOrDefault(colD, 1),
                NeedsManualCheck = len <= 0
            };
        }

        // 2. Box Culvert: "Box:600x600-L11m-i0.2%"
        var mBox = BoxCulvertRegex().Match(normalized);
        if (mBox.Success)
        {
            double w = ParseDouble(mBox.Groups["w"].Value);
            double h = ParseDouble(mBox.Groups["h"].Value);
            string lenStr = mBox.Groups["len"].Success ? mBox.Groups["len"].Value : mBox.Groups["len2"].Value;
            double len = ParseDouble(lenStr);
            double slope = ParseDouble(mBox.Groups["slope"].Value);
            return new CadParsedItemDto
            {
                SourceRow = sourceRow,
                RawText = colA,
                ItemType = "BOX_CULVERT",
                CategoryKey = $"Box:{(int)w}x{(int)h}",
                WidthMm = w,
                HeightMm = h,
                SlopePct = slope,
                LengthM = len,
                Qty = ParseQtyOrDefault(colD, 1),
                NeedsManualCheck = len <= 0
            };
        }

        // 3. Round Pipe: "D200-L44m- i0.5%", "D200-L4m", "HDPE D300 L=25.5m"
        var mPipe = PipeRegex().Match(normalized);
        if (mPipe.Success)
        {
            double dia = ParseDouble(mPipe.Groups["dia"].Value);
            double len = ParseDouble(mPipe.Groups["len"].Value);
            double slope = ParseDouble(mPipe.Groups["slope"].Value);
            string mat = mPipe.Groups["mat"].Value.Trim().ToUpperInvariant();
            string key = string.IsNullOrEmpty(mat) || mat == "D" ? $"D{(int)dia}" : $"{mat}-D{(int)dia}";
            return new CadParsedItemDto
            {
                SourceRow = sourceRow,
                RawText = colA,
                ItemType = "PIPE",
                CategoryKey = key,
                DiameterMm = dia,
                SlopePct = slope,
                LengthM = len,
                Qty = ParseQtyOrDefault(colD, 1),
                NeedsManualCheck = len <= 0
            };
        }

        // 4. Manhole Block Attribute: ColA="TN-TENGA2" & ColB=MG & ColC=DC (or MG/DC present)
        if (!string.IsNullOrWhiteSpace(colB) && !string.IsNullOrWhiteSpace(colC) &&
            double.TryParse(colB.Trim().Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double mg) &&
            double.TryParse(colC.Trim().Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double dc))
        {
            double depth = Math.Max(0, Math.Round(mg - dc, 3));
            return new CadParsedItemDto
            {
                SourceRow = sourceRow,
                RawText = colA,
                ItemType = "MANHOLE",
                CategoryKey = colA.Trim(),
                MgElev = mg,
                DcElev = dc,
                DepthM = depth,
                Qty = ParseQtyOrDefault(colD, 1),
                NeedsManualCheck = false
            };
        }

        return new CadParsedItemDto
        {
            SourceRow = sourceRow,
            RawText = colA,
            ItemType = "UNKNOWN",
            CategoryKey = "UNKNOWN",
            NeedsManualCheck = true
        };
    }

    private static double ParseDouble(string? s)
    {
        if (string.IsNullOrWhiteSpace(s)) return 0;
        return double.TryParse(s.Trim().Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double val) ? val : 0;
    }

    private static double ParseQtyOrDefault(string? s, double defVal)
    {
        double v = ParseDouble(s);
        return v > 0 ? v : defVal;
    }
}
