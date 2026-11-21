namespace ZOfit.Engine.Generators;

public sealed record BarShapeMeta(string Code, int BendCount, bool HasLap, string LenExpr);

public static class BarShapeCatalog
{
    private static readonly Dictionary<string, BarShapeMeta> Map = new(StringComparer.OrdinalIgnoreCase)
    {
        ["D01"] = new("D01", 0, true, "F{r}"),
        ["D02"] = new("D02", 2, true, "F{r}+M{r}+N{r}"),
        ["D03"] = new("D03", 2, true, "F{r}+G{r}+M{r}"),
        ["D04"] = new("D04", 2, true, "F{r}+M{r}+N{r}"),
        ["D05"] = new("D05", 2, true, "F{r}+G{r}+M{r}"),
        ["D06"] = new("D06", 1, true, "F{r}+G{r}"),
        ["D07"] = new("D07", 2, true, "F{r}+G{r}+H{r}"),
        ["D08"] = new("D08", 4, true, "F{r}+G{r}+H{r}+I{r}+J{r}"),
        ["D09"] = new("D09", 2, true, "F{r}+G{r}+H{r}"),
        ["D10"] = new("D10", 2, true, "F{r}+G{r}+H{r}"),
        ["D11"] = new("D11", 1, true, "F{r}+G{r}"),
        ["D12"] = new("D12", 2, true, "F{r}+G{r}+M{r}"),
        ["D13"] = new("D13", 3, true, "F{r}+G{r}+M{r}+N{r}"),
        ["D14"] = new("D14", 2, true, "F{r}+G{r}+H{r}"),
        ["D15"] = new("D15", 2, true, "F{r}+G{r}+H{r}"),
        ["D16"] = new("D16", 2, true, "F{r}+G{r}+H{r}"),
        ["D17"] = new("D17", 3, true, "F{r}+G{r}+H{r}+I{r}"),
        ["D18"] = new("D18", 4, true, "F{r}+G{r}+H{r}+I{r}+J{r}"),
        ["D19"] = new("D19", 5, true, "F{r}+G{r}+H{r}+I{r}+J{r}+K{r}"),
        ["D20"] = new("D20", 3, true, "F{r}+G{r}+H{r}+I{r}"),
        ["D21"] = new("D21", 3, true, "F{r}+G{r}+H{r}+I{r}"),
        ["D22"] = new("D22", 5, true, "F{r}+G{r}+H{r}+I{r}+J{r}+K{r}"),
        ["D23"] = new("D23", 6, false, "F{r}+G{r}+H{r}+I{r}+J{r}+K{r}"),
        ["D24"] = new("D24", 2, false, "F{r}+M{r}+N{r}"),
        ["D25"] = new("D25", 2, true, "F{r}+G{r}+M{r}"),
        ["D26"] = new("D26", 2, true, "F{r}+M{r}+N{r}"),
        ["D27"] = new("D27", 2, true, "F{r}+G{r}+M{r}"),
        ["D28"] = new("D28", 5, true, "F{r}+G{r}+H{r}+I{r}+J{r}+K{r}"),
        ["D29"] = new("D29", 5, true, "F{r}+G{r}+H{r}+I{r}+J{r}+K{r}"),
        ["D30"] = new("D30", 3, true, "F{r}+G{r}+H{r}+I{r}"),
        ["D31"] = new("D31", 6, true, "F{r}+G{r}+H{r}+I{r}+J{r}+K{r}+L{r}"),
        ["D32"] = new("D32", 6, true, "F{r}+G{r}+H{r}+I{r}+J{r}+K{r}+L{r}"),
        ["D33"] = new("D33", 4, true, "F{r}+G{r}+H{r}+I{r}+J{r}"),
        ["D34"] = new("D34", 2, false, "F{r}+G{r}+H{r}"),
        ["D35"] = new("D35", 3, false, "F{r}+G{r}+H{r}+I{r}"),
        ["D36"] = new("D36", 4, false, "F{r}+G{r}+H{r}+I{r}+J{r}"),
        ["D37"] = new("D37", 4, true, "F{r}+G{r}+H{r}+I{r}+J{r}"),
        ["D38"] = new("D38", 4, true, "F{r}+G{r}+H{r}+I{r}+J{r}"),
        ["D39"] = new("D39", 4, false, "2*(F{r}+G{r})+2*M{r}"),
        ["D40"] = new("D40", 6, true, "F{r}+G{r}+H{r}+I{r}+J{r}+K{r}+L{r}"),
        ["D41"] = new("D41", 0, false, "2*PI()*F{r}+M{r}")
    };

    public static BarShapeMeta Get(string? code)
    {
        string k = (code ?? "D01").Trim().ToUpperInvariant();
        return Map.TryGetValue(k, out var meta) ? meta : Map["D01"];
    }
}
