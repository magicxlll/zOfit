using ZOfit.Engine.Generators;
using ZOfit.Engine.Models;

namespace ZOfit.Engine.Sync;

public sealed class SyncEngine
{
    private readonly RcDetailGenerator _detailGen = new();
    private readonly RebarBbsGenerator _tktGen = new();

    public GenerateStructureResponse GenerateAndSync(GenerateStructureRequest req)
    {
        var buildings = req.Elements
            .Select(e => e.Building?.Trim() ?? string.Empty)
            .Where(b => !string.IsNullOrEmpty(b))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (!string.IsNullOrWhiteSpace(req.TargetBuilding))
        {
            buildings = buildings
                .Where(b => string.Equals(b, req.TargetBuilding.Trim(), StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        var response = new GenerateStructureResponse();

        foreach (var bld in buildings)
        {
            var bElements = req.Elements
                .Where(e => string.Equals(e.Building?.Trim(), bld, StringComparison.OrdinalIgnoreCase))
                .ToList();

            var detailSheet = SyncDetailForBuilding(bld, req.Spec, bElements, req.ExistingDetailRows);
            response.DetailSheets.Add(detailSheet);

            var cfgIds = new HashSet<string>(
                (req.RebarConfigs ?? new List<RebarConfigRowDto>())
                    .Where(c => !string.Equals(c.RowKind, "BLK_HDR", StringComparison.OrdinalIgnoreCase)
                             && !string.Equals(c.RowKind, "BLK_END", StringComparison.OrdinalIgnoreCase))
                    .Select(c => c.ElementId?.Trim() ?? ""),
                StringComparer.OrdinalIgnoreCase);

            var tktElements = bElements
                .Where(e => (!string.IsNullOrWhiteSpace(e.RebarTemplate)
                             && !string.Equals(e.RebarTemplate.Trim(), "NONE", StringComparison.OrdinalIgnoreCase))
                            || cfgIds.Contains(e.ElementId))
                .ToList();

            if (tktElements.Count > 0 || req.ExistingTktRows.Any(r => string.Equals(r.TagId, "MANUAL", StringComparison.OrdinalIgnoreCase)))
            {
                var tktSheet = SyncTktForBuilding(bld, req.Spec, bElements, req.ExistingTktRows, req.RebarConfigs);
                response.TktSheets.Add(tktSheet);
            }
        }

        return response;
    }

    private GeneratedSheetDto SyncDetailForBuilding(
        string building,
        ProjectSpecDto spec,
        List<ElementDataDto> elements,
        List<ExistingDetailRowDto> existingRows)
    {
        var dto = _detailGen.Generate(building, spec, elements, existingRows);

        var oldElementIds = ExtractElementIds(existingRows.Select(r => r.TagId));
        var newElementIds = new HashSet<string>(elements.Select(e => e.ElementId), StringComparer.OrdinalIgnoreCase);

        int added = newElementIds.Count(id => !oldElementIds.Contains(id));
        int updated = newElementIds.Count(id => oldElementIds.Contains(id));

        var removedRows = existingRows
            .Where(r => IsElementRowTag(r.TagId) && !IsTagRemoved(r.TagId))
            .Where(r =>
            {
                string eid = ExtractElementIdFromTag(r.TagId);
                return !string.IsNullOrEmpty(eid) && !newElementIds.Contains(eid);
            })
            .ToList();

        if (removedRows.Count > 0)
        {
            dto.Rows.Add(RcDetailGenerator.MakeRow("BLANK", 0, "", "REMOVED"));
            dto.Rows.Add(RcDetailGenerator.MakeRow("WORK_HDR", 0, "SEC_REMOVED", "REMOVED", "!", "[REMOVED ITEMS] CÁC MỤC ĐÃ XÓA KHỎI QS_DATA (KIỂM TRA ĐỐI CHIẾU)", "", "", null, null, null, null, null, null, null, 0, "SEC_REMOVED", "REMOVED"));
            foreach (var rem in removedRows)
            {
                string desc = rem.Description.StartsWith("[REMOVED]", StringComparison.OrdinalIgnoreCase)
                    ? rem.Description
                    : "[REMOVED] " + rem.Description;
                dto.Rows.Add(RcDetailGenerator.MakeRow("REMOVED", 1, "REMOVED::" + rem.TagId, rem.SectionKey,
                    rem.Code, desc, rem.SpecText, rem.Unit, rem.NoExpr, rem.LExpr, rem.WExpr, rem.HExpr,
                    null, null, null, 0, "REMOVED::" + rem.TagId, rem.SectionKey));
            }
        }

        dto.AddedCount = added;
        dto.UpdatedCount = updated;
        dto.RemovedFlaggedCount = removedRows.Count;
        return dto;
    }

    private GeneratedSheetDto SyncTktForBuilding(
        string building,
        ProjectSpecDto spec,
        List<ElementDataDto> elements,
        List<ExistingTktRowDto> existingRows,
        IReadOnlyList<RebarConfigRowDto>? rebarConfigs)
    {
        var dto = _tktGen.Generate(building, spec, elements, existingRows, rebarConfigs);

        var oldIds = ExtractElementIds(existingRows.Select(r => r.TagId));
        var newIds = new HashSet<string>(elements.Select(e => e.ElementId), StringComparer.OrdinalIgnoreCase);

        int added = newIds.Count(id => !oldIds.Contains(id));
        int updated = newIds.Count(id => oldIds.Contains(id));

        var removedRows = existingRows
            .Where(r => IsElementRowTag(r.TagId) && !IsTagRemoved(r.TagId))
            .Where(r =>
            {
                string eid = ExtractElementIdFromTag(r.TagId);
                return !string.IsNullOrEmpty(eid) && !newIds.Contains(eid);
            })
            .ToList();

        int stt = dto.Rows.Count + 1;
        foreach (var rem in removedRows)
        {
            var vals = new object?[RebarBbsGenerator.TotalCols];
            vals[0] = stt++;
            vals[1] = rem.CompName.StartsWith("[REMOVED]") ? rem.CompName : "[REMOVED] " + rem.CompName;
            vals[2] = rem.BarMark;
            vals[3] = string.IsNullOrWhiteSpace(rem.ShapeCode) ? "D01" : rem.ShapeCode;
            vals[4] = vals[3];
            vals[14] = rem.Dia;
            vals[15] = 0;
            vals[16] = 0;
            vals[17] = 0;
            vals[18] = 0;
            vals[19] = 0;
            vals[20] = 0;
            vals[21] = 0;
            vals[22] = 0;
            vals[23] = "REMOVED from QS_Data";
            vals[24] = "REMOVED::" + rem.TagId;
            vals[25] = rem.BoqBucket;

            dto.Rows.Add(new SheetCellRowDto
            {
                RowKind = "REMOVED",
                OutlineLevel = 0,
                TagId = "REMOVED::" + rem.TagId,
                SectionKey = rem.BoqBucket,
                Values = vals.ToList()
            });
        }

        dto.AddedCount = added;
        dto.UpdatedCount = updated;
        dto.RemovedFlaggedCount = removedRows.Count;
        return dto;
    }

    private static bool IsTagRemoved(string? tagId) =>
        !string.IsNullOrEmpty(tagId) && tagId.StartsWith("REMOVED::", StringComparison.OrdinalIgnoreCase);

    private static bool IsElementRowTag(string? tagId)
    {
        if (string.IsNullOrWhiteSpace(tagId)) return false;
        if (string.Equals(tagId, "MANUAL", StringComparison.OrdinalIgnoreCase)) return false;
        return tagId.Contains("::", StringComparison.Ordinal);
    }

    private static string ExtractElementIdFromTag(string? tagId)
    {
        if (string.IsNullOrWhiteSpace(tagId)) return string.Empty;
        int idx = tagId.IndexOf("::", StringComparison.Ordinal);
        return idx > 0 ? tagId.Substring(0, idx) : string.Empty;
    }

    private static HashSet<string> ExtractElementIds(IEnumerable<string?> tags)
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var t in tags)
        {
            if (IsElementRowTag(t) && !IsTagRemoved(t))
            {
                string eid = ExtractElementIdFromTag(t);
                if (!string.IsNullOrEmpty(eid)) set.Add(eid);
            }
        }
        return set;
    }
}
